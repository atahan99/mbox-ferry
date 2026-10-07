#!/usr/bin/env python3
"""Mbox Ferry: resumable mbox-to-IMAP migration through a local bridge."""

from __future__ import annotations

import argparse
import base64
import collections
import email
import email.policy
import email.utils
import getpass
import hashlib
import imaplib
import ipaddress
import json
import mailbox
import os
import re
import ssl
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable


VERSION = "0.2.0"
STATE_VERSION = 1
RETRY_DELAYS = (2, 5, 10, 20, 40, 60, 120, 180)
ENV_KEYS = {
    "SOURCE_FOLDER",
    "DESTINATION_FOLDER",
    "BRIDGE_USERNAME",
    "BRIDGE_PASSWORD",
    "FALLBACK_SENDER_EMAIL",
    "SENT_FROM_CUTOFF",
    "SENT_FROM_BEFORE_EMAIL",
    "SENT_FROM_AFTER_EMAIL",
}
IGNORED_NAMES = {
    "msgfilterrules.dat",
    "virtualfolders.dat",
    "foldertree.json",
    "feeds.json",
    "feeditems.json",
    "junklog.html",
    "panacea.dat",
}


def load_env_file(path: Path) -> set[str]:
    """Load Mbox Ferry values from a simple dotenv file without dependencies."""
    loaded: set[str] = set()
    if not path.is_file():
        return loaded
    for line_number, original in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
        line = original.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            raise RuntimeError(f"Invalid .env line {line_number}: expected NAME=VALUE")
        name, value = line.split("=", 1)
        name = name.strip()
        value = value.strip()
        if name not in ENV_KEYS:
            raise RuntimeError(
                f"Invalid .env line {line_number}: expected one of {', '.join(sorted(ENV_KEYS))}"
            )
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        # Values set by the shell take precedence over the local .env file.
        if name not in os.environ:
            os.environ[name] = value
            loaded.add(name)
    return loaded


@dataclass(frozen=True)
class SourceMailbox:
    path: Path
    relative_name: str
    destination: str
    count: int


def normalize_message_id(value: str | None) -> str:
    return re.sub(r"\s+", "", value or "").lower()


def source_fingerprint(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def modified_utf7(value: str) -> str:
    """Encode a Unicode mailbox name using IMAP modified UTF-7."""
    output: list[str] = []
    non_ascii: list[str] = []

    def flush() -> None:
        if not non_ascii:
            return
        encoded = base64.b64encode("".join(non_ascii).encode("utf-16-be")).decode("ascii")
        output.append("&" + encoded.rstrip("=").replace("/", ",") + "-")
        non_ascii.clear()

    for char in value:
        code = ord(char)
        if 0x20 <= code <= 0x7E:
            flush()
            output.append("&-" if char == "&" else char)
        else:
            non_ascii.append(char)
    flush()
    return "".join(output)


def quote_mailbox(name: str) -> str:
    encoded = modified_utf7(name)
    return '"' + encoded.replace("\\", "\\\\").replace('"', '\\"') + '"'


def clean_segment(value: str) -> str:
    value = value[:-5] if value.lower().endswith(".mbox") else value
    value = value.replace("/", "-").replace("\\", "-").strip()
    return value or "Unnamed"


def source_parts(root: Path, path: Path) -> list[str]:
    relative = path.relative_to(root)
    parts: list[str] = []
    for directory in relative.parts[:-1]:
        if directory.lower().endswith(".sbd"):
            directory = directory[:-4]
        parts.append(clean_segment(directory))
    parts.append(clean_segment(relative.name))
    return parts


def destination_for(root: str, parts: list[str], max_depth: int) -> str:
    base = [part for part in root.strip("/").split("/") if part]
    if not base:
        raise ValueError("Destination root cannot be empty")
    available = max_depth - len(base)
    if available < 1:
        raise ValueError(
            f"Destination {root!r} already uses the configured maximum depth of {max_depth}"
        )
    # Keep deep source trees within Proton's folder-depth limit.
    if len(parts) > available:
        parts = parts[: available - 1] + [" - ".join(parts[available - 1 :])]
    return "/".join(base + parts)


def is_candidate_mbox(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size == 0:
        return False
    name = path.name.lower()
    if name.startswith(".") or name.endswith(".msf") or name in IGNORED_NAMES:
        return False
    return path.suffix.lower() in {"", ".mbox"}


def discover_mboxes(root: Path, destination_root: str, max_depth: int) -> list[SourceMailbox]:
    if not root.is_dir():
        raise RuntimeError(f"Source directory does not exist: {root}")
    found: list[SourceMailbox] = []
    for path in sorted(root.rglob("*"), key=lambda item: str(item).lower()):
        if not is_candidate_mbox(path):
            continue
        box = mailbox.mbox(path, create=False)
        try:
            count = len(box)
        finally:
            box.close()
        if count == 0:
            continue
        parts = source_parts(root, path)
        found.append(
            SourceMailbox(
                path=path,
                relative_name=path.relative_to(root).as_posix(),
                destination=destination_for(destination_root, parts, max_depth),
                count=count,
            )
        )
    return found


def parse_internal_date(message: mailbox.mboxMessage) -> str | None:
    try:
        parsed = email.utils.parsedate_to_datetime(message.get("Date", ""))
        if parsed is None:
            return None
        if parsed.tzinfo is None:
            parsed = parsed.astimezone()
        return imaplib.Time2Internaldate(parsed)
    except (TypeError, ValueError, OverflowError):
        return None


def flags_for(message: mailbox.mboxMessage) -> str | None:
    raw = (message.get("X-Mozilla-Status") or "").strip()
    try:
        if int(raw, 16) & 0x0001:
            return "(\\Seen)"
    except ValueError:
        pass
    return None


def repaired_copy(
    raw: bytes,
    replacement_from: str,
    *,
    force_from: bool = False,
    message_id_salt: str | None = None,
) -> bytes:
    """Return a standards-compliant copy while preserving broken headers."""
    message = email.message_from_bytes(raw, policy=email.policy.default)

    def preserve_and_replace(name: str, replacement: str | None) -> None:
        values = message.get_all(name, [])
        while name in message:
            del message[name]
        for value in values:
            encoded = base64.b64encode(str(value).encode("utf-8", "replace")).decode("ascii")
            message[f"X-Original-{name}-Base64"] = encoded
        if replacement is not None:
            message[name] = replacement

    from_values = message.get_all("From", [])
    from_addresses = email.utils.getaddresses([str(value) for value in from_values])
    # Normal repair fixes invalid From headers; the optional Sent rule forces a known address.
    if force_from or not from_values or not any("@" in address for _, address in from_addresses):
        preserve_and_replace("From", replacement_from)
    for header in ("To", "Cc", "Bcc", "Reply-To"):
        values = message.get_all(header, [])
        if not values:
            continue
        addresses = email.utils.getaddresses([str(value) for value in values])
        malformed = any(address and "@" not in address for _, address in addresses)
        if malformed:
            preserve_and_replace(header, "Undisclosed recipients:;" if header == "To" else None)

    if not message.get("Date"):
        message["Date"] = email.utils.format_datetime(datetime.now().astimezone())
    if message_id_salt is not None:
        # Sender rewrites need a different, repeatable ID to avoid server-side deduplication.
        preserve_and_replace("Message-ID", None)
        digest = hashlib.sha256(raw + b"\0" + message_id_salt.encode()).hexdigest()[:32]
        message["Message-ID"] = f"<{digest}@mbox-ferry.invalid>"
    elif not message.get("Message-ID"):
        digest = hashlib.sha256(raw).hexdigest()[:32]
        message["Message-ID"] = f"<{digest}@recovered.invalid>"
    message["X-Mbox-Ferry-Repair"] = "Header-only repair; original mbox message was not modified"
    return message.as_bytes(policy=email.policy.SMTP)


def is_loopback(host: str) -> bool:
    if host.lower() == "localhost":
        return True
    try:
        return ipaddress.ip_address(host).is_loopback
    except ValueError:
        return False


def sent_sender_for(
    source: SourceMailbox,
    message: mailbox.mboxMessage,
    cutoff: date | None,
    before: str | None,
    after: str | None,
) -> str | None:
    """Choose the configured address only for recognized Sent folders."""
    if cutoff is None or before is None or after is None:
        return None
    if clean_segment(source.path.name).casefold() not in {"sent", "sent items", "sent mail"}:
        return None
    try:
        sent_at = email.utils.parsedate_to_datetime(message.get("Date", ""))
    except (TypeError, ValueError, OverflowError) as exc:
        raise RuntimeError(f"Sent message has no usable date in {source.relative_name}") from exc
    if sent_at is None:
        raise RuntimeError(f"Sent message has no usable date in {source.relative_name}")
    return before if sent_at.date() < cutoff else after


def batches(values: list[bytes], size: int = 250) -> Iterable[list[bytes]]:
    for index in range(0, len(values), size):
        yield values[index : index + size]


class Importer:
    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        state_path: Path,
        replacement_from: str,
        allow_remote: bool,
    ):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.state_path = state_path
        self.replacement_from = replacement_from
        self.allow_remote = allow_remote
        self.conn: imaplib.IMAP4 | None = None
        self.state = self.load_state()

    def load_state(self) -> dict:
        if self.state_path.exists():
            loaded = json.loads(self.state_path.read_text(encoding="utf-8"))
            if loaded.get("version") != STATE_VERSION:
                raise RuntimeError("The state file was created by an incompatible Mbox Ferry version")
            return loaded
        return {"version": STATE_VERSION, "completed": {}, "repairs": [], "failures": []}

    def save_state(self) -> None:
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.state_path.with_suffix(self.state_path.suffix + ".tmp")
        temporary.write_text(json.dumps(self.state, indent=2, sort_keys=True), encoding="utf-8")
        # Atomic replacement prevents Ctrl+C from leaving a half-written checkpoint.
        temporary.replace(self.state_path)

    def connect(self) -> None:
        self.disconnect()
        local = is_loopback(self.host)
        if not local and not self.allow_remote:
            raise RuntimeError(
                "Refusing a non-loopback IMAP host. Use --allow-remote-imap only if this is intentional."
            )
        connection = imaplib.IMAP4(self.host, self.port, timeout=120)
        context = ssl.create_default_context()
        if local:
            # Bridge uses a self-signed certificate, but only on the local machine.
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
        connection.starttls(ssl_context=context)
        status, response = connection.login(self.username, self.password)
        if status != "OK":
            raise RuntimeError(f"IMAP login failed: {response!r}")
        self.conn = connection

    def disconnect(self) -> None:
        if self.conn is not None:
            try:
                self.conn.logout()
            except Exception:
                pass
            self.conn = None

    def ensure_mailbox(self, name: str) -> None:
        assert self.conn is not None
        parts = [part for part in name.strip("/").split("/") if part]
        for index in range(1, len(parts) + 1):
            current = "/".join(parts[:index])
            status, _ = self.conn.select(quote_mailbox(current), readonly=True)
            if status == "OK":
                self.conn.unselect()
                continue
            status, response = self.conn.create(quote_mailbox(current))
            details = b" ".join(response or []).lower()
            if status != "OK" and b"exist" not in details:
                raise RuntimeError(f"Could not create destination {current!r}: {response!r}")

    def remote_inventory(self, name: str) -> collections.Counter[str]:
        assert self.conn is not None
        status, response = self.conn.select(quote_mailbox(name), readonly=True)
        if status != "OK":
            raise RuntimeError(f"Could not open destination {name!r}: {response!r}")
        result: collections.Counter[str] = collections.Counter()
        status, data = self.conn.uid("search", None, "ALL")
        uids = (data[0] or b"").split() if status == "OK" else []
        # Small batches avoid oversized IMAP commands on large folders.
        for group in batches(uids):
            sequence = b",".join(group).decode("ascii")
            status, fetched = self.conn.uid(
                "fetch", sequence, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])"
            )
            if status != "OK":
                continue
            for item in fetched:
                if not (isinstance(item, tuple) and isinstance(item[1], bytes)):
                    continue
                header = email.message_from_bytes(item[1], policy=email.policy.default)
                message_id = normalize_message_id(header.get("Message-ID"))
                if message_id:
                    result[message_id] += 1
        self.conn.unselect()
        return result

    def append_once(
        self, destination: str, raw: bytes, flags: str | None, internal_date: str | None
    ) -> tuple[str, list[bytes]]:
        assert self.conn is not None
        return self.conn.append(quote_mailbox(destination), flags, internal_date, raw)

    def append_with_retries(
        self, destination: str, raw: bytes, flags: str | None, internal_date: str | None
    ) -> tuple[bool, bool, str]:
        def is_format_error(detail: str) -> bool:
            lowered = detail.lower()
            return any(term in lowered for term in ("rfc5322", "failed to parse", "required header"))

        def is_temporary(detail: str) -> bool:
            lowered = detail.lower()
            return any(
                term in lowered
                for term in ("temporary", "try again", "timeout", "unavailable", "connection")
            )

        last_error = "unknown failure"
        for attempt, delay in enumerate(RETRY_DELAYS, start=1):
            try:
                if self.conn is None:
                    self.connect()
                status, response = self.append_once(destination, raw, flags, internal_date)
                detail = b" ".join(response or []).decode("utf-8", "replace")
                if status == "OK":
                    return True, False, ""
                if is_format_error(detail):
                    # Never rewrite a message unless Bridge rejects the original as malformed.
                    repaired = repaired_copy(raw, self.replacement_from)
                    status, repaired_response = self.append_once(
                        destination, repaired, flags, internal_date
                    )
                    if status == "OK":
                        return True, True, detail
                    repaired_detail = b" ".join(repaired_response or []).decode("utf-8", "replace")
                    return False, False, f"{detail}; repaired copy rejected: {repaired_detail}"
                if "no such mailbox" in detail.lower():
                    self.ensure_mailbox(destination)
                elif not is_temporary(detail):
                    return False, False, detail
                last_error = detail
            except (imaplib.IMAP4.abort, OSError, TimeoutError) as exc:
                last_error = str(exc)
                self.disconnect()
            except imaplib.IMAP4.error as exc:
                detail = str(exc)
                if is_format_error(detail):
                    try:
                        repaired = repaired_copy(raw, self.replacement_from)
                        status, response = self.append_once(
                            destination, repaired, flags, internal_date
                        )
                        if status == "OK":
                            return True, True, detail
                    except Exception as repair_exc:
                        return False, False, f"{detail}; repaired copy rejected: {repair_exc}"
                if not is_temporary(detail):
                    return False, False, detail
                last_error = detail

            if attempt < len(RETRY_DELAYS):
                print(f"    Temporary failure; retrying in {delay}s: {last_error}", flush=True)
                time.sleep(delay)
        return False, False, last_error


def build_parser() -> argparse.ArgumentParser:
    source_default = os.environ.get("SOURCE_FOLDER") or None
    fallback_sender = os.environ.get("FALLBACK_SENDER_EMAIL", "unknown@invalid.local")
    parser = argparse.ArgumentParser(
        description="Safely copy Thunderbird-style mbox folders through an IMAP bridge"
    )
    parser.add_argument("--version", action="version", version=f"Mbox Ferry {VERSION}")
    parser.add_argument(
        "--source",
        type=Path,
        default=Path(source_default) if source_default else None,
        required=source_default is None,
        help="Directory containing mbox files",
    )
    parser.add_argument(
        "--destination",
        default=os.environ.get("DESTINATION_FOLDER", "Folders/Imported Mail"),
        help="Destination mailbox root",
    )
    parser.add_argument("--host", default="127.0.0.1", help="IMAP host")
    parser.add_argument("--port", type=int, default=1143, help="IMAP STARTTLS port")
    parser.add_argument("--username", default=os.environ.get("BRIDGE_USERNAME"), help="IMAP username")
    parser.add_argument(
        "--state",
        type=Path,
        default=Path("mbox-ferry-state.json"),
        help="Checkpoint file",
    )
    parser.add_argument(
        "--replacement-from",
        default=f"Recovered message <{fallback_sender}>",
        help="Sender used only when repairing a malformed From header",
    )
    parser.add_argument(
        "--sent-from-cutoff",
        type=date.fromisoformat,
        default=os.environ.get("SENT_FROM_CUTOFF") or None,
        help="Date when Sent-folder messages switch sender address (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--sent-from-before",
        default=os.environ.get("SENT_FROM_BEFORE_EMAIL") or None,
        help="Sender address before --sent-from-cutoff",
    )
    parser.add_argument(
        "--sent-from-after",
        default=os.environ.get("SENT_FROM_AFTER_EMAIL") or None,
        help="Sender address on or after --sent-from-cutoff",
    )
    parser.add_argument("--max-depth", type=int, default=3, help="Maximum destination hierarchy depth")
    parser.add_argument("--delay", type=float, default=0.15, help="Seconds between uploads")
    parser.add_argument("--execute", action="store_true", help="Perform uploads")
    parser.add_argument(
        "--allow-remote-imap", action="store_true", help="Permit a non-loopback IMAP host"
    )
    return parser


def print_plan(sources: list[SourceMailbox], args: argparse.Namespace) -> None:
    total = sum(source.count for source in sources)
    print("Mbox Ferry import plan")
    print(f"Source: {args.source.resolve()}")
    print(f"Destination server: {args.host}:{args.port}")
    print(f"Destination root: {args.destination}")
    print(f"Mailboxes: {len(sources)}; messages: {total}")
    print("Source files will not be modified, moved, or deleted.\n")
    for source in sources:
        print(f"  {source.relative_name} ({source.count}) -> {source.destination}")
    if not args.execute:
        print("\nDry run only. Repeat with --execute to begin copying.")


def run_import(args: argparse.Namespace, sources: list[SourceMailbox]) -> int:
    if not args.username:
        raise RuntimeError("--username is required for an actual import")
    password = os.environ.get("BRIDGE_PASSWORD")
    if password:
        print("Using the Bridge password from .env.")
    else:
        password = getpass.getpass("IMAP Bridge password: ")
    if not password:
        raise RuntimeError("An IMAP password is required")

    importer = Importer(
        args.host,
        args.port,
        args.username,
        password,
        args.state,
        args.replacement_from,
        args.allow_remote_imap,
    )
    uploaded_total = skipped_total = failed_total = 0
    importer.connect()
    try:
        for source in sources:
            importer.ensure_mailbox(source.destination)
            remote_ids = importer.remote_inventory(source.destination)
            completed = set(importer.state["completed"].get(source.relative_name, []))
            source_occurrences: collections.Counter[str] = collections.Counter()
            uploaded = skipped = failed = 0
            print(f"\n{source.relative_name}: {source.count} messages", flush=True)

            box = mailbox.mbox(source.path, create=False)
            try:
                for index, message in enumerate(box, start=1):
                    raw = message.as_bytes(unixfrom=False)
                    # Transform before hashing so checkpoints describe the corrected copy.
                    sender = sent_sender_for(
                        source,
                        message,
                        args.sent_from_cutoff,
                        args.sent_from_before,
                        args.sent_from_after,
                    )
                    if sender:
                        salt = f"sent-from:{args.sent_from_cutoff.isoformat()}:{sender}"
                        raw = repaired_copy(
                            raw,
                            sender,
                            force_from=True,
                            message_id_salt=salt,
                        )
                    digest = source_fingerprint(raw)
                    prepared = email.message_from_bytes(raw, policy=email.policy.default)
                    message_id = normalize_message_id(prepared.get("Message-ID"))
                    if message_id:
                        source_occurrences[message_id] += 1
                    # Remote IDs cover fresh reruns; fingerprints cover messages without IDs.
                    already_remote = bool(
                        message_id and remote_ids[message_id] >= source_occurrences[message_id]
                    )
                    if digest in completed or already_remote:
                        skipped += 1
                        continue

                    ok, repaired, detail = importer.append_with_retries(
                        source.destination, raw, flags_for(message), parse_internal_date(message)
                    )
                    if ok:
                        uploaded += 1
                        completed.add(digest)
                        if message_id:
                            remote_ids[message_id] += 1
                        importer.state["completed"][source.relative_name] = sorted(completed)
                        if repaired:
                            importer.state["repairs"].append(
                                {"folder": source.relative_name, "index": index, "reason": detail}
                            )
                    else:
                        failed += 1
                        importer.state["failures"].append(
                            {"folder": source.relative_name, "index": index, "reason": detail}
                        )
                    importer.save_state()

                    if index % 25 == 0 or index == source.count:
                        print(
                            f"  {index}/{source.count}: uploaded={uploaded}, "
                            f"already present={skipped}, failed={failed}",
                            flush=True,
                        )
                    time.sleep(max(0.0, args.delay))
            finally:
                box.close()

            uploaded_total += uploaded
            skipped_total += skipped
            failed_total += failed
    except KeyboardInterrupt:
        print("\nPaused safely. Repeat the same command to resume.")
        return 130
    finally:
        importer.disconnect()

    print("\nImport pass finished")
    print(f"Uploaded: {uploaded_total}")
    print(f"Already present/checkpointed: {skipped_total}")
    print(f"Failed after retries: {failed_total}")
    print(f"Private checkpoint: {args.state.resolve()}")
    return 0 if failed_total == 0 else 2


def main(argv: list[str] | None = None) -> int:
    # Support both running beside .env and running from a configured working directory.
    project_env = Path(__file__).resolve().with_name(".env")
    load_env_file(project_env)
    working_env = Path.cwd() / ".env"
    if working_env.resolve() != project_env:
        load_env_file(working_env)
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.max_depth < 2:
        parser.error("--max-depth must be at least 2")
    if args.delay < 0:
        parser.error("--delay cannot be negative")
    sender_rule = (args.sent_from_cutoff, args.sent_from_before, args.sent_from_after)
    if any(sender_rule) and not all(sender_rule):
        parser.error(
            "SENT_FROM_CUTOFF, SENT_FROM_BEFORE_EMAIL, and SENT_FROM_AFTER_EMAIL "
            "must be set together"
        )
    for address in (args.sent_from_before, args.sent_from_after):
        if address and "@" not in email.utils.parseaddr(address)[1]:
            parser.error(f"Invalid sender email address: {address}")

    sources = discover_mboxes(args.source, args.destination, args.max_depth)
    if not sources:
        raise RuntimeError("No non-empty mbox files were found in the source directory")
    print_plan(sources, args)
    if not args.execute:
        return 0
    return run_import(args, sources)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"\nMbox Ferry stopped: {exc}", file=sys.stderr)
        raise SystemExit(1)
