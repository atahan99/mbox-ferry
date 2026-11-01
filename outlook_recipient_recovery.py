#!/usr/bin/env python3
"""Recover missing Sent-folder recipients from classic Outlook on Windows."""

from __future__ import annotations

import argparse
import base64
import email
import email.policy
import email.utils
import getpass
import json
import mailbox
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import mbox_ferry as ferry


# Outlook COM is Windows-only; PowerShell avoids adding a pywin32 dependency.
POWERSHELL_EXPORT = r"""
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')
function Get-SentFolders($folder) {
    foreach ($child in $folder.Folders) {
        # shortcut: English folder names; add localized names when needed.
        if ($child.Name -match '^(Sent|Sent Items|Sent Mail)$') { $child }
        Get-SentFolders $child
    }
}
$rows = foreach ($store in $namespace.Stores) {
    foreach ($folder in @(Get-SentFolders ($store.GetRootFolder()))) {
      foreach ($item in $folder.Items) {
        if ($item.Class -ne 43) { continue }
        $recipients = foreach ($recipient in $item.Recipients) {
            $null = $recipient.Resolve()
            $smtp = ''
            try {
                $smtp = [string]$recipient.PropertyAccessor.GetProperty(
                    'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'
                )
            } catch {}
            if (-not $smtp) {
                try {
                    $address = [string]$recipient.AddressEntry.Address
                    if ($address -match '@') { $smtp = $address }
                } catch {}
            }
            [PSCustomObject]@{
                Type = [int]$recipient.Type
                Name = [string]$recipient.Name
                Smtp = $smtp
            }
        }
        [PSCustomObject]@{
            Store = [string]$store.DisplayName
            Folder = [string]$folder.FolderPath
            Subject = [string]$item.Subject
            SentOn = $item.SentOn.ToUniversalTime().ToString('o')
            Recipients = @($recipients)
        }
      }
    }
}
ConvertTo-Json -InputObject @($rows) -Depth 5 -Compress
"""


def load_environment() -> None:
    project_env = Path(__file__).resolve().with_name(".env")
    ferry.load_env_file(project_env)
    working_env = Path.cwd() / ".env"
    if working_env.resolve() != project_env:
        ferry.load_env_file(working_env)


def export_outlook(store: str | None) -> list[dict]:
    if sys.platform != "win32":
        raise RuntimeError("Outlook recipient recovery is available only on Windows")
    command = base64.b64encode(POWERSHELL_EXPORT.encode("utf-16le")).decode("ascii")
    result = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", command],
        capture_output=True,
        text=True,
        check=True,
    )
    rows = json.loads(result.stdout.lstrip("\ufeff") or "[]")
    if store:
        rows = [row for row in rows if row["Store"].casefold() == store.casefold()]
        if not rows:
            raise RuntimeError(f"Outlook store was not found: {store}")
    return rows


def normalized_subject(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def match_messages(messages: list[mailbox.mboxMessage], rows: list[dict]) -> list[dict]:
    available = set(range(len(rows)))
    matches = []
    for message in messages:
        sent_at = email.utils.parsedate_to_datetime(message.get("Date", "")).astimezone(timezone.utc)
        subject = normalized_subject(str(message.get("Subject", "")))
        candidates = sorted(
            (
                abs(
                    (
                        datetime.fromisoformat(rows[index]["SentOn"].replace("Z", "+00:00"))
                        .astimezone(timezone.utc)
                        - sent_at
                    ).total_seconds()
                ),
                index,
            )
            for index in available
            if normalized_subject(rows[index]["Subject"]) == subject
        )
        # Same subject, nearest timestamp, and never guess an equally close match.
        if not candidates or candidates[0][0] > 180:
            raise RuntimeError(f"No Outlook match for {message.get('Subject')!r}")
        if len(candidates) > 1 and abs(candidates[1][0] - candidates[0][0]) < 0.5:
            raise RuntimeError(f"Ambiguous Outlook match for {message.get('Subject')!r}")
        _, index = candidates[0]
        available.remove(index)
        matches.append(rows[index])
    return matches


def preserve_header(message: email.message.Message, name: str) -> None:
    """Keep the imported value before replacing it with Outlook's address."""
    backup = f"X-Original-{name}-Base64"
    if message.get(backup):
        return
    for value in message.get_all(name, []):
        encoded = base64.b64encode(str(value).encode("utf-8", "replace")).decode("ascii")
        message[backup] = encoded


def corrected_copy(
    source: ferry.SourceMailbox,
    message: mailbox.mboxMessage,
    outlook_row: dict,
    args: argparse.Namespace,
) -> tuple[bytes, str, int]:
    raw = message.as_bytes(unixfrom=False)
    sender = ferry.sent_sender_for(
        source,
        message,
        args.sent_from_cutoff,
        args.sent_from_before,
        args.sent_from_after,
    )
    # A stable new ID prevents Proton from reusing the earlier malformed copy.
    fixed = ferry.repaired_copy(
        raw,
        sender or args.replacement_from,
        force_from=bool(sender),
        message_id_salt=f"outlook-recipient-recovery-v1:{sender or ''}",
    )
    parsed = email.message_from_bytes(fixed, policy=email.policy.default)
    # Outlook recipient types: 1=To, 2=Cc, 3=Bcc.
    groups = {1: [], 2: [], 3: []}
    for recipient in outlook_row["Recipients"]:
        address = recipient["Smtp"]
        if "@" not in address:
            raise RuntimeError(f"Outlook recipient has no SMTP address: {recipient['Name']!r}")
        groups[int(recipient["Type"])].append(email.utils.formataddr((recipient["Name"], address)))

    for name, kind in (("To", 1), ("Cc", 2), ("Bcc", 3)):
        preserve_header(parsed, name)
        while name in parsed:
            del parsed[name]
        if groups[kind]:
            parsed[name] = ", ".join(groups[kind])
    parsed["X-Mbox-Ferry-Recipient-Fix"] = "Recipients restored from Outlook"
    return (
        parsed.as_bytes(policy=email.policy.SMTP),
        str(parsed["Message-ID"]).lower(),
        sum(map(len, groups.values())),
    )


def main(argv: list[str] | None = None) -> int:
    load_environment()
    parser = ferry.build_parser()
    parser.description = "Recover Sent-folder recipients from classic Outlook"
    parser.add_argument("--outlook-store", default=os.environ.get("OUTLOOK_STORE") or None)
    parser.add_argument(
        "--recovery-folder",
        default=os.environ.get("OUTLOOK_RECOVERY_FOLDER", "Sent Items Recovered"),
    )
    args = parser.parse_args(argv)
    ferry.validate_args(parser, args)

    sources = [
        source
        for source in ferry.discover_mboxes(args.source, args.destination, args.max_depth)
        if ferry.clean_segment(source.path.name).casefold() in {"sent", "sent items", "sent mail"}
    ]
    if len(sources) != 1:
        raise RuntimeError(f"Expected one Sent mailbox, found {len(sources)}")
    source = sources[0]
    box = mailbox.mbox(source.path, create=False)
    try:
        messages = list(box)
    finally:
        box.close()

    rows = export_outlook(args.outlook_store)
    matches = match_messages(messages, rows)
    prepared = [
        (message, *corrected_copy(source, message, row, args))
        for message, row in zip(messages, matches, strict=True)
    ]
    recipients = sum(count for _, _, _, count in prepared)
    destination = ferry.destination_for(
        args.destination,
        [ferry.clean_segment(args.recovery_folder)],
        args.max_depth,
    )
    print(f"Matched {len(prepared)} Sent messages with {recipients} resolved recipients")
    print(f"Destination: {destination}")
    # Dry-run is the default so Outlook and Proton stay untouched until explicitly requested.
    if not args.execute:
        print("Dry run only. Repeat with --execute to upload.")
        return 0

    if not args.username:
        raise RuntimeError("BRIDGE_USERNAME or --username is required")
    password = os.environ.get("BRIDGE_PASSWORD") or getpass.getpass("IMAP Bridge password: ")
    importer = ferry.Importer(
        args.host,
        args.port,
        args.username,
        password,
        args.state,
        args.replacement_from,
        args.allow_remote_imap,
    )
    uploaded = skipped = failed = 0
    importer.connect()
    try:
        importer.ensure_mailbox(destination)
        remote_ids = importer.remote_inventory(destination)
        for index, (message, raw, message_id, _) in enumerate(prepared, 1):
            if remote_ids[message_id]:
                skipped += 1
                continue
            ok, _, detail = importer.append_with_retries(
                destination, raw, ferry.flags_for(message), ferry.parse_internal_date(message)
            )
            if ok:
                uploaded += 1
                remote_ids[message_id] += 1
            else:
                failed += 1
                print(f"Message {index} failed: {detail}")
        final_ids = importer.remote_inventory(destination)
        verified = sum(bool(final_ids[message_id]) for _, _, message_id, _ in prepared)
    finally:
        importer.disconnect()

    print(f"Uploaded={uploaded}, skipped={skipped}, failed={failed}, verified={verified}")
    return 0 if verified == len(prepared) and failed == 0 else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"\nOutlook recovery stopped: {exc}", file=sys.stderr)
        raise SystemExit(1)
