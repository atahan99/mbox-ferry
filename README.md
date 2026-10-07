# Mbox Ferry

Mbox Ferry is a small Python tool I made to copy Thunderbird mbox folders into
Proton Mail through Proton Bridge.

## Why I made it

Proton Easy Switch supports Outlook, but school and work Microsoft 365 accounts
may require administrator approval. That was a problem for my university email.
Proton's local-file import workflow also uses MBOX or EML files rather than a
direct Outlook PST upload.

One workaround is to get the Outlook mailbox into Thunderbird Local Folders,
connect Proton to Thunderbird through Proton Bridge, and copy the messages from
there. Large manual copies can stall, and malformed older messages may be
rejected. Mbox Ferry automates that last step: it uploads messages one at a time,
retries temporary failures, saves its progress, and resumes after being stopped.

## What it does

- Copies mail without changing the original mbox files.
- Performs a dry run unless `--execute` is provided.
- Resumes from a local checkpoint.
- Avoids ordinary duplicate uploads using Message-ID values.
- Repairs some malformed legacy email headers when Proton rejects them.
- Preserves nested Thunderbird folders, flattening very deep folder trees.

## Requirements

- Python 3.10 or newer
- Proton Bridge running on the computer
- Thunderbird mbox files or folders

## Setup

Copy `.env.example` to `.env`, then enter the source folder and the IMAP details
shown in Proton Bridge under **Mailbox details**:

```dotenv
SOURCE_FOLDER=C:\path\to\Thunderbird\Mail\Local Folders
DESTINATION_FOLDER=Folders/Imported Mail
BRIDGE_USERNAME=bridge-username
BRIDGE_PASSWORD=bridge-generated-password
```

The Bridge password is not your normal Proton account password.

### Optional Sent-folder sender repair

Some Outlook exports contain only a display name in the `From` header. To
replace those invalid senders by message date, add these values to `.env`:

```dotenv
FALLBACK_SENDER_EMAIL=unknown@invalid.local
SENT_FROM_CUTOFF=2023-01-01
SENT_FROM_BEFORE_EMAIL=old-address@example.com
SENT_FROM_AFTER_EMAIL=new-address@example.com
```

The rule applies only to standard Sent folders. Corrected copies receive stable
new Message-IDs so the destination server cannot reuse an earlier malformed
copy. Leave the three `SENT_FROM_*` values blank to disable the rule.

## Run it

First run a preview:

```console
python mbox_ferry.py
```

If the folder mapping looks correct, start the import:

```console
python mbox_ferry.py --execute
```

Keep Proton Bridge running. You can press Ctrl+C at any time and run the same
command again later to resume.

Mbox Ferry uses `127.0.0.1:1143` by default. If Bridge shows something different,
use the `--host` and `--port` options.

The progress file is `mbox-ferry-state.json`. Keep it until you have verified
the imported mail.

## Important

Back up your mailbox before importing important mail. Mbox Ferry copies messages
and does not delete the source, but it is still an early personal project and is
not affiliated with Proton or Mozilla.
