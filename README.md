# word-bridge

Let a Claude Code session on a remote machine make **tracked edits in a Word document that is open on your Mac**. The change appears in the window you are looking at, as a revision you can accept or reject. No file is replaced, and nothing needs to be closed and reopened.

[中文说明](README.zh-CN.md)

## Why this exists

If you write in Word on a Mac while an AI agent works on a remote server, the obvious approach is to sync the `.docx` file. That does not work well:

- Word does not reload a document whose file was replaced on disk. The open window keeps showing the old text, and your next save overwrites the incoming version.
- With AutoSave on a OneDrive file it is worse: in our test Word never refreshed, and the first keystroke silently overwrote the replaced file.
- Word's real-time co-authoring only admits Word clients. There is no API for a program to join a session.

word-bridge takes the other route: the remote agent never touches the file. It sends small instructions ("replace this sentence with that one"), and Word itself applies them to the document it has in memory.

## How it works

```
Claude Code on the remote host
   │  MCP tool call, e.g. word_replace_text
   ▼
client.py (stdio MCP server, remote side)
   │  Unix socket on the remote host
   ▼
SSH reverse tunnel, inside the SSH connection you already have open
   │
   ▼
server.py on the Mac        checks the token and the document allow-list
   │  osascript
   ▼
word_ops.applescript        turns on Track Changes, locates the text, rewrites it
   ▼
Microsoft Word              the open window updates immediately
```

Only text travels over the tunnel: the instructions, and document text when the agent reads. The `.docx` file is never transferred.

## Tools

| Tool | What it does |
|---|---|
| `word_list_documents` | Lists the allowed documents and whether each is open in Word |
| `word_read_document` | Returns the current text by paragraph, including unsaved edits |
| `word_find_text` | Finds a string and shows each hit with context |
| `word_replace_text` | Replaces text as a tracked change (empty replacement deletes) |
| `word_insert_text` | Inserts text before or after an anchor, as a tracked change |
| `word_add_comment` | Attaches a Word comment to a piece of text |
| `word_list_changes` | Lists pending revisions and comments |

Every write is a tracked change or a comment, authored as the name set in `config.json` (default `Claude`). The bridge never saves, closes, opens, accepts or rejects anything.

## Requirements

- macOS with Microsoft Word, and Python 3.9 or later at `/usr/bin/python3`
- A remote host you reach over SSH with connection sharing (`ControlMaster`) enabled, with Python 3.9 or later
- Claude Code on the remote host (any MCP client that launches stdio servers should work)

## Install

**1. On the Mac**

```bash
git clone https://github.com/PeixiongYuan/word-bridge.git
cd word-bridge
scripts/install-mac.sh
```

This copies the code to `~/.word-bridge`, generates an access token, and creates `config.json` from the example. Edit `~/.word-bridge/config.json`:

- `allowed_documents`: file names the agent may touch
- `ssh_host`: the `Host` in your `~/.ssh/config` where Claude Code runs
- `remote_socket`: where the tunnel should appear on that host (the next step prints the right value)

**2. On the remote host**, while you are logged in to it:

```bash
scripts/install-remote.sh myserver
```

It copies `client.py` and the token, and prints the `claude mcp add` command to run in your project directory there.

**3. Start the server**

```bash
/usr/bin/python3 ~/.word-bridge/server.py
```

or have it start at every login:

```bash
scripts/install-mac.sh --launch-agent
```

The first time the server talks to Word, macOS may ask whether to allow it to control Microsoft Word.

**4. Optional: stop the permission prompts** by allowing the tools in the remote project's `.claude/settings.local.json`:

```json
{ "permissions": { "allow": ["mcp__word-bridge__word_read_document", "mcp__word-bridge__word_replace_text"] } }
```

## The SSH tunnel

The server keeps a reverse tunnel alive by itself. It only ever rides the SSH master connection you already opened; if that connection is gone it does nothing and never tries to authenticate. This matters on hosts that need a one-time password: after you log in again, the tunnel is back within about 30 seconds.

Your SSH config needs connection sharing for the host, for example:

```
Host myserver
    ControlMaster auto
    ControlPath ~/.ssh/sockets/%r@%h-%p
    ControlPersist 8h
```

## Security model

- The Mac socket and the remote socket are Unix sockets readable only by you (`0600` inside a `0700` directory). On a shared login or compute node, other users cannot connect.
- Every connection must present the token before anything else is processed.
- Only documents named in `allowed_documents`, and currently open in Word, can be read or edited.
- All writes are logged to `~/.word-bridge/audit.log`.
- The agent can read the text of allowed documents. Do not allow documents you would not show it.

## Limitations

- **Text only.** No figures, table structure or layout. Tables are read as flattened text.
- **Matching is exact.** The anchor text must match character for character and be unique, or the call is refused. Anchors that contain citation fields, footnote marks or existing revisions may fail the verification step; the bridge then makes no change.
- **Documents are matched by file name, not path.** Two files with the same name are indistinguishable.
- **Authorship is set by briefly switching Word's user name** during each write and restoring it afterwards. If you type in that instant, your keystrokes may be attributed to the agent.
- **Reading excludes tracked deletions.** The text is what you would get after accepting all changes.
- **Tested on one setup:** Word for Mac 16.114, macOS client with OpenSSH 10, a Linux remote host, Claude Code 2.1. Word for Windows is not supported (the bridge uses AppleScript).
- Tool descriptions and messages are currently written in Chinese.

## Troubleshooting

- `python3 ~/.word-bridge/client.py --ping` on the remote host prints `ok` when the whole chain works.
- `~/.word-bridge/server.log` on the Mac shows the tunnel state (`up`, `no-master`, `forward-failed`).
- If the remote host refuses Unix-socket forwarding, its `sshd` has `AllowStreamLocalForwarding` disabled.

## License

MIT
