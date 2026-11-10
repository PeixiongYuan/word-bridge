#!/bin/sh
# Copy the client and the access token to the remote host where Claude Code runs.
#   scripts/install-remote.sh <ssh-host>
# <ssh-host> is a Host from your ~/.ssh/config. Run this while you are logged in to it.
set -eu
[ $# -eq 1 ] || { echo "usage: $0 <ssh-host>" >&2; exit 2; }
HOST="$1"
SRC="$(cd "$(dirname "$0")/.." && pwd)"
LOCAL="${WORD_BRIDGE_HOME:-$HOME/.word-bridge}"
[ -s "$LOCAL/token" ] || { echo "no token in $LOCAL - run scripts/install-mac.sh first" >&2; exit 1; }
ssh -o BatchMode=yes "$HOST" 'umask 077; mkdir -p ~/.word-bridge && chmod 700 ~/.word-bridge'
scp -q -o BatchMode=yes "$SRC/client.py" "$LOCAL/token" "$HOST:.word-bridge/"
ssh -o BatchMode=yes "$HOST" 'chmod 600 ~/.word-bridge/client.py ~/.word-bridge/token; echo "installed on $(hostname): $HOME/.word-bridge"; echo "set remote_socket in the Mac config.json to: $HOME/.word-bridge/bridge.sock"; echo "register with Claude Code there, from your project directory:"; echo "  claude mcp add word-bridge -- python3 $HOME/.word-bridge/client.py"'
