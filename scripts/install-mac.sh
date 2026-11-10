#!/bin/sh
# Install or update the bridge on the Mac that runs Word.
#   scripts/install-mac.sh                  copy the code, create token and config if missing
#   scripts/install-mac.sh --launch-agent   also start the server at login (launchd)
# Never overwrites an existing token or config.json.
set -eu
SRC="$(cd "$(dirname "$0")/.." && pwd)"
DEST="${WORD_BRIDGE_HOME:-$HOME/.word-bridge}"
umask 077
mkdir -p "$DEST"
chmod 700 "$DEST"
for f in server.py client.py word_ops.applescript; do
	cp "$SRC/$f" "$DEST/$f"
	chmod 600 "$DEST/$f"
done
if [ ! -s "$DEST/token" ]; then
	/usr/bin/python3 -c 'import secrets; print(secrets.token_hex(32))' > "$DEST/token"
	echo "created a new access token: $DEST/token"
fi
chmod 600 "$DEST/token"
if [ ! -f "$DEST/config.json" ]; then
	cp "$SRC/config.example.json" "$DEST/config.json"
	chmod 600 "$DEST/config.json"
	echo "created $DEST/config.json from the example - edit it before starting the server"
fi
echo "installed code into $DEST"

if [ "${1:-}" = "--launch-agent" ]; then
	PLIST="${WORD_BRIDGE_PLIST:-$HOME/Library/LaunchAgents/local.word-bridge.plist}"
	mkdir -p "$(dirname "$PLIST")"
	sed "s|__BRIDGE_HOME__|$DEST|g" "$SRC/launchd/local.word-bridge.plist.template" > "$PLIST"
	echo "wrote $PLIST"
	if [ -z "${WORD_BRIDGE_PLIST:-}" ]; then
		launchctl bootout "gui/$(id -u)/local.word-bridge" 2>/dev/null || true
		launchctl bootstrap "gui/$(id -u)" "$PLIST"
		echo "server started; it will also start at every login"
	fi
else
	echo "start the server with: /usr/bin/python3 $DEST/server.py"
fi
