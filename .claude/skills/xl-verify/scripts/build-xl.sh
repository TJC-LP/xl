#!/usr/bin/env bash
# Build the xl CLI from a checkout and write a launcher for it.
# Usage: build-xl.sh <checkout> <launcher-path>
#   .claude/skills/xl-verify/scripts/build-xl.sh "$PWD" "$RUN/xl-branch"
# The launcher runs the assembly JAR on the JDK Mill provisions for xl-cli (falling back to
# $JAVA_HOME, then `java` on PATH, as in cloud sessions), filters the JVM's "WARNING:" lines from
# stderr, and runs a copy of the JAR kept beside it, so a later build in the same checkout cannot
# change what it runs.
set -euo pipefail
[ $# -eq 2 ] || { echo "usage: $0 <checkout> <launcher-path>" >&2; exit 2; }
CHECKOUT=$(cd "$1" && pwd)
LAUNCHER=$2
mkdir -p "$(dirname "$LAUNCHER")"
LAUNCHER=$(cd "$(dirname "$LAUNCHER")" && pwd)/$(basename "$LAUNCHER")
cd "$CHECKOUT"
./mill xl-cli.assembly >/dev/null
JAVA_BIN=""
JH=$(./mill show xl-cli.javaHome 2>/dev/null | tail -1 | tr -d '"' | sed 's/^qref:v1:[^:]*://' || true)
if [ -n "$JH" ] && [ -x "$JH/bin/java" ]; then JAVA_BIN="$JH/bin/java"
elif [ -n "${JAVA_HOME:-}" ] && [ -x "$JAVA_HOME/bin/java" ]; then JAVA_BIN="$JAVA_HOME/bin/java"
elif command -v java >/dev/null && java -version >/dev/null 2>&1; then JAVA_BIN=$(command -v java)
fi
[ -n "$JAVA_BIN" ] || { echo "no usable java (tried xl-cli.javaHome, JAVA_HOME, PATH)" >&2; exit 1; }
JAR="$LAUNCHER.jar"
cp out/xl-cli/assembly.dest/out.jar "$JAR"
cat > "$LAUNCHER" <<LAUNCH
#!/usr/bin/env bash
exec "$JAVA_BIN" -jar "$JAR" "\$@" 2> >(grep -v '^WARNING:' >&2)
LAUNCH
chmod +x "$LAUNCHER"
echo "$LAUNCHER -> $(git -C "$CHECKOUT" rev-parse --short HEAD) ($("$LAUNCHER" --version 2>/dev/null))"
