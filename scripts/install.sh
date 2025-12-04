#!/usr/bin/env bash
# XL CLI Installer
# Installs xl CLI from a packaged skill distribution
#
# Usage: ./install.sh [PREFIX]
#   PREFIX defaults to ~/.local
#
# This script is designed to work in containers and CI environments.

set -e

PREFIX="${1:-${PREFIX:-$HOME/.local}}"
BINDIR="$PREFIX/bin"
SHAREDIR="$PREFIX/share/xl"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "Installing xl CLI to $PREFIX..."

# Create directories
mkdir -p "$BINDIR"
mkdir -p "$SHAREDIR"

# Copy JAR
if [ -f "$SCRIPT_DIR/xl.jar" ]; then
    cp "$SCRIPT_DIR/xl.jar" "$SHAREDIR/"
else
    echo "Error: xl.jar not found in $SCRIPT_DIR"
    exit 1
fi

# Create wrapper script with JVM flags:
#   -Dcats.effect.warnOnNonMainThreadDetected=false : CLI runs synchronously on main thread
#   --add-opens java.base/java.lang=ALL-UNNAMED     : Required for Scala 3 reflection on JDK 17+
#   -XX:+IgnoreUnrecognizedVMOptions                : Cross-JDK compatibility
#   stderr filter: Suppress Scala 3.5+ LazyVals deprecation warnings
cat > "$BINDIR/xl" << 'WRAPPER'
#!/usr/bin/env bash
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
JAR_PATH="$SCRIPT_DIR/../share/xl/xl.jar"
exec java \
  -Dcats.effect.warnOnNonMainThreadDetected=false \
  --add-opens java.base/java.lang=ALL-UNNAMED \
  -XX:+IgnoreUnrecognizedVMOptions \
  -jar "$JAR_PATH" "$@" \
  2> >(grep -vE "(sun.misc.Unsafe|terminally deprecated|LazyVals|will be removed)" >&2)
WRAPPER

chmod +x "$BINDIR/xl"

echo "Installed xl to $BINDIR/xl"
echo "JAR installed to $SHAREDIR/xl.jar"
echo ""
echo "Ensure $BINDIR is in your PATH:"
echo "  export PATH=\"\$HOME/.local/bin:\$PATH\""
