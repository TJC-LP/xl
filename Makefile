# XL CLI Installation
# Usage: make install    (installs native binary to ~/.local/bin/xl)
#        make install-jar (installs JAR-based version requiring JDK)
#        make uninstall  (removes ~/.local/bin/xl)

PREFIX ?= $(HOME)/.local
BINDIR ?= $(PREFIX)/bin
SHAREDIR ?= $(PREFIX)/share/xl
JAR_PATH = out/xl-cli/assembly.dest/out.jar
NATIVE_PATH = out/xl-cli/nativeImage.dest/native-executable

.PHONY: build build-native install install-jar uninstall clean help package-skill package-dist

help:
	@echo "XL CLI Makefile"
	@echo ""
	@echo "Usage:"
	@echo "  make install       Build and install native binary to $(BINDIR) (recommended)"
	@echo "  make install-jar   Build and install JAR version (requires JDK 17+)"
	@echo "  make uninstall     Remove xl from $(BINDIR)"
	@echo "  make build         Build the fat JAR only"
	@echo "  make build-native  Build the native binary only"
	@echo "  make package-skill Package skill files only for Anthropic API (<1MB)"
	@echo "  make package-dist  Full distribution with JAR (~30MB)"
	@echo "  make clean         Remove build artifacts"
	@echo ""
	@echo "Options:"
	@echo "  PREFIX=<path>   Install prefix (default: ~/.local)"
	@echo "  BINDIR=<path>   Binary directory (default: PREFIX/bin)"

build:
	./mill xl-cli.assembly

build-native:
	./mill xl-cli.nativeImage

# Default install: native binary (no JDK required, instant startup)
install: build-native
	@mkdir -p $(BINDIR)
	@cp $(NATIVE_PATH) $(BINDIR)/xl
	@chmod +x $(BINDIR)/xl
	@echo "Installed xl native binary to $(BINDIR)/xl"
	@echo ""
	@echo "Ensure $(BINDIR) is in your PATH:"
	@echo '  export PATH="$$HOME/.local/bin:$$PATH"'

# JAR-based install (requires JDK 17+)
install-jar: build
	@mkdir -p $(BINDIR)
	@mkdir -p $(SHAREDIR)
	@cp $(JAR_PATH) $(SHAREDIR)/xl.jar
# Generate relocatable wrapper script with JVM flags:
#   -Dcats.effect.warnOnNonMainThreadDetected=false : CLI runs synchronously on main thread
#   --add-opens java.base/java.lang=ALL-UNNAMED     : Required for Scala 3 reflection on JDK 17+
#   -XX:+IgnoreUnrecognizedVMOptions                : Cross-JDK compatibility (older JDKs ignore --add-opens)
#   stderr filter: Suppress Scala 3.5+ LazyVals deprecation warnings from transitive dependencies
#
# Path resolution: Uses dirname to find JAR relative to script location.
# Compatible with macOS (no readlink -f) and Linux.
	@echo '#!/usr/bin/env bash' > $(BINDIR)/xl
	@echo 'SCRIPT_DIR="$$(cd "$$(dirname "$$0")" && pwd)"' >> $(BINDIR)/xl
	@echo 'JAR_PATH="$$SCRIPT_DIR/../share/xl/xl.jar"' >> $(BINDIR)/xl
	@echo 'exec java \' >> $(BINDIR)/xl
	@echo '  -Dcats.effect.warnOnNonMainThreadDetected=false \' >> $(BINDIR)/xl
	@echo '  --add-opens java.base/java.lang=ALL-UNNAMED \' >> $(BINDIR)/xl
	@echo '  -XX:+IgnoreUnrecognizedVMOptions \' >> $(BINDIR)/xl
	@echo '  -jar "$$JAR_PATH" "$$@" \' >> $(BINDIR)/xl
	@echo '  2> >(grep -vE "(sun.misc.Unsafe|terminally deprecated|LazyVals|will be removed)" >&2)' >> $(BINDIR)/xl
	@chmod +x $(BINDIR)/xl
	@echo "Installed xl to $(BINDIR)/xl"
	@echo "JAR installed to $(SHAREDIR)/xl.jar"
	@echo ""
	@echo "Ensure $(BINDIR) is in your PATH:"
	@echo '  export PATH="$$HOME/.local/bin:$$PATH"'

uninstall:
	@rm -f $(BINDIR)/xl
	@rm -f $(SHAREDIR)/xl.jar
	@rmdir $(SHAREDIR) 2>/dev/null || true
	@echo "Uninstalled xl from $(BINDIR)/xl"

clean:
	./mill clean
	@rm -rf dist/

# Package skill files only for Anthropic Skills API (<1MB)
# Does NOT include JAR - CLI must be pre-installed in container
package-skill:
	@echo "Packaging skill files for Anthropic API..."
	@rm -rf dist/xl-skill
	@mkdir -p dist/xl-skill/reference
	@cp .claude/skills/xl-cli/SKILL.md dist/xl-skill/
	@cp .claude/skills/xl-cli/reference/*.md dist/xl-skill/reference/
	@cd dist && zip -r xl-skill.zip xl-skill
	@echo ""
	@echo "Created dist/xl-skill.zip"
	@unzip -l dist/xl-skill.zip
	@du -h dist/xl-skill.zip
	@echo ""
	@echo "Upload to Anthropic Skills API:"
	@echo "  - Upload .claude/skills/xl-cli/ directory, OR"
	@echo "  - Upload dist/xl-skill.zip"
	@echo ""
	@echo "Note: xl CLI must be pre-installed in the container."
	@echo "Use 'make package-dist' for full distribution with JAR."

# Full distribution with JAR for local/container install (~30MB)
package-dist: build
	@echo "Packaging full distribution with JAR..."
	@rm -rf dist/xl-dist
	@mkdir -p dist/xl-dist/reference
	@cp $(JAR_PATH) dist/xl-dist/xl.jar
	@cp .claude/skills/xl-cli/SKILL.md dist/xl-dist/
	@cp .claude/skills/xl-cli/reference/*.md dist/xl-dist/reference/
	@cp scripts/install.sh dist/xl-dist/
	@chmod +x dist/xl-dist/install.sh
	@cd dist && tar czf xl-cli.tar.gz xl-dist
	@echo ""
	@echo "Created dist/xl-cli.tar.gz"
	@du -h dist/xl-cli.tar.gz
	@echo ""
	@echo "Install:"
	@echo "  tar xzf xl-cli.tar.gz && cd xl-dist && ./install.sh"
