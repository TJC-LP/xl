# XL CLI Installation
# Usage: make install    (installs to ~/.local/bin/xl)
#        make uninstall  (removes ~/.local/bin/xl)

PREFIX ?= $(HOME)/.local
BINDIR ?= $(PREFIX)/bin
SHAREDIR ?= $(PREFIX)/share/xl
JAR_PATH = out/xl-cli/assembly.dest/out.jar

.PHONY: build install uninstall clean help package-skill

help:
	@echo "XL CLI Makefile"
	@echo ""
	@echo "Usage:"
	@echo "  make install       Build and install xl to $(BINDIR)"
	@echo "  make uninstall     Remove xl from $(BINDIR)"
	@echo "  make build         Build the fat JAR only"
	@echo "  make package-skill Create distributable skill zip for Anthropic API"
	@echo "  make clean         Remove build artifacts"
	@echo ""
	@echo "Options:"
	@echo "  PREFIX=<path>   Install prefix (default: ~/.local)"
	@echo "  BINDIR=<path>   Binary directory (default: PREFIX/bin)"

build:
	./mill xl-cli.assembly

install: build
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

# Package skill for Anthropic Skills API
# Creates a distributable zip with JAR, skill files, and install script
package-skill: build
	@echo "Packaging xl-cli skill..."
	@rm -rf dist/xl-skill
	@mkdir -p dist/xl-skill/reference
	@cp $(JAR_PATH) dist/xl-skill/xl.jar
	@cp .claude/skills/xl-cli/SKILL.md dist/xl-skill/
	@cp .claude/skills/xl-cli/reference/*.md dist/xl-skill/reference/
	@cp scripts/install.sh dist/xl-skill/
	@chmod +x dist/xl-skill/install.sh
	@cd dist && zip -r xl-skill.zip xl-skill
	@echo ""
	@echo "Created dist/xl-skill.zip"
	@echo ""
	@echo "Contents:"
	@unzip -l dist/xl-skill.zip
	@echo ""
	@echo "Distribution: dist/xl-skill.zip (includes JAR for local install)"
	@echo ""
	@echo "For Anthropic Skills API (skill files only, <8MB):"
	@echo "  Upload .claude/skills/xl-cli/ directory"
	@echo ""
	@echo "For local/container install:"
	@echo "  unzip xl-skill.zip && cd xl-skill && ./install.sh"
