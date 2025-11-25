# XL CLI Installation
# Usage: make install    (installs to ~/.local/bin/xl)
#        make uninstall  (removes ~/.local/bin/xl)

PREFIX ?= $(HOME)/.local
BINDIR ?= $(PREFIX)/bin
JAR_PATH = out/xl-cli/assembly.dest/out.jar

.PHONY: build install uninstall clean help

help:
	@echo "XL CLI Makefile"
	@echo ""
	@echo "Usage:"
	@echo "  make install    Build and install xl to $(BINDIR)"
	@echo "  make uninstall  Remove xl from $(BINDIR)"
	@echo "  make build      Build the fat JAR only"
	@echo "  make clean      Remove build artifacts"
	@echo ""
	@echo "Options:"
	@echo "  PREFIX=<path>   Install prefix (default: ~/.local)"
	@echo "  BINDIR=<path>   Binary directory (default: PREFIX/bin)"

build:
	./mill xl-cli.assembly

install: build
	@mkdir -p $(BINDIR)
	@echo '#!/usr/bin/env bash' > $(BINDIR)/xl
	@echo 'exec java -jar "$(CURDIR)/$(JAR_PATH)" "$$@"' >> $(BINDIR)/xl
	@chmod +x $(BINDIR)/xl
	@echo "Installed xl to $(BINDIR)/xl"
	@echo ""
	@echo "Ensure $(BINDIR) is in your PATH:"
	@echo '  export PATH="$$HOME/.local/bin:$$PATH"'

uninstall:
	@rm -f $(BINDIR)/xl
	@echo "Uninstalled xl from $(BINDIR)/xl"

clean:
	./mill clean
