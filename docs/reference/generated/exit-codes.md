<!-- GENERATED from `xl schema --json` by DocsGenSpec; do not edit by hand.
     A diff here is a contract change. Regenerate: XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec -->

# Exit codes

Every `xl` invocation ends with one of these four codes; `xl --help` prints the same table.
Exit 1 is reserved for "completed with findings or a failed gate" and is never a failure.
Branch on the exit code and the `code:` line (or `error.code` under `--json`), never on
message text.

| exit | meaning |
| --- | --- |
| `0` | ok |
| `1` | completed with findings or a failed gate (diff differs, lint findings, audit --fail-on-findings, --strict) — never a failure; with -o the file is written, with -i the input is left untouched |
| `2` | usage — the command line is wrong; nothing read, nothing written |
| `3` | failed — the operation could not complete; nothing written |
