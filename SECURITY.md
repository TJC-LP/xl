# Security Policy

## Supported Versions

Only the latest minor release line (currently 0.11.x) receives security fixes. Older versions
should upgrade — releases are API-stable within a minor line and artifacts are immutable on
Maven Central.

## Reporting a Vulnerability

Report vulnerabilities **privately** via GitHub Security Advisories:
<https://github.com/TJC-LP/xl/security/advisories/new>

Please do **not** open a public issue or pull request for a suspected vulnerability. You should
receive an acknowledgement within a few business days; coordinated disclosure is appreciated.

## Built-in Hardening

XLSX files are ZIP archives, and the reader treats them as untrusted input: decompression is
capped at **100 MB uncompressed by default** and suspicious compression ratios are rejected with
a structured `SecurityError` — a deliberate zip-bomb defense. The CLI exposes the cap as
`--max-size <MB>` (`0` = unlimited); raise it only for files you trust, or prefer `--stream` for
large files.
