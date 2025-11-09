
# Security & Hardening

## Threat model
- Untrusted `.xlsx` inputs; adversarial zip bombs; malformed XML; formula injection vectors.

## Controls
- Zip entry count/size/ratio caps; path traversal prevention.
- XML parser limits: entity expansion disabled; depth limits; attribute count limits.
- Formula text sanitizer for untrusted text writes (opt‑in allowlist).

## Guidance
- Never evaluate macros; do not embed active content; prefer typed formulas.
