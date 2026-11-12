# Workflow action pin policy

Run `python3 scripts/check-action-pins.py` from the checkout. CI checks every
`.yml` and `.yaml` file directly under `.github/workflows`.

Remote `uses:` values require a full lowercase 40-hex commit and a release-version
comment. Local `./` references and Docker image references pinned by SHA-256 are
accepted. Action upgrades, tag provenance and action-internal downloads are
separate concerns.

```yaml
steps:
  - uses: owner/action@0123456789012345678901234567890123456789 # v1.2.3
```

The checker enforces a restricted repository format, rather than parsing all YAML.
Use spaces for indentation and sequence-dash separation, and block mappings
with plain identifier keys or simple quoted keys without escapes. The `uses` key itself must be unquoted and written `uses:`. Scalar values
may be plain text, complete single-line quoted strings, or `|`/`>` script blocks
with ordinary indentation/chomping modifiers. Comments, scalar payloads and quoted
path-filter list items are not action declarations. One leading UTF-8 BOM is
accepted without hiding the first mapping.

Safe inline exceptions are empty maps, complete plain-scalar lists such as
`branches: [main, release]`, and one-line option maps under `env`, `with` or
`permissions`, such as `with: { fetch-depth: 0 }`. Option-map keys and values are
simple ASCII identifier/value atoms (letters, digits, `_`, `.`, `/`, `*`, `+`, `-`);
Keys may not be `uses`, and values cannot begin with alias/property indicators. Use block form for quoted values, expressions, comments
inside a map, nested collections or other punctuation.

Other structural forms fail the check, including document markers/directives,
anchors, aliases, tags, merge/explicit/escaped keys, nonempty general flow maps,
multiline flow collections and multiline quoted strings. Use the supported block
forms instead. Nonempty noncomment lines outside a script block cannot silently
fall through as unrecognized syntax. A `uses:` value is checked before any empty
map/list/scalar exception.

This policy does not validate the complete GitHub workflow schema or prove that
any synthetic rejected example would execute on GitHub. Literal or folded script
content is deliberately not parsed as workflow structure. YAML indentation and
BOM handling follow the [YAML 1.2.2 specification](https://yaml.org/spec/1.2.2/);
the allowed repository forms are intentionally narrower.
