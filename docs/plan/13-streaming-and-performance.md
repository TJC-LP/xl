
# Streaming & Performance — Targets, Techniques, Patterns

## Targets
- Parse & write 100k+ rows/sheet without OOM (1–2 GB heap).
- ≤ 2× POI on typical read/transform/write; faster on streaming transforms.
- Core < 5 MB.

## Techniques
- **fs2‑data‑xml** pull parsing; no DOM for big sheets.
- Persistent maps/vectors with structural sharing for edits.
- SST LRU cache + optional spill to temp file.
- Inline givens for codecs to avoid dictionary allocs.
- Macro sugar compiles away (no runtime parsing).

## Patterns
- Chunking: `rowsStream.chunkN(2048)` + backpressure.
- Filter‑map‑fold pipelines fuse with the JVM JIT; optional stream fusion macros available.
