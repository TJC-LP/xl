# Output Format Reference

Detailed specifications for `view --format` options.

## Format Summary

| Format | Output | Requires |
|--------|--------|----------|
| markdown | Text table (default) | - |
| json | Structured JSON | - |
| csv | CSV text | - |
| html | HTML table | - |
| svg | SVG vector | - |
| png | PNG image | `--raster-output` (auto fallback chain) |
| jpeg | JPEG image | `--raster-output` (auto fallback chain) |
| webp | WebP image | `--raster-output`, ImageMagick |
| pdf | PDF document | `--raster-output` (auto fallback chain) |

## JSON Schema

```json
{
  "sheet": "Sheet1",
  "range": "A1:C2",
  "rows": [
    {
      "row": 1,
      "cells": [
        {"ref": "A1", "type": "text", "value": "Header", "formatted": "Header"},
        {"ref": "B1", "type": "number", "value": 100, "formatted": "100"},
        {"ref": "C1", "type": "formula", "value": "=A1*2", "formatted": "200"}
      ]
    }
  ]
}
```

**Cell types**: `text`, `number`, `formula`, `bool`, `error`, `datetime`, `empty`

## Raster Options

| Option | Description | Default |
|--------|-------------|---------|
| `--raster-output <path>` | Output file (required) | - |
| `--dpi <n>` | Resolution | 144 |
| `--quality <n>` | JPEG quality 1-100 | 90 |
| `--show-labels` | Add row/column headers | false |
| `--rasterizer <name>` | Force specific rasterizer | auto |

## CSV Options

| Option | Description |
|--------|-------------|
| `--show-labels` | Add column letters as header, row numbers as first column |
| `--formulas` | Show formulas instead of computed values |

## Examples

```bash
# JSON for programmatic parsing
xl -f data.xlsx view A1:D10 --format json

# CSV with labels
xl -f data.xlsx view A1:D10 --format csv --show-labels

# High-quality PNG for visual analysis
xl -f data.xlsx view A1:E20 --format png --raster-output chart.png --show-labels --dpi 300

# JPEG with compression
xl -f data.xlsx view A1:E20 --format jpeg --raster-output chart.jpg --quality 85
```

## Rasterizer Fallback Chain

The CLI automatically tries rasterizers in order until one succeeds:

1. **Batik** (built-in) - Works in JVM mode, not in native image
2. **cairosvg** - Python, very portable (`pip install cairosvg`)
3. **rsvg-convert** - Fast C/Rust binary (`apt install librsvg2-bin`)
4. **resvg** - Best quality Rust (`cargo install resvg`)
5. **ImageMagick** - Widely available, last resort

**Format Support**:
- PNG: All rasterizers
- JPEG: Batik, ImageMagick
- PDF: Batik, cairosvg, rsvg-convert
- WebP: ImageMagick only

**Force Specific Rasterizer**:
```bash
xl -f data.xlsx view A1:E10 --format png --raster-output out.png --rasterizer cairosvg
```

**Installation**:
```bash
# cairosvg (recommended for native image)
pip install cairosvg

# rsvg-convert
brew install librsvg          # macOS
sudo apt install librsvg2-bin # Ubuntu/Debian

# resvg (best quality)
cargo install resvg

# ImageMagick
brew install imagemagick      # macOS
sudo apt install imagemagick  # Ubuntu/Debian
```
