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
| png | PNG image | ImageMagick, `--raster-output` |
| jpeg | JPEG image | ImageMagick, `--raster-output` |
| webp | WebP image | ImageMagick, `--raster-output` |
| pdf | PDF document | ImageMagick, `--raster-output` |

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

## ImageMagick Installation

Required for PNG/JPEG/WebP/PDF export:

```bash
# macOS
brew install imagemagick

# Ubuntu/Debian
sudo apt-get install imagemagick
```
