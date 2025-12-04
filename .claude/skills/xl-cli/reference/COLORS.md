# Color Reference

Colors for `--bg`, `--fg`, and `--border-color` options.

## Color Formats

| Format | Syntax | Example |
|--------|--------|---------|
| Named | `colorname` | `--bg red` |
| Hex | `#RRGGBB` | `--bg "#FF6600"` |
| RGB | `rgb(r,g,b)` | `--bg "rgb(100,150,200)"` |

## Named Colors

| Color | RGB Value |
|-------|-----------|
| black | 0, 0, 0 |
| white | 255, 255, 255 |
| red | 255, 0, 0 |
| green | 0, 128, 0 |
| blue | 0, 0, 255 |
| yellow | 255, 255, 0 |
| orange | 255, 165, 0 |
| purple | 128, 0, 128 |
| pink | 255, 192, 203 |
| cyan | 0, 255, 255 |
| magenta | 255, 0, 255 |
| gray / grey | 128, 128, 128 |
| lightgray / lightgrey | 211, 211, 211 |
| darkgray / darkgrey | 169, 169, 169 |
| brown | 139, 69, 19 |
| navy | 0, 0, 128 |
| teal | 0, 128, 128 |
| olive | 128, 128, 0 |
| maroon | 128, 0, 0 |
| silver | 192, 192, 192 |
| gold | 255, 215, 0 |
| lime | 0, 255, 0 |

## Examples

```bash
# Named color
xl -f data.xlsx -o out.xlsx style A1:D1 --bg navy --fg white

# Hex color
xl -f data.xlsx -o out.xlsx style A1:D1 --bg "#FF6600"

# RGB color
xl -f data.xlsx -o out.xlsx style A1:D1 --bg "rgb(100, 150, 200)"

# Border color
xl -f data.xlsx -o out.xlsx style A1:D1 --border thin --border-color red
```
