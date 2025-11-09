
# Drawings & Graphics — Anchors, Pictures, Vector Shapes

## Anchors
- `TwoCell(fromRow, fromCol, fromDx, fromDy) → (toRow, toCol, toDx, toDy)`
- `OneCell(atRow, atCol, dx, dy, widthEmu, heightEmu)`
- **Normalization:** ensure `(from ≤ to)` lexicographically; clamp negative EMUs to 0.

## Pictures
- `ImageId = sha256(bytes)` → dedup single media part.
- Alt text, rotation degrees, lock aspect ratio modeled as pure fields.

## Shapes
- `Rect/Ellipse/Line/Polygon` + `Stroke`, `SolidFill | NoFill`.
- Pure affine transforms with composition law; identity is no‑op.

## OOXML mapping
- Drawings reside in `xl/drawings/drawing#.xml` with anchors `xdr:twoCellAnchor`/`xdr:oneCellAnchor`.
- Media stored under `/xl/media/*.png|*.jpeg`; relations via `drawing#.xml.rels` and `workbook.xml.rels`.
