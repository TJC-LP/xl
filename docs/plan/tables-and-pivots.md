
# Tables & Pivots — Pure Specs, Deterministic Writers

## Tables
- `TableSpec(range, style, showFilterButtons, totalsRow)`
- Banding computed by parity (rows, columns, or both); independent of content.

## Pivots
- `PivotSpec(sourceRange or cache, rows: List[Field], cols: List[Field], values: List[Aggregation])`
- Aggregations: `Sum/Avg/Count/Min/Max` over numeric fields.

## Slicers
- `SlicerSpec(field, selectionSet)` → pure filter state.
- Serialization to `slicer*.xml` total; interactions adjust pivot cache query.
