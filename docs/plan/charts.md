
# Charts — Typed Grammar → Deterministic ChartML

## Spec
- Marks: `Column|Bar`(clustered/stacked), `Line`(smooth/markers), `Area`(stacked), `Scatter`, `Pie`.
- Axis: Category/Value/Time; Scales: Linear/Log10/Time.
- Series: `(mark, encoding{x,y,color?,size?}, name?)`

## Semantics
- **Normalization:** series order → by `name` stable; color palette resolved via theme.
- **Stacking:** grouped by X; y-values summed; gaps handled as zeros or missing per policy.
- **Legend & titles:** pure values; printed deterministically.

## Mapping
- `xl/charts/chart#.xml`: `c:barChart|lineChart|...`, `c:plotArea`, `c:legend`, axes.
- Data ranges emit `c:numRef` / `c:strRef` with `f` formulas referencing sheet ranges.
