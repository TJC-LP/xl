
# FAQ

**Q: Why not reuse POI objects?**  
A: We want immutability, type safety, and compile‑time guarantees; POI’s design makes that difficult.

**Q: Will charts look identical to Excel’s?**  
A: Yes for supported types: the mapping targets Excel’s ChartML directly with theme‑aware colors and deterministic ordering.

**Q: Do you support `.xls`?**  
A: Not initially. Focus is `.xlsx`/`.xlsm`. BIFF can be a separate module later.
