package com.tjclp.xl

/**
 * Package-level forwarders for the recalculating writes (GH-360 `Excel.writeRecalculated`, GH-589
 * `Excel.writeChecked`), so `import com.tjclp.xl.{*, given}` sees the same `Excel` facade the
 * scripting prelude does: a formula model built under the base import has the cache-completing
 * write at hand instead of only `Excel.write`. The prelude re-exports the object separately; a file
 * never imports both (see `scripting`).
 */
export io.ExcelRecalc.*
