package com.tjclp.xl

/**
 * Excel IO export for convenient access via `import com.tjclp.xl.*`.
 *
 * When xl-cats-effect is on the classpath, this makes `Excel` available alongside the core API:
 * {{{
 * import com.tjclp.xl.*
 * import com.tjclp.xl.unsafe.*
 *
 * val wb = Excel.read("data.xlsx")
 * Excel.write(wb, "output.xlsx")
 * }}}
 */
val Excel: io.Excel.type = io.Excel
