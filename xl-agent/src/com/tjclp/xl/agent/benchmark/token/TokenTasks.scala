package com.tjclp.xl.agent.benchmark.token

/**
 * Task definitions for the token efficiency benchmark.
 *
 * These tasks are designed to compare the token efficiency of the xl CLI (custom skill) vs
 * Anthropic's built-in xlsx skill (openpyxl).
 *
 * Each task has parallel prompts for both approaches and an expected answer for correctness
 * grading.
 */
object TokenTasks:

  /** Standard tasks on sample.xlsx */
  val standardTasks: List[TokenTask] = List(
    TokenTask(
      id = "list_sheets",
      name = "List Sheets",
      description = "List all sheets in the workbook with basic info",
      xlPrompt = "List all sheets in the Excel file with their cell counts.",
      xlsxPrompt = "List all sheets in the Excel file with their cell counts.",
      expectedAnswer = Some("""The workbook contains 2 sheets:
1. Sales - Dimension A1:F6 (36 cells)
2. Summary - Dimension A1:B6 (12 cells)""")
    ),
    TokenTask(
      id = "view_range",
      name = "View Range",
      description = "Display a range of cells as a table",
      xlPrompt = "Show the data in the Sales sheet from A1 to F6.",
      xlsxPrompt = "Show the data in the Sales sheet from A1 to F6.",
      expectedAnswer = Some("""Sales A1:F6 contains product sales data:
- Row 1: Headers (Product, Q1, Q2, Q3, Q4, Total)
- Row 2: Widget A with quarterly values $12,500, $14,200, $13,800, $15,600 and Total formula
- Row 3: Widget B with quarterly values $8,900, $9,100, $8,700, $9,500 and Total formula
- Row 4: Widget C with quarterly values $21,000, $19,500, $22,100, $23,400 and Total formula
- Row 5: Total row with SUM formulas (42400, 42800, 44600, 48500, 178300)
- Row 6: Average row with AVERAGE formulas""")
    ),
    TokenTask(
      id = "statistics",
      name = "Calculate Statistics",
      description = "Compute statistics on a numeric range",
      xlPrompt =
        "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
      xlsxPrompt =
        "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
      expectedAnswer = Some("""Statistics for B2:E4 (12 quarterly values):
- Count: 12
- Sum: 178,300
- Min: 8,700
- Max: 23,400
- Mean/Average: ~14,858""")
    ),
    TokenTask(
      id = "show_formulas",
      name = "Show Formulas",
      description = "Display formulas instead of computed values",
      xlPrompt = "Show the formulas in the Sales sheet column F (F2:F6).",
      xlsxPrompt = "Show the formulas in the Sales sheet column F (F2:F6).",
      expectedAnswer = Some("""Formulas in Sales column F (F2:F6):
- F2: =SUM(B2:E2)
- F3: =SUM(B3:E3)
- F4: =SUM(B4:E4)
- F5: =SUM(F2:F4)
- F6: =AVERAGE(F2:F4)""")
    ),
    TokenTask(
      id = "search",
      name = "Search Content",
      description = "Find cells containing a specific pattern",
      xlPrompt = "Search for cells containing 'Widget' in the workbook.",
      xlsxPrompt = "Search for cells containing 'Widget' in the workbook.",
      expectedAnswer = Some("""Found 3 cells containing 'Widget':
- Sales!A2: Widget A
- Sales!A3: Widget B
- Sales!A4: Widget C""")
    ),
    TokenTask(
      id = "cell_details",
      name = "Cell Details",
      description = "Get detailed info about a specific cell including dependencies",
      xlPrompt =
        "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
      xlsxPrompt =
        "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
      expectedAnswer = Some("""Cell F5 in Sales sheet:
- Type: Formula
- Formula: =SUM(F2:F4)
- Dependencies: F2, F3, F4
- F5 is the grand total summing the individual product totals""")
    ),
    TokenTask(
      id = "cross_sheet_ref",
      name = "Cross-Sheet Reference",
      description = "Understand cross-sheet formula references",
      xlPrompt =
        "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
      xlsxPrompt =
        "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
      expectedAnswer = Some("""Cell B4 in Summary sheet:
- Formula: =Sales!F5
- This references the grand total from the Sales sheet
- Value: 178,300 (sum of all product totals)""")
    )
  )

  /** Large file tasks (require large_sample.xlsx) */
  val largeFileTasks: List[TokenTask] = List(
    TokenTask(
      id = "large_view",
      name = "View Large Range",
      description = "Display first 20 rows of a large dataset",
      xlPrompt = "Show the first 20 rows of data in the Data sheet.",
      xlsxPrompt = "Show the first 20 rows of data in the Data sheet.",
      expectedAnswer = None
    ),
    TokenTask(
      id = "large_stats",
      name = "Large Statistics",
      description = "Compute statistics on a large range",
      xlPrompt = "Calculate statistics for column B (B2:B1001) in the Data sheet.",
      xlsxPrompt = "Calculate statistics for column B (B2:B1001) in the Data sheet.",
      expectedAnswer = None
    ),
    TokenTask(
      id = "large_search",
      name = "Large Search",
      description = "Search in a large file",
      xlPrompt = "Find the first 5 cells containing 'Row 500' in the workbook.",
      xlsxPrompt = "Find the first 5 cells containing 'Row 500' in the workbook.",
      expectedAnswer = None
    )
  )

  /** Get all tasks, optionally including large file tasks */
  def allTasks(includeLarge: Boolean = false): List[TokenTask] =
    if includeLarge then standardTasks ++ largeFileTasks
    else standardTasks

  /** Find a task by ID */
  def findById(id: String): Option[TokenTask] =
    (standardTasks ++ largeFileTasks).find(_.id == id)

  /** Task IDs */
  val standardTaskIds: List[String] = standardTasks.map(_.id)
  val largeTaskIds: List[String] = largeFileTasks.map(_.id)
