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
      xlPrompt =
        "List all sheets in the Excel file. For each sheet, report its name, dimension (used range), and cell count.",
      xlsxPrompt =
        "List all sheets in the Excel file. For each sheet, report its name, dimension (used range), and cell count.",
      expectedAnswer = Some("""The workbook contains 2 sheets:
1. Sales - Dimension A1:F6 (36 cells)
2. Summary - Dimension A1:B6 (12 cells)""")
    ),
    TokenTask(
      id = "view_range",
      name = "View Range",
      description = "Display a range of cells as a table",
      xlPrompt =
        "Display the contents of Sales sheet range A1:F6. Show all cell values including the specific dollar amounts and formulas.",
      xlsxPrompt =
        "Display the contents of Sales sheet range A1:F6. Show all cell values including the specific dollar amounts and formulas.",
      expectedAnswer = Some("""Sales A1:F6 contents:
| Row | A        | B          | C          | D          | E          | F           |
|-----|----------|------------|------------|------------|------------|-------------|
| 1   | Product  | Q1         | Q2         | Q3         | Q4         | Total       |
| 2   | Widget A | $12,500    | $14,200    | $13,800    | $15,600    | =SUM(B2:E2) |
| 3   | Widget B | $8,900     | $9,100     | $8,700     | $9,500     | =SUM(B3:E3) |
| 4   | Widget C | $21,000    | $19,500    | $22,100    | $23,400    | =SUM(B4:E4) |
| 5   | Total    | =SUM(B2:B4)| =SUM(C2:C4)| =SUM(D2:D4)| =SUM(E2:E4)| =SUM(F2:F4) |
| 6   | Average  | =AVERAGE   | =AVERAGE   | =AVERAGE   | =AVERAGE   | =AVERAGE    |""")
    ),
    TokenTask(
      id = "statistics",
      name = "Calculate Statistics",
      description = "Compute statistics on a numeric range",
      xlPrompt =
        "Calculate statistics for the quarterly data in Sales B2:E4. Report the count, sum, min, max, and mean with their numeric values.",
      xlsxPrompt =
        "Calculate statistics for the quarterly data in Sales B2:E4. Report the count, sum, min, max, and mean with their numeric values.",
      expectedAnswer = Some("""Statistics for Sales B2:E4:
- Count: 12 values
- Sum: $178,300
- Min: $8,700
- Max: $23,400
- Mean: $14,858.33""")
    ),
    TokenTask(
      id = "show_formulas",
      name = "Show Formulas",
      description = "Display formulas instead of computed values",
      xlPrompt =
        "List all formulas in Sales sheet column F (cells F2 through F6). Show each cell reference and its exact formula.",
      xlsxPrompt =
        "List all formulas in Sales sheet column F (cells F2 through F6). Show each cell reference and its exact formula.",
      expectedAnswer = Some("""Formulas in Sales column F:
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
      xlPrompt =
        "Search for all cells containing 'Widget' in the workbook. List each match with its sheet name, cell reference, and cell value.",
      xlsxPrompt =
        "Search for all cells containing 'Widget' in the workbook. List each match with its sheet name, cell reference, and cell value.",
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
        "Get details about cell F5 in the Sales sheet. Report: cell type, formula (if any), cell dependencies, and what the cell represents.",
      xlsxPrompt =
        "Get details about cell F5 in the Sales sheet. Report: cell type, formula (if any), cell dependencies, and what the cell represents.",
      expectedAnswer = Some("""Cell Sales!F5 details:
- Type: Formula
- Formula: =SUM(F2:F4)
- Dependencies: F2, F3, F4 (the individual product totals)
- Purpose: Grand total summing all product totals""")
    ),
    TokenTask(
      id = "cross_sheet_ref",
      name = "Cross-Sheet Reference",
      description = "Understand cross-sheet formula references",
      xlPrompt =
        "Examine cell B4 in the Summary sheet. Report its formula, what sheet/cell it references, and its computed value.",
      xlsxPrompt =
        "Examine cell B4 in the Summary sheet. Report its formula, what sheet/cell it references, and its computed value.",
      expectedAnswer = Some("""Cell Summary!B4:
- Formula: =Sales!F5
- References: Cell F5 in the Sales sheet (grand total)
- Value: $178,300""")
    )
  )

  /** Large file tasks (require large_sample.xlsx) */
  val largeFileTasks: List[TokenTask] = List(
    TokenTask(
      id = "large_view",
      name = "View Large Range",
      description = "Display first 20 rows of a large dataset",
      xlPrompt =
        "Display the first 20 rows of data in the Data sheet as a table showing all columns and their values.",
      xlsxPrompt =
        "Display the first 20 rows of data in the Data sheet as a table showing all columns and their values.",
      expectedAnswer = None
    ),
    TokenTask(
      id = "large_stats",
      name = "Large Statistics",
      description = "Compute statistics on a large range",
      xlPrompt =
        "Calculate statistics for column B (B2:B1001) in the Data sheet. Report count, sum, min, max, and mean with their numeric values.",
      xlsxPrompt =
        "Calculate statistics for column B (B2:B1001) in the Data sheet. Report count, sum, min, max, and mean with their numeric values.",
      expectedAnswer = None
    ),
    TokenTask(
      id = "large_search",
      name = "Large Search",
      description = "Search in a large file",
      xlPrompt =
        "Find the first 5 cells containing 'Row 500' in the workbook. List each match with sheet name, cell reference, and cell value.",
      xlsxPrompt =
        "Find the first 5 cells containing 'Row 500' in the workbook. List each match with sheet name, cell reference, and cell value.",
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
