"""
Task definitions for xl CLI vs Anthropic xlsx skill benchmark.

Each task defines prompts for both approaches to accomplish the same goal,
plus expected answers for correctness grading.
"""

TASKS = [
    {
        "id": "list_sheets",
        "name": "List Sheets",
        "description": "List all sheets in the workbook with basic info",
        "xl_prompt": "List all sheets in the Excel file with their cell counts.",
        "xlsx_prompt": "List all sheets in the Excel file with their cell counts.",
        "expected_answer": """The workbook contains 2 sheets:
1. Sales - Dimension A1:F6 (36 cells)
2. Summary - Dimension A1:B6 (12 cells)""",
    },
    {
        "id": "view_range",
        "name": "View Range",
        "description": "Display a range of cells as a table",
        "xl_prompt": "Show the data in the Sales sheet from A1 to F6.",
        "xlsx_prompt": "Show the data in the Sales sheet from A1 to F6.",
        "expected_answer": """Sales A1:F6 contains product sales data:
- Row 1: Headers (Product, Q1, Q2, Q3, Q4, Total)
- Row 2: Widget A with quarterly values $12,500, $14,200, $13,800, $15,600 and Total formula
- Row 3: Widget B with quarterly values $8,900, $9,100, $8,700, $9,500 and Total formula
- Row 4: Widget C with quarterly values $21,000, $19,500, $22,100, $23,400 and Total formula
- Row 5: Total row with SUM formulas (42400, 42800, 44600, 48500, 178300)
- Row 6: Average row with AVERAGE formulas""",
    },
    {
        "id": "statistics",
        "name": "Calculate Statistics",
        "description": "Compute statistics on a numeric range",
        "xl_prompt": "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
        "xlsx_prompt": "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
        "expected_answer": """Statistics for B2:E4 (12 quarterly values):
- Count: 12
- Sum: 178,300
- Min: 8,700
- Max: 23,400
- Mean/Average: ~14,858""",
    },
    {
        "id": "show_formulas",
        "name": "Show Formulas",
        "description": "Display formulas instead of computed values",
        "xl_prompt": "Show the formulas in the Sales sheet column F (F2:F6).",
        "xlsx_prompt": "Show the formulas in the Sales sheet column F (F2:F6).",
        "expected_answer": """Formulas in Sales column F (F2:F6):
- F2: =SUM(B2:E2)
- F3: =SUM(B3:E3)
- F4: =SUM(B4:E4)
- F5: =SUM(F2:F4)
- F6: =AVERAGE(F2:F4)""",
    },
    {
        "id": "search",
        "name": "Search Content",
        "description": "Find cells containing a specific pattern",
        "xl_prompt": "Search for cells containing 'Widget' in the workbook.",
        "xlsx_prompt": "Search for cells containing 'Widget' in the workbook.",
        "expected_answer": """Found 3 cells containing 'Widget':
- Sales!A2: Widget A
- Sales!A3: Widget B
- Sales!A4: Widget C""",
    },
    {
        "id": "cell_details",
        "name": "Cell Details",
        "description": "Get detailed info about a specific cell including dependencies",
        "xl_prompt": "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
        "xlsx_prompt": "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
        "expected_answer": """Cell F5 in Sales sheet:
- Type: Formula
- Formula: =SUM(F2:F4)
- Dependencies: F2, F3, F4
- F5 is the grand total summing the individual product totals""",
    },
    {
        "id": "cross_sheet_ref",
        "name": "Cross-Sheet Reference",
        "description": "Understand cross-sheet formula references",
        "xl_prompt": "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
        "xlsx_prompt": "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
        "expected_answer": """Cell B4 in Summary sheet:
- Formula: =Sales!F5
- This references the grand total from the Sales sheet
- Value: 178,300 (sum of all product totals)""",
    },
]

# Tasks that are specifically designed for larger files
LARGE_FILE_TASKS = [
    {
        "id": "large_view",
        "name": "View Large Range",
        "description": "Display first 20 rows of a large dataset",
        "xl_prompt": "Show the first 20 rows of data in the Data sheet.",
        "xlsx_prompt": "Show the first 20 rows of data in the Data sheet.",
    },
    {
        "id": "large_stats",
        "name": "Large Statistics",
        "description": "Compute statistics on a large range",
        "xl_prompt": "Calculate statistics for column B (B2:B1001) in the Data sheet.",
        "xlsx_prompt": "Calculate statistics for column B (B2:B1001) in the Data sheet.",
    },
    {
        "id": "large_search",
        "name": "Large Search",
        "description": "Search in a large file",
        "xl_prompt": "Find the first 5 cells containing 'Row 500' in the workbook.",
        "xlsx_prompt": "Find the first 5 cells containing 'Row 500' in the workbook.",
    },
]
