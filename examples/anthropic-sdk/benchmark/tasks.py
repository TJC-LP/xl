"""
Task definitions for xl CLI vs Anthropic xlsx skill benchmark.

Each task defines prompts for both approaches to accomplish the same goal.
"""

TASKS = [
    {
        "id": "list_sheets",
        "name": "List Sheets",
        "description": "List all sheets in the workbook with basic info",
        "xl_prompt": "List all sheets in the Excel file with their cell counts.",
        "xlsx_prompt": "List all sheets in the Excel file with their cell counts.",
    },
    {
        "id": "view_range",
        "name": "View Range",
        "description": "Display a range of cells as a table",
        "xl_prompt": "Show the data in the Sales sheet from A1 to F6.",
        "xlsx_prompt": "Show the data in the Sales sheet from A1 to F6.",
    },
    {
        "id": "statistics",
        "name": "Calculate Statistics",
        "description": "Compute statistics on a numeric range",
        "xl_prompt": "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
        "xlsx_prompt": "Calculate statistics (count, sum, min, max, mean) for the quarterly data in Sales B2:E4.",
    },
    {
        "id": "show_formulas",
        "name": "Show Formulas",
        "description": "Display formulas instead of computed values",
        "xl_prompt": "Show the formulas in the Sales sheet column F (F2:F6).",
        "xlsx_prompt": "Show the formulas in the Sales sheet column F (F2:F6).",
    },
    {
        "id": "search",
        "name": "Search Content",
        "description": "Find cells containing a specific pattern",
        "xl_prompt": "Search for cells containing 'Widget' in the workbook.",
        "xlsx_prompt": "Search for cells containing 'Widget' in the workbook.",
    },
    {
        "id": "cell_details",
        "name": "Cell Details",
        "description": "Get detailed info about a specific cell including dependencies",
        "xl_prompt": "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
        "xlsx_prompt": "Get details about cell F5 in the Sales sheet including its formula and what it depends on.",
    },
    {
        "id": "cross_sheet_ref",
        "name": "Cross-Sheet Reference",
        "description": "Understand cross-sheet formula references",
        "xl_prompt": "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
        "xlsx_prompt": "What does cell B4 in the Summary sheet reference? Show the formula and its value.",
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
