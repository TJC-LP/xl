# XL Examples

Standalone demo scripts showcasing XL library features.

## Running Examples

These examples use [scala-cli](https://scala-cli.virtuslab.org/) for easy standalone execution.

### Prerequisites

1. **Install scala-cli** (if not already installed):
   ```bash
   brew install Virtuslab/scala-cli/scala-cli  # macOS
   # or: curl -sSLf https://scala-cli.virtuslab.org/get | sh
   ```

2. **Publish xl-core locally**:
   ```bash
   ./mill xl-core.publishLocal
   ```

### Running demo.sc

```bash
scala-cli run examples/demo.sc
```

The script showcases:
- Compile-time validated literals (`ref"A1"`, `money"$1,234.56"`)
- Creating workbooks and sheets
- Addressing operations (shift, intersects, contains)
- Sheet-qualified references (`ref"Sales!A1"`)
- Formatted literals (money, percent, date)
- Rich text formatting

### Running easy-mode-demo.sc

```bash
scala-cli run examples/easy-mode-demo.sc
```

The script showcases **Easy Mode API** (LLM-friendly ergonomics):
- String-based cell references (`.put("A1", value)`)
- Template-first styling (`.style("A1:B1", style)`)
- Inline styling (`.put("A1", value, style)`)
- Safe lookups (`.cell("A1")` returns `Option`)
- Rich text formatting (`"Error: ".red.bold + "Fix!"`)
- Simplified IO (`Excel.read/write/modify`)
- Structured exception handling (`XLException` wraps `XLError`)

## Adding New Examples

Create a new `.sc` script file with scala-cli directives:

```scala
//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using repository ivy2Local

import com.tjclp.xl.*

@main def myExample(): Unit =
  // Your code here
```

Run with: `scala-cli run examples/my-example.sc`
