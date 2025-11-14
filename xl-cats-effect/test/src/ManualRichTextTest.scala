import cats.effect.{IO, IOApp}
import com.tjclp.xl.api.*
import com.tjclp.xl.io.Excel
import com.tjclp.xl.richtext.RichText.{*, given}
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheet.syntax.*
import java.nio.file.Paths

/** Manual test to verify rich text formatting in Excel.
  *
  * Run this test and open the generated rich-text-demo.xlsx in Excel to verify that:
  * 1. Bold/italic/underline formatting appears correctly
  * 2. Colors (red, green, blue) appear correctly
  * 3. Font sizes work
  * 4. Multiple runs within a cell display properly
  *
  * To run:
  * {{{
  * ./mill xl-cats-effect.test.runMain ManualRichTextTest
  * }}}
  */
object ManualRichTextTest extends IOApp.Simple:

  def run: IO[Unit] =
    val excel = Excel.forIO
    val outputPath = Paths.get("rich-text-demo.xlsx")

    // Build a financial report with rich text formatting
    val report = Sheet("Q1 Performance Report").getOrElse(sys.error("Failed to create sheet"))
      .put(
        // Title with large bold text
        ref"A1" -> ("Q1 2025 ".size(18.0).bold + "Performance Report".size(18.0).italic),

        // Section headers
        ref"A3" -> "Metric".bold,
        ref"B3" -> "Change".bold,

        // Revenue (positive - green)
        ref"A4" -> "Revenue",
        ref"B4" -> ("+12.5%".green.bold + " (strong growth)"),

        // Expenses (negative - red)
        ref"A5" -> "Expenses",
        ref"B5" -> ("+8.2%".red.bold + " (cost increase)"),

        // Profit (positive - green)
        ref"A6" -> "Net Profit",
        ref"B6" -> ("+4.3%".green.bold + " (improved margin)"),

        // Warning message
        ref"A8" -> ("Warning: ".red.bold.underline + "Review quarterly targets".italic),

        // Mixed formatting example
        ref"A10" -> ("This cell has ".fontFamily("Calibri") +
          "bold ".bold +
          "italic ".italic +
          "underline ".underline +
          "and ".size(14.0) +
          "colored ".blue.bold +
          "text!".red.size(16.0).bold),

        // Plain text for comparison
        ref"A12" -> "This is plain text (no formatting)",

        // Numbers for context
        ref"A14" -> "Revenue ($M):",
        ref"B14" -> BigDecimal("1250.50")
      )

    val workbook = Workbook(Vector(report))

    for
      _ <- excel.write(workbook, outputPath)
      _ <- IO.println(s"✅ Wrote $outputPath")
      _ <- IO.println("\n📝 Open rich-text-demo.xlsx in Excel to verify:")
      _ <- IO.println("   1. Cell A1: Large title with bold + italic")
      _ <- IO.println("   2. Cells B4-B6: Green/red colors with bold")
      _ <- IO.println("   3. Cell A8: Red bold underline + italic")
      _ <- IO.println("   4. Cell A10: Multiple font styles and sizes")
      _ <- IO.println("   5. Cell B14: Decimal number format")

      // Export to HTML
      htmlTable = report.toHtml(ref"A1:B14")
      _ <- IO.println(s"\n📊 HTML Export:\n$htmlTable")
      _ <- IO.println("\n✅ You can paste the HTML into a web page to see the formatting")
    yield ()
