package spreadsheetbench

import zio.*
import zio.json.*
import java.nio.file.{Files, Path, Paths}
import org.apache.commons.compress.archivers.tar.TarArchiveInputStream
import org.apache.commons.compress.compressors.gzip.GzipCompressorInputStream
import java.io.{BufferedInputStream, FileInputStream}
import scala.jdk.CollectionConverters.*

object TaskLoader:

  /** Load benchmark tasks from a dataset */
  def load(config: BenchConfig): Task[List[BenchTask]] =
    for
      dataDir <- ensureExtracted(config.dataDir, config.dataset)
      jsonPath = dataDir.resolve("dataset.json")
      content <- ZIO.attemptBlocking(Files.readString(jsonPath))
      tasks   <- ZIO.fromEither(content.fromJson[List[DatasetTask]])
                   .mapError(e => new Exception(s"JSON parse error: $e"))

      // Apply filters in order
      filtered = applyFilters(tasks, config)

    yield filtered.map(toBenchTask(dataDir, _))

  /** Apply all configured filters to the task list */
  private def applyFilters(tasks: List[DatasetTask], config: BenchConfig): List[DatasetTask] =
    var result = tasks

    // Filter by specific task IDs if provided
    config.taskIds match
      case Some(ids) =>
        val idSet = ids.toSet
        result = result.filter(t => idSet.contains(extractId(t.id)))
      case None => ()

    // Remove skipped task IDs
    if config.skipIds.nonEmpty then
      result = result.filterNot(t => config.skipIds.contains(extractId(t.id)))

    // Filter by category if provided
    config.category match
      case Some(cat) =>
        result = result.filter(_.instruction_type.toLowerCase.contains(cat.toLowerCase))
      case None => ()

    // Filter VBA tasks (always, unless specific task IDs were requested)
    if config.taskIds.isEmpty then
      result = result.filterNot(isVbaTask)

    // Apply limit
    config.limit match
      case Some(n) => result.take(n)
      case None    => result

  /** Extract tarball if needed, return path to extracted data */
  private def ensureExtracted(dataDir: Path, dataset: String): Task[Path] =
    val tarPath = dataDir.resolve(s"$dataset.tar.gz")
    val extractedPath = dataDir.resolve(dataset)

    ZIO.ifZIO(ZIO.attemptBlocking(Files.exists(extractedPath) && Files.isDirectory(extractedPath)))(
      onTrue = ZIO.logInfo(s"Dataset already extracted at $extractedPath") *> ZIO.succeed(extractedPath),
      onFalse = for
        _ <- ZIO.logInfo(s"Extracting $tarPath to $extractedPath")
        _ <- extractTarGz(tarPath, dataDir)
        _ <- ZIO.logInfo(s"Extraction complete")
      yield extractedPath
    )

  /** Extract a .tar.gz file */
  private def extractTarGz(tarPath: Path, destDir: Path): Task[Unit] =
    ZIO.attemptBlocking {
      val fis = new FileInputStream(tarPath.toFile)
      val bis = new BufferedInputStream(fis)
      val gzis = new GzipCompressorInputStream(bis)
      val tais = new TarArchiveInputStream(gzis)

      try
        var entry = tais.getNextEntry
        while entry != null do
          val destPath = destDir.resolve(entry.getName)
          if entry.isDirectory then
            Files.createDirectories(destPath)
          else
            Files.createDirectories(destPath.getParent)
            Files.copy(tais, destPath)
          entry = tais.getNextEntry
      finally
        tais.close()
        gzis.close()
        bis.close()
        fis.close()
    }

  /** Check if a task is VBA-based */
  private def isVbaTask(t: DatasetTask): Boolean =
    val lower = t.instruction.toLowerCase
    lower.contains("vba") || lower.contains("macro")

  /** Extract string ID from JSON (handles both Int and String) */
  private def extractId(json: zio.json.ast.Json): String =
    json match
      case zio.json.ast.Json.Str(s) => s
      case zio.json.ast.Json.Num(n) => n.intValue.toString
      case other                    => other.toString

  /** Convert DatasetTask to BenchTask with test cases */
  private def toBenchTask(dataDir: Path, t: DatasetTask): BenchTask =
    val taskId = extractId(t.id)
    val spreadsheetDir = dataDir.resolve(t.spreadsheet_path)

    BenchTask(
      id = taskId,
      instruction = t.instruction,
      instructionType = t.instruction_type,
      answerPosition = t.answer_position,
      testCases = (1 to 3).toList.map { caseNum =>
        TestCase(
          caseNum = caseNum,
          inputPath = spreadsheetDir.resolve(s"${caseNum}_${taskId}_input.xlsx"),
          answerPath = spreadsheetDir.resolve(s"${caseNum}_${taskId}_answer.xlsx")
        )
      }
    )
