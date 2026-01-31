package com.tjclp.xl.agent.benchmark.tracing

import cats.effect.{Fiber, IO}
import cats.effect.std.Queue
import cats.effect.syntax.all.*
import cats.syntax.all.*
import com.tjclp.xl.agent.AgentEvent

import java.util.concurrent.atomic.AtomicReference

/** Real-time colored console output for streaming conversations */
object StreamingConsole:

  case class StreamContext(taskId: String, skillName: String, caseNum: Int)

  private sealed trait StreamItem
  private object StreamItem:
    case class CaseStart(ctx: StreamContext) extends StreamItem
    case class CaseComplete(ctx: StreamContext, passed: Boolean, durationMs: Long)
        extends StreamItem
    case class Event(ctx: StreamContext, traced: TracedEvent) extends StreamItem
    case object Shutdown extends StreamItem

  private case class State(
    queue: Queue[IO, StreamItem],
    fiber: Fiber[IO, Throwable, Unit]
  )

  private val stateRef = new AtomicReference[Option[State]](None)

  private val DefaultQueueSize = 2048
  private val DefaultMaxTextChars = 4096

  // ANSI color codes
  private val Reset = "\u001b[0m"
  private val Cyan = "\u001b[36m"
  private val Yellow = "\u001b[33m"
  private val Green = "\u001b[32m"
  private val Red = "\u001b[31m"
  private val Gray = "\u001b[90m"
  private val Bold = "\u001b[1m"

  def withStreaming[A](enabled: Boolean)(fa: IO[A]): IO[A] =
    if !enabled then fa
    else start(DefaultQueueSize, DefaultMaxTextChars) *> fa.guarantee(stop())

  def start(queueSize: Int, maxTextChars: Int): IO[Unit] =
    IO(stateRef.get).flatMap {
      case Some(_) => IO.unit
      case None =>
        for
          queue <- Queue.bounded[IO, StreamItem](queueSize)
          fiber <- drain(queue, maxTextChars).start
          state = State(queue, fiber)
          set = stateRef.compareAndSet(None, Some(state))
          _ <- if set then IO.unit else fiber.cancel
        yield ()
    }

  def stop(): IO[Unit] =
    IO(stateRef.getAndSet(None)).flatMap {
      case Some(state) =>
        state.queue.offer(StreamItem.Shutdown) *> state.fiber.join.void.handleError(_ => ())
      case None => IO.unit
    }

  def enqueueCaseStart(ctx: StreamContext): IO[Unit] =
    enqueue(StreamItem.CaseStart(ctx))

  def enqueueCaseComplete(ctx: StreamContext, passed: Boolean, durationMs: Long): IO[Unit] =
    enqueue(StreamItem.CaseComplete(ctx, passed, durationMs))

  def enqueueEvent(ctx: StreamContext, traced: TracedEvent): IO[Unit] =
    enqueue(StreamItem.Event(ctx, traced))

  private def enqueue(item: StreamItem): IO[Unit] =
    IO(stateRef.get).flatMap {
      case Some(state) => state.queue.tryOffer(item).void
      case None => IO.unit
    }

  private def drain(queue: Queue[IO, StreamItem], maxTextChars: Int): IO[Unit] =
    def flushText(ctx: StreamContext, buffer: StringBuilder): IO[Unit] =
      if buffer.isEmpty then IO.unit
      else
        IO.blocking {
          Console.print(s"$Cyan${buffer.toString}$Reset")
          Console.flush()
        }

    def loop(pending: Option[(StreamContext, StringBuilder)]): IO[Unit] =
      queue.take.flatMap {
        case StreamItem.Shutdown =>
          pending.fold(IO.unit) { case (ctx, buffer) => flushText(ctx, buffer) }
        case StreamItem.CaseStart(ctx) =>
          pending.fold(IO.unit) { case (pctx, buffer) => flushText(pctx, buffer) } *>
            printCaseHeader(ctx) *> loop(None)
        case StreamItem.CaseComplete(ctx, passed, durationMs) =>
          pending.fold(IO.unit) { case (pctx, buffer) => flushText(pctx, buffer) } *>
            printComplete(ctx.skillName, ctx.taskId, ctx.caseNum, passed, durationMs) *> loop(None)
        case StreamItem.Event(ctx, traced) =>
          traced.event match
            case AgentEvent.TextOutput(text, _) =>
              pending match
                case Some((pctx, buffer)) if pctx == ctx =>
                  buffer.append(text)
                  if buffer.length >= maxTextChars then flushText(ctx, buffer) *> loop(None)
                  else loop(Some((ctx, buffer)))
                case Some((pctx, buffer)) =>
                  flushText(pctx, buffer) *>
                    loop(Some((ctx, new StringBuilder().append(text))))
                case None =>
                  val buffer = new StringBuilder().append(text)
                  if buffer.length >= maxTextChars then flushText(ctx, buffer) *> loop(None)
                  else loop(Some((ctx, buffer)))
            case _ =>
              pending.fold(IO.unit) { case (pctx, buffer) => flushText(pctx, buffer) } *>
                printEvent(ctx, traced) *> loop(None)
      }

    loop(None)

  private def formatContext(ctx: StreamContext): String =
    s"${ctx.skillName} ${ctx.taskId}.${ctx.caseNum}"

  private def printEvent(ctx: StreamContext, traced: TracedEvent): IO[Unit] =
    IO.blocking {
      val timeStr = f"[${traced.relativeMs / 1000.0}%6.2fs]"
      val ctxStr = formatContext(ctx)
      val prefix = s"$timeStr [$ctxStr]"

      traced.event match
        case AgentEvent.Prompts(_, _) =>
          ()

        case AgentEvent.TextOutput(_, _) =>
          ()

        case AgentEvent.ToolInvocation(name, _, _, command) =>
          Console.println(s"\n$Yellow$prefix TOOL: $name$Reset")
          command.foreach { cmd =>
            val truncated = truncateLines(cmd, 30)
            Console.println(s"$Yellow$$$$Reset $truncated")
          }

        case AgentEvent.ToolResult(_, stdout, stderr, exitCode, files) =>
          val success = exitCode.forall(_ == 0)
          val color = if success then Green else Red

          if stdout.nonEmpty then
            val truncated = truncateLines(stdout, 20)
            Console.println(s"$color$truncated$Reset")

          if stderr.nonEmpty then
            Console.println(s"${Red}STDERR: ${truncateLines(stderr, 5)}$Reset")

          if files.nonEmpty then Console.println(s"${Gray}FILES: ${files.mkString(", ")}$Reset")

        case AgentEvent.FileCreated(fileId, filename) =>
          Console.println(s"$Green$prefix FILE: $filename ($fileId)$Reset")

        case AgentEvent.Error(message) =>
          Console.println(s"$Red$prefix ERROR: $message$Reset")

        case AgentEvent.McpToolResult(_, content, isError) =>
          val color = if isError then Red else Green
          val header = if isError then "MCP ERROR" else "MCP RESULT"
          Console.println(s"$color$prefix $header$Reset")
          if content.nonEmpty then
            val truncated = truncateLines(content, 10)
            Console.println(s"$Gray$truncated$Reset")

        case AgentEvent.ViewResult(content) =>
          Console.println(s"$Green$prefix VIEW RESULT$Reset")
          if content.nonEmpty then
            val truncated = truncateLines(content, 10)
            Console.println(s"$Gray$truncated$Reset")

        case AgentEvent.TurnComplete(usage) =>
          Console.println(
            s"$Gray$prefix Turn ${usage.turnNum}: +${usage.inputTokens} in / +${usage.outputTokens} out (cumulative: ${usage.cumulativeInputTokens}/${usage.cumulativeOutputTokens})$Reset"
          )

        case AgentEvent.SubTurnComplete(usage) =>
          val toolStr = if usage.hasToolCall then " [tool]" else ""
          Console.println(
            s"$Gray$prefix Sub-turn ${usage.subTurnNum}$toolStr: ${usage.durationMs}ms$Reset"
          )
    }

  private def printCaseHeader(ctx: StreamContext): IO[Unit] =
    IO.blocking {
      Console.println(
        s"$Gray--- Task ${ctx.taskId} Case ${ctx.caseNum} (${ctx.skillName}) ---$Reset"
      )
    }

  /** Print a skill header */
  def printSkillHeader(skillName: String): IO[Unit] =
    IO.blocking {
      val line = "=" * 60
      Console.println(s"\n$Bold$Cyan$line$Reset")
      Console.println(s"$Bold$Cyan  $skillName$Reset")
      Console.println(s"$Bold$Cyan$line$Reset\n")
    }

  /** Print a task header */
  def printTaskHeader(taskId: String, caseNum: Int): IO[Unit] =
    IO.blocking {
      Console.println(s"$Gray--- Task $taskId Case $caseNum ---$Reset")
    }

  /** Print completion message */
  def printComplete(
    skillName: String,
    taskId: String,
    caseNum: Int,
    passed: Boolean,
    durationMs: Long
  ): IO[Unit] =
    IO.blocking {
      val status = if passed then s"${Green}PASSED" else s"${Red}FAILED"
      val duration = f"${durationMs / 1000.0}%.1fs"
      Console.println(s"\n$status$Reset Task $taskId:$caseNum ($duration)")
    }

  private def truncateLines(s: String, maxLines: Int): String =
    val lines = s.split("\n")
    if lines.length <= maxLines then s
    else
      val shown = lines.take(maxLines).mkString("\n")
      val remaining = lines.length - maxLines
      s"$shown\n... ($remaining more lines)"
