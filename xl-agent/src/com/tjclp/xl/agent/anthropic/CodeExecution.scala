package com.tjclp.xl.agent.anthropic

import cats.effect.{IO, Ref}
import cats.effect.std.{Dispatcher, Queue}
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.helpers.BetaMessageAccumulator
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.{AgentConfig, AgentEvent, TokenUsage, UploadedFile}
import com.tjclp.xl.agent.error.AgentError

import java.util.concurrent.atomic.AtomicBoolean
import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

/** Handles code execution requests to the Anthropic API */
object CodeExecution:

  /**
   * Build the request params for one API turn.
   *
   * On a pause_turn resume (issue #344) the request is the original one with each paused assistant
   * message appended to `messages` verbatim (`BetaMessage.toParam`) — per the Anthropic contract
   * for server-side tools, the API detects the trailing server-tool state and resumes the turn; a
   * synthetic "continue" user message must NOT be added. The last paused turn's container id is
   * passed back so files written by earlier bash cycles stay visible to the resumed turn.
   *
   * The container id is applied AFTER `configureRequest`: strategies declare a fresh skills
   * container via `builder.container(BetaContainerParams)` and the last `.container()` write on the
   * builder wins, so setting the id first would provision a new container and lose the paused
   * turn's files. Per the documented container-reuse shape, the id alone is passed on resume — the
   * container already holds the skill files loaded at creation, so skills are not re-declared.
   */
  private[anthropic] def buildParams(
    config: AgentConfig,
    systemPrompt: String,
    userPrompt: String,
    containerUploads: List[String],
    resumeTurns: Vector[BetaMessage],
    configureRequest: MessageCreateParams.Builder => MessageCreateParams.Builder = identity
  ): MessageCreateParams =
    // Prompt caching (5m TTL): the code-execution loop re-samples the
    // conversation on every server-side sub-turn, and a task's cases
    // share tools + system + skill context — both re-read from cache.
    val cacheMarker = BetaCacheControlEphemeral.builder().build()

    // System block breakpoint: shared read point across a task's cases
    val systemBlock = BetaTextBlockParam
      .builder()
      .text(systemPrompt)
      .cacheControl(cacheMarker)
      .build()

    // Build content blocks: text + container uploads
    val contentBlocks = new java.util.ArrayList[BetaContentBlockParam]()
    contentBlocks.add(
      BetaContentBlockParam.ofText(BetaTextBlockParam.builder().text(userPrompt).build())
    )
    containerUploads.foreach { fileId =>
      contentBlocks.add(
        BetaContentBlockParam.ofContainerUpload(
          BetaContainerUploadBlockParam.builder().fileId(fileId).build()
        )
      )
    }

    val baseBuilder = MessageCreateParams
      .builder()
      .model(config.model)
      .maxTokens(config.maxTokens.toLong)
      .systemOfBetaTextBlockParams(java.util.List.of(systemBlock))
      .addUserMessageOfBetaContentBlockParams(contentBlocks)
      // Top-level cache_control auto-places on the last cacheable
      // block, covering the whole initial prompt for the loop
      .cacheControl(cacheMarker)

    // pause_turn resume: paused assistant turns go back into messages verbatim
    resumeTurns.foreach(paused => baseBuilder.addMessage(paused.toParam()))

    // Apply strategy-specific configuration (tools, betas, container)
    val configured = configureRequest(baseBuilder)

    // Reuse the paused turn's container so the resumed bash cycles see its files. Must come
    // after configureRequest: the strategy's container(BetaContainerParams) write would
    // otherwise clobber the id and provision a fresh, empty container (see scaladoc).
    resumeTurns.lastOption
      .flatMap(_.container().toScala)
      .foreach(container => configured.container(container.id()))

    configured.build()

  /**
   * Send a message with code execution capability and stream the response.
   *
   * `resumeTurns` carries the paused assistant turns of a pause_turn resume (issue #344), oldest
   * first; empty for the initial request. See [[buildParams]] for the re-send shape.
   */
  def sendRequest(
    client: JAnthropicClient,
    config: AgentConfig,
    systemPrompt: String,
    userPrompt: String,
    containerUploads: List[String], // File IDs to upload to container
    eventQueue: Queue[IO, AgentEvent],
    configureRequest: MessageCreateParams.Builder => MessageCreateParams.Builder =
      identity, // Strategy-specific configuration (tools, betas, container)
    onEvent: AgentEvent => IO[Unit] = _ => IO.unit, // Real-time event callback for tracing
    resumeTurns: Vector[BetaMessage] = Vector.empty // Paused turns to append (pause_turn resume)
  ): IO[BetaMessage] =
    val promptsEvent = AgentEvent.Prompts(systemPrompt, userPrompt)

    // Use Dispatcher.sequential to preserve event order while avoiding unsafeRunSync()
    // which blocks the CE compute pool and causes CPU starvation warnings
    Dispatcher.sequential[IO].use { dispatcher =>
      for
        // Emit prompts for tracing
        _ <- eventQueue.offer(promptsEvent)
        _ <- onEvent(promptsEvent)
        streamProcessor <- StreamEventProcessor.create(eventQueue, onEvent, config.verbose)
        result <- IO
          .blocking {
            val params = buildParams(
              config,
              systemPrompt,
              userPrompt,
              containerUploads,
              resumeTurns,
              configureRequest
            )

            // Stream response
            val accumulator = BetaMessageAccumulator.create()
            val streamResponse = client.beta().messages().createStreaming(params)
            val interrupted = new AtomicBoolean(false)

            try
              streamResponse.stream().forEach { event =>
                if Thread.currentThread().isInterrupted() then
                  interrupted.set(true)
                  streamResponse.close()
                else
                  accumulator.accumulate(event)
                  // Fire-and-forget: doesn't block the OkHttp callback thread
                  // Dispatcher.sequential ensures events are processed in order
                  dispatcher.unsafeRunAndForget(streamProcessor.process(event))
              }
              if config.verbose then println() // Final newline after streaming

              if interrupted.get() then throw new InterruptedException("Stream interrupted by user")

              accumulator.message()
            finally streamResponse.close()
          }
          .onError { case _ =>
            // Best-effort recovery of the open turn's token usage when the stream dies
            // mid-turn (issue #340). Routed through the same sequential dispatcher so it
            // runs after every already-submitted stream event (a pending message_stop
            // must close the turn first, or its usage would be double-counted).
            IO.fromFuture(IO(dispatcher.unsafeToFuture(streamProcessor.flushPartialTurn)))
              .attempt
              .void
          }
          .adaptError { case e: Exception =>
            AgentError.StreamingError(e.getMessage)
          }
      yield result
    }

  /** Extract text content from response */
  def extractResponseText(response: BetaMessage): String =
    response
      .content()
      .asScala
      .flatMap(_.text().toScala)
      .map(_.text())
      .mkString("\n")

  /**
   * Extract output file ID from response content blocks.
   *
   * Files saved to $OUTPUT_DIR appear in bashCodeExecutionToolResult blocks.
   */
  def extractOutputFileId(response: BetaMessage, verbose: Boolean = false): Option[String] =
    val fileIds = response
      .content()
      .asScala
      .flatMap { block =>
        // Method 1: containerUpload blocks
        val containerFileId = block.containerUpload().toScala.map(_.fileId())

        // Method 2: bashCodeExecutionToolResult via typed API
        val bashOutputFileIds =
          block.bashCodeExecutionToolResult().toScala.toList.flatMap { toolResult =>
            toolResult.content().betaBashCodeExecutionResultBlock().toScala.toList.flatMap {
              resultBlock =>
                resultBlock.content().asScala.map(_.fileId())
            }
          }

        // Method 3: bashCodeExecutionToolResult via raw JSON (fallback)
        val rawBashFileIds =
          block.bashCodeExecutionToolResult().toScala.toList.flatMap { toolResult =>
            toolResult._content().asKnown().toScala.toList.flatMap { contentField =>
              contentField._json().toScala.toList.flatMap { json =>
                json.asObject().toScala.toList.flatMap { obj =>
                  obj.asScala.get("content").toList.flatMap { contentArray =>
                    contentArray.asArray().toScala.toList.flatMap { arr =>
                      arr.asScala.flatMap { item =>
                        item.asObject().toScala.flatMap { outputBlock =>
                          outputBlock.asScala.get("file_id").flatMap(_.asString().toScala)
                        }
                      }
                    }
                  }
                }
              }
            }
          }

        (containerFileId.toList ++ bashOutputFileIds ++ rawBashFileIds).distinct
      }
      .distinct

    if verbose && fileIds.nonEmpty then println(s"    DEBUG: Found file_ids: $fileIds")

    fileIds.lastOption
