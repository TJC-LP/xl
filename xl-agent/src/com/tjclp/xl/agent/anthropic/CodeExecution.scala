package com.tjclp.xl.agent.anthropic

import cats.effect.{IO, Ref}
import cats.effect.std.Queue
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.core.JsonValue
import com.anthropic.helpers.BetaMessageAccumulator
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.{AgentConfig, AgentEvent, TokenUsage, UploadedFile}
import com.tjclp.xl.agent.error.AgentError

import java.util.concurrent.atomic.AtomicBoolean
import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

/** Handles code execution requests to the Anthropic API */
object CodeExecution:

  /** Send a message with code execution capability and stream the response */
  def sendRequest(
    client: JAnthropicClient,
    config: AgentConfig,
    systemPrompt: String,
    userPrompt: String,
    containerUploads: List[String], // File IDs to upload to container
    eventQueue: Queue[IO, AgentEvent],
    configureRequest: MessageCreateParams.Builder => MessageCreateParams.Builder =
      identity // Strategy-specific configuration (tools, betas, container)
  ): IO[BetaMessage] =
    IO.blocking {
      import java.util.{List as JList, Map as JMap}

      // Build content blocks: text + container uploads
      val contentBlocks = new java.util.ArrayList[JMap[String, Any]]()
      contentBlocks.add(JMap.of("type", "text", "text", userPrompt))
      containerUploads.foreach { fileId =>
        contentBlocks.add(JMap.of("type", "container_upload", "file_id", fileId))
      }

      val baseBuilder = MessageCreateParams
        .builder()
        .model(config.model)
        .maxTokens(config.maxTokens.toLong)
        .system(systemPrompt)
        .addUserMessage("placeholder")
        // Override messages to include container_upload blocks
        .putAdditionalBodyProperty(
          "messages",
          JsonValue.from(
            JList.of(
              JMap.of(
                "role",
                "user",
                "content",
                contentBlocks
              )
            )
          )
        )

      // Apply strategy-specific configuration (tools, betas, container)
      val params = configureRequest(baseBuilder).build()

      // Stream response
      val accumulator = BetaMessageAccumulator.create()
      val streamResponse = client.beta().messages().createStreaming(params)
      val streamProcessor = new StreamEventProcessor(eventQueue, config.verbose)
      val interrupted = new AtomicBoolean(false)

      try
        streamResponse.stream().forEach { event =>
          if Thread.currentThread().isInterrupted() then
            interrupted.set(true)
            streamResponse.close()
          else
            accumulator.accumulate(event)
            // Process event synchronously (we're already in IO.blocking)
            import cats.effect.unsafe.implicits.global
            streamProcessor.process(event).unsafeRunSync()
        }
        if config.verbose then println() // Final newline after streaming

        if interrupted.get() then throw new InterruptedException("Stream interrupted by user")

        accumulator.message()
      finally streamResponse.close()
    }.adaptError { case e: Exception =>
      AgentError.StreamingError(e.getMessage)
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
