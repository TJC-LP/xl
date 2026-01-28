package com.tjclp.xl.agent.anthropic

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.client.okhttp.AnthropicOkHttpClient
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.files.*
import com.tjclp.xl.agent.{AgentConfig, UploadedFile}
import com.tjclp.xl.agent.error.AgentError
import io.github.cdimascio.dotenv.Dotenv

import java.nio.file.{Files, Path, StandardCopyOption}

/** Wrapper around the Anthropic Java SDK with Cats Effect IO */
trait AnthropicClientIO:
  /** Upload a file to the Anthropic Files API */
  def uploadFile(path: Path): IO[UploadedFile]

  /** Download a file from the Anthropic Files API */
  def downloadFile(fileId: String, outputPath: Path): IO[Unit]

  /** Delete a file from the Anthropic Files API */
  def deleteFile(fileId: String): IO[Unit]

  /** List files in the Anthropic Files API */
  def listFiles(limit: Int = 100): IO[List[FileMetadata]]

  /** Get the underlying Java client for advanced operations */
  def underlying: JAnthropicClient

object AnthropicClientIO:

  /** Load API key from environment or .env file */
  def loadApiKey: IO[String] =
    IO.blocking {
      // Try multiple directories for .env file
      val directories =
        List(".", "..", "examples/anthropic-sdk", "examples/anthropic-sdk/spreadsheetbench")
      val dotenv = directories.view
        .map(dir => Dotenv.configure().directory(dir).ignoreIfMissing().load())
        .find(d => d.get("ANTHROPIC_API_KEY") != null)
        .getOrElse(Dotenv.configure().ignoreIfMissing().load())

      Option(java.lang.System.getenv("ANTHROPIC_API_KEY"))
        .orElse(Option(dotenv.get("ANTHROPIC_API_KEY")))
        .getOrElse(
          throw AgentError.ApiKeyMissing(
            "ANTHROPIC_API_KEY not found. Set in .env or as environment variable."
          )
        )
    }

  /** Create an AnthropicClientIO resource that manages client lifecycle */
  def resource(apiKey: String): Resource[IO, AnthropicClientIO] =
    Resource
      .make(
        IO.blocking(
          AnthropicOkHttpClient
            .builder()
            .apiKey(apiKey)
            .build()
        )
      )(client => IO.blocking(client.close()).handleError(_ => ()))
      .map(client =>
        new AnthropicClientIO:
          override def underlying: JAnthropicClient = client

          override def uploadFile(path: Path): IO[UploadedFile] =
            IO.blocking {
              val params = FileUploadParams
                .builder()
                .file(path)
                .addBeta(AnthropicBeta.FILES_API_2025_04_14)
                .build()
              val metadata = client.beta().files().upload(params)
              UploadedFile(metadata.id(), metadata.filename())
            }.adaptError { case e: Exception =>
              AgentError.FileUploadFailed(path.toString, e.getMessage)
            }

          override def downloadFile(fileId: String, outputPath: Path): IO[Unit] =
            IO.blocking {
              val params = FileDownloadParams
                .builder()
                .fileId(fileId)
                .addBeta(AnthropicBeta.FILES_API_2025_04_14)
                .build()

              val httpResponse = client.beta().files().download(params)
              try Files.copy(httpResponse.body(), outputPath, StandardCopyOption.REPLACE_EXISTING)
              finally httpResponse.close()
            }.void
              .adaptError { case e: Exception =>
                AgentError.FileDownloadFailed(fileId, e.getMessage)
              }

          override def deleteFile(fileId: String): IO[Unit] =
            IO.blocking {
              val params = FileDeleteParams
                .builder()
                .fileId(fileId)
                .addBeta(AnthropicBeta.FILES_API_2025_04_14)
                .build()
              client.beta().files().delete(params)
            }.void
              .handleError(_ => ()) // Ignore deletion errors

          override def listFiles(limit: Int = 100): IO[List[FileMetadata]] =
            IO.blocking {
              import scala.jdk.CollectionConverters.*
              val params = FileListParams
                .builder()
                .addBeta(AnthropicBeta.FILES_API_2025_04_14)
                .limit(limit.toLong)
                .build()

              client.beta().files().list(params).data().asScala.toList
            }
      )

  /** Create a client from environment, managing its lifecycle */
  def fromEnv: Resource[IO, AnthropicClientIO] =
    Resource.eval(loadApiKey).flatMap(resource)
