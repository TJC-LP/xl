package com.tjclp.xl.agent.anthropic

import cats.effect.IO
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.tjclp.xl.agent.error.AgentError
import io.circe.*
import io.circe.generic.semiauto.*
import io.circe.parser.*
import io.circe.syntax.*
import okhttp3.{MediaType, MultipartBody, OkHttpClient, Request, RequestBody}

import java.nio.file.{Files, Path}
import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*
import scala.util.Using

/** Metadata for a custom skill */
case class SkillMetadata(
  id: String,
  displayTitle: String,
  version: Option[String] = None
)

object SkillMetadata:
  given Decoder[SkillMetadata] = Decoder.instance { c =>
    for
      id <- c.get[String]("id")
      displayTitle <- c.get[String]("display_title")
      version <- c.get[Option[String]]("version")
    yield SkillMetadata(id, displayTitle, version)
  }

/** Wrapper for Anthropic Skills API (beta) */
object SkillsApi:

  private val ApiBase = "https://api.anthropic.com/v1"
  private val SkillsBeta = "skills-2025-10-02"
  private val JsonMediaType = MediaType.get("application/json")
  private val OctetMediaType = MediaType.get("application/octet-stream")

  /** List custom skills */
  def listCustomSkills(apiKey: String): IO[List[SkillMetadata]] =
    IO.blocking {
      val client = new OkHttpClient()
      val request = new Request.Builder()
        .url(s"$ApiBase/skills?source=custom")
        .addHeader("x-api-key", apiKey)
        .addHeader("anthropic-version", "2023-06-01")
        .addHeader("anthropic-beta", SkillsBeta)
        .get()
        .build()

      val response = client.newCall(request).execute()
      try
        if !response.isSuccessful then
          throw AgentError.SkillsApiError(
            s"Failed to list skills: ${response.code()} ${response.message()}"
          )

        val body = response.body().string()
        decode[SkillsListResponse](body) match
          case Right(r) => r.data
          case Left(e) =>
            throw AgentError.SkillsApiError(s"Failed to parse skills response: ${e.getMessage}")
      finally response.close()
    }

  /** Create a new custom skill from files */
  def createSkill(
    apiKey: String,
    displayTitle: String,
    files: List[(String, Array[Byte])]
  ): IO[SkillMetadata] =
    IO.blocking {
      val client = new OkHttpClient()

      // Build multipart form
      val bodyBuilder = new MultipartBody.Builder()
        .setType(MultipartBody.FORM)
        .addFormDataPart("display_title", displayTitle)

      // Add each file
      files.foreach { case (filename, content) =>
        bodyBuilder.addFormDataPart(
          "files",
          filename,
          RequestBody.create(content, OctetMediaType)
        )
      }

      val request = new Request.Builder()
        .url(s"$ApiBase/skills")
        .addHeader("x-api-key", apiKey)
        .addHeader("anthropic-version", "2023-06-01")
        .addHeader("anthropic-beta", SkillsBeta)
        .post(bodyBuilder.build())
        .build()

      val response = client.newCall(request).execute()
      try
        if !response.isSuccessful then
          val errorBody = Option(response.body()).map(_.string()).getOrElse("")
          throw AgentError.SkillsApiError(
            s"Failed to create skill: ${response.code()} ${response.message()} - $errorBody"
          )

        val body = response.body().string()
        decode[SkillMetadata](body) match
          case Right(skill) => skill
          case Left(e) =>
            throw AgentError.SkillsApiError(s"Failed to parse skill response: ${e.getMessage}")
      finally response.close()
    }

  /** Get existing xl-cli skill or create a new one from zip file */
  def getOrCreateXlSkill(apiKey: String, skillZipPath: Path): IO[String] =
    for
      // Check for existing skill
      existingSkills <- listCustomSkills(apiKey)
      existingId = existingSkills.find(_.displayTitle == "xl-cli").map(_.id)

      skillId <- existingId match
        case Some(id) =>
          IO.println(s"   Found existing xl-cli skill: $id") *> IO.pure(id)
        case None =>
          for
            _ <- IO.println(s"   Creating xl-cli skill from ${skillZipPath.getFileName}...")
            files <- extractZipFiles(skillZipPath)
            skill <- createSkill(apiKey, "xl-cli", files)
            _ <- IO.println(s"   Created skill: ${skill.id}")
          yield skill.id
    yield skillId

  /** Extract files from a zip archive with xl-cli/ prefix */
  private def extractZipFiles(zipPath: Path): IO[List[(String, Array[Byte])]] =
    IO.blocking {
      Using.resource(new ZipFile(zipPath.toFile)) { zip =>
        zip.entries().asScala.toList.flatMap { entry =>
          if entry.isDirectory then None
          else
            val content = Using.resource(zip.getInputStream(entry)) { is =>
              is.readAllBytes()
            }
            // Add xl-cli/ prefix for Skills API requirement
            Some((s"xl-cli/${entry.getName}", content))
        }
      }
    }

  // Response wrapper for skills list
  private case class SkillsListResponse(data: List[SkillMetadata])

  private object SkillsListResponse:
    given Decoder[SkillsListResponse] = deriveDecoder
