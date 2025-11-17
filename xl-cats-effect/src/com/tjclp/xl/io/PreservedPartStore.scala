package com.tjclp.xl.io

import java.nio.file.Path
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import cats.effect.{IO, Resource}
import com.tjclp.xl.ooxml.PartManifest

/** Provides lazy streaming access to ZIP entries that XL did not modify. */
trait PreservedPartStore:
  def open: Resource[IO, PreservedPartHandle]

trait PreservedPartHandle:
  def exists(path: String): Boolean
  def listAll: Set[String]
  def streamTo(path: String, output: ZipOutputStream): IO[Unit]

object PreservedPartStore:
  def fromPath(sourcePath: Path, manifest: PartManifest): PreservedPartStore =
    new PreservedPartStoreImpl(Some(sourcePath), manifest)

  val empty: PreservedPartStore = new PreservedPartStoreImpl(None, PartManifest.empty)

private final class PreservedPartStoreImpl(
  sourcePath: Option[Path],
  manifest: PartManifest
) extends PreservedPartStore:

  def open: Resource[IO, PreservedPartHandle] =
    sourcePath match
      case Some(path) =>
        Resource
          .make(IO.blocking(new ZipFile(path.toFile)))(zip => IO.blocking(zip.close()))
          .map(zip => new ZipPreservedPartHandle(zip, manifest))
      case None => Resource.pure(new EmptyPreservedPartHandle)

private final class ZipPreservedPartHandle(zip: ZipFile, manifest: PartManifest)
    extends PreservedPartHandle:
  def exists(path: String): Boolean = manifest.contains(path)

  def listAll: Set[String] = manifest.entries.keySet

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def streamTo(path: String, output: ZipOutputStream): IO[Unit] =
    if !manifest.contains(path) then
      IO.raiseError(new IllegalArgumentException(s"Unknown entry: $path"))
    else
      for
        entry <- IO
          .blocking(Option(zip.getEntry(path)))
          .flatMap(opt =>
            IO.fromOption(opt)(new IllegalStateException(s"Entry missing from source: $path"))
          )
        _ <- Resource
          .make(IO.blocking(zip.getInputStream(entry)))(in =>
            IO.blocking(in.close()).handleErrorWith(_ => IO.unit)
          )
          .use { in =>
            IO.blocking {
              val newEntry = new ZipEntry(path)
              newEntry.setTime(0L)
              newEntry.setMethod(entry.getMethod)
              if entry.getMethod == ZipEntry.STORED then
                newEntry.setSize(entry.getSize)
                newEntry.setCompressedSize(entry.getCompressedSize)
                newEntry.setCrc(entry.getCrc)

              output.putNextEntry(newEntry)
              val buffer = new Array[Byte](8192)
              var read = in.read(buffer)
              while read != -1 do
                output.write(buffer, 0, read)
                read = in.read(buffer)
              output.closeEntry()
            }
          }
      yield ()

private final class EmptyPreservedPartHandle extends PreservedPartHandle:
  def exists(path: String): Boolean = false
  def listAll: Set[String] = Set.empty
  def streamTo(path: String, output: ZipOutputStream): IO[Unit] =
    IO.raiseError(new IllegalStateException("No preserved parts available"))
