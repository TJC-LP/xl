package com.tjclp.xl.io

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import com.tjclp.xl.ooxml.{PartManifest, PartManifestBuilder}
import munit.CatsEffectSuite

class PreservedPartStoreSpec extends CatsEffectSuite:

  test("empty store exposes no entries") {
    PreservedPartStore.empty.open.use { handle =>
      IO {
        assertEquals(handle.listAll, Set.empty[String])
        assert(!handle.exists("xl/workbook.xml"))
      }
    }
  }

  test("streamTo copies bytes without materializing") {
    val data = "<chart/>".getBytes("UTF-8")
    tempZip(Map("xl/charts/chart1.xml" -> data)).use { source =>
      val manifest = new PartManifestBuilder().recordUnparsed("xl/charts/chart1.xml").build()
      val store = PreservedPartStore.fromPath(source, manifest)
      store.open.use { handle =>
        assert(handle.exists("xl/charts/chart1.xml"))
        val baos = new ByteArrayOutputStream()
        val zipOut = new ZipOutputStream(baos)
        for
          _ <- handle.streamTo("xl/charts/chart1.xml", zipOut)
          _ <- IO.blocking(zipOut.close())
          bytes <- IO(baos.toByteArray)
        yield {
          val copied = readEntry(bytes, "xl/charts/chart1.xml")
          assertEquals(copied.toSeq, data.toSeq)
        }
      }
    }
  }

  private def tempZip(entries: Map[String, Array[Byte]]): Resource[IO, Path] =
    Resource.make {
      IO.blocking {
        val path = Files.createTempFile("preserved-store", ".zip")
        val out = new ZipOutputStream(Files.newOutputStream(path))
        try
          entries.foreach { case (name, bytes) =>
            val entry = new ZipEntry(name)
            out.putNextEntry(entry)
            out.write(bytes)
            out.closeEntry()
          }
        finally
          out.close()
        path
      }
    }(path => IO.blocking(Files.deleteIfExists(path)).void)

  private def readEntry(bytes: Array[Byte], name: String): Array[Byte] =
    val in = new ZipInputStream(new ByteArrayInputStream(bytes))
    val buffer = new Array[Byte](1024)
    @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
    def readFully(): Array[Byte] =
      val baos = new ByteArrayOutputStream()
      var read = in.read(buffer)
      while read != -1 do
        baos.write(buffer, 0, read)
        read = in.read(buffer)
      baos.toByteArray
    @annotation.tailrec
    def loop(entry: ZipEntry | Null): Array[Byte] =
      if entry == null then throw new IllegalStateException(s"Entry $name not found in zip stream")
      else if entry.getName == name then readFully()
      else loop(in.getNextEntry)
    try loop(in.getNextEntry)
    finally in.close()
