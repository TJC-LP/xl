package com.tjclp.xl.platform

import java.nio.charset.StandardCharsets

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * NIST FIPS 180-4 known-answer tests plus an incremental ≡ one-shot property. These pin the shim's
 * output so the pure-Scala twin planned for non-JVM platforms (ADR-016 wave B1) can be verified
 * against the exact same vectors.
 */
class Sha256Spec extends ScalaCheckSuite:

  private def hex(bytes: Array[Byte]): String =
    bytes.map("%02x".format(_)).mkString

  private def ascii(s: String): Array[Byte] = s.getBytes(StandardCharsets.US_ASCII)

  test("digest: empty input (NIST KAT)") {
    assertEquals(
      hex(Sha256.digest(Array.emptyByteArray)),
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    )
  }

  test("digest: \"abc\" (NIST KAT)") {
    assertEquals(
      hex(Sha256.digest(ascii("abc"))),
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    )
  }

  test("digest: 56-byte two-block message (NIST KAT)") {
    val msg = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
    assertEquals(msg.length, 56)
    assertEquals(
      hex(Sha256.digest(ascii(msg))),
      "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    )
  }

  test("hasher: chunked updates with offsets match the one-shot digest") {
    val msg = ascii("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq")
    val h = Sha256.hasher()
    h.update(msg, 0, 10)
    h.update(msg, 10, 0) // zero-length update is a no-op
    h.update(msg, 10, msg.length - 10)
    assert(Sha256.isEqual(h.digest(), Sha256.digest(msg)))
  }

  property("hasher: any chunking of any input equals the one-shot digest") {
    val genBytes = Gen.containerOf[Array, Byte](Gen.choose(Byte.MinValue, Byte.MaxValue))
    forAll(genBytes, Gen.choose(1, 64)) { (bytes: Array[Byte], chunk: Int) =>
      val h = Sha256.hasher()
      bytes.grouped(chunk).foldLeft(0) { (off, group) =>
        h.update(bytes, off, group.length)
        off + group.length
      }
      Prop(Sha256.isEqual(h.digest(), Sha256.digest(bytes))) :| s"chunk=$chunk len=${bytes.length}"
    }
  }

  property("isEqual agrees with element-wise equality") {
    val genBytes = Gen.containerOf[Array, Byte](Gen.choose(Byte.MinValue, Byte.MaxValue))
    forAll(genBytes, genBytes) { (a: Array[Byte], b: Array[Byte]) =>
      (Sha256.isEqual(a, b) ?= a.sameElements(b)) && Prop(Sha256.isEqual(a, a.clone()))
    }
  }
