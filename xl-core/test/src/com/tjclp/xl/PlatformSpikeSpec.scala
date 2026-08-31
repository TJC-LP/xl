package com.tjclp.xl

import com.tjclp.xl.platform.Sha256

/** Spike probes: platform shim correctness + classpath-resource availability per platform. */
class PlatformSpikeSpec extends munit.FunSuite:

  private def hex(b: Array[Byte]): String = b.map(x => f"$x%02x").mkString

  test("sha256 known-answer vectors (NIST)"):
    assertEquals(
      hex(Sha256.digest(Array.empty[Byte])),
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    )
    assertEquals(
      hex(Sha256.digest("abc".getBytes("UTF-8"))),
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    )
    // 56-byte message exercises the two-block padding path
    assertEquals(
      hex(
        Sha256.digest("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq".getBytes("UTF-8"))
      ),
      "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    )

  test("incremental hasher matches one-shot digest"):
    val data = Array.tabulate[Byte](100000)(i => (i % 251).toByte)
    val h = Sha256.hasher()
    h.update(data, 0, 33333)
    h.update(data, 33333, 66667)
    assert(Sha256.isEqual(h.digest(), Sha256.digest(data)))

  test("classpath resource is readable (embed-resources probe)"):
    val in = Option(getClass.getResourceAsStream("/spike-probe.txt"))
    assert(in.isDefined, "getResourceAsStream returned null for /spike-probe.txt")
    val content = in.map(s => scala.io.Source.fromInputStream(s).mkString.trim)
    assertEquals(content, Some("spike-ok"))
