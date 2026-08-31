package com.tjclp.xl.cli.raster

/**
 * Rasterizer backends bundled into this build (ADR-016 platform indirection point).
 *
 * JVM builds bundle Batik — pure JVM, no subprocess, only unavailable where AWT is missing. A build
 * without AWT (Scala Native) supplies an empty list here and [[RasterizerChain.defaultChain]] falls
 * through to the external-tool backends.
 */
object PlatformRasterizers:
  val bundled: List[Rasterizer] = List(BatikRasterizer)
