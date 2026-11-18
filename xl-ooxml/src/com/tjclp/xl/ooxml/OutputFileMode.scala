package com.tjclp.xl.ooxml

import java.io.IOException
import java.nio.file.{Files, Path}
import java.nio.file.attribute.{PosixFileAttributeView, PosixFilePermission}
import java.util.UUID

/**
 * The mode a staged output takes on just before it is renamed into place (GH-868).
 *
 * Every atomic write stages its output in a temp file that `Files.createTempFile` makes owner-only
 * (0600), and a rename keeps that mode, so outputs used to end up 0600 whatever the umask. The
 * staged file stays owner-only while it is written; [[prepare]] then gives it the mode the output
 * should have:
 *   - an output that already exists keeps its own mode (a private file stays private);
 *   - a new output gets the mode any new file gets in its directory — `0666 & ~umask`, narrowed by
 *     a default ACL if the directory has one — read from an empty probe file there, since the JVM
 *     cannot read the umask. The probe holds no data, so it exposes nothing.
 *
 * Best effort: on a file system without POSIX permissions, or if the mode cannot be read or set,
 * the staged file keeps its owner-only mode — never a wider one.
 */
private[xl] object OutputFileMode:

  /** Give `staged`, about to replace `target`, the mode `target` should end up with. */
  def prepare(staged: Path, target: Path): Unit =
    Option(Files.getFileAttributeView(staged, classOf[PosixFileAttributeView])).foreach { view =>
      try modeFor(target).foreach(view.setPermissions)
      catch case _: IOException | _: UnsupportedOperationException | _: SecurityException => ()
    }

  private def modeFor(target: Path): Option[java.util.Set[PosixFilePermission]] =
    if Files.isRegularFile(target) then Some(Files.getPosixFilePermissions(target))
    else if Files.exists(target) then None // not a file (a directory): no mode to carry over
    else freshMode(Option(target.toAbsolutePath.getParent).getOrElse(Path.of(".")))

  /** The mode a file newly created in `dir` gets, read from an empty probe. */
  private def freshMode(dir: Path): Option[java.util.Set[PosixFilePermission]] =
    val probe = dir.resolve(s".xl-mode-${UUID.randomUUID()}.tmp")
    Files.createFile(probe)
    try Some(Files.getPosixFilePermissions(probe))
    finally Files.deleteIfExists(probe): Unit
