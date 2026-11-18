package com.tjclp.xl.ooxml

import java.io.IOException
import java.nio.file.{Files, Path}
import java.nio.file.attribute.{
  GroupPrincipal,
  PosixFileAttributeView,
  PosixFileAttributes,
  PosixFilePermission
}
import java.util.UUID

import scala.jdk.CollectionConverters.*

/**
 * The mode a staged output takes on just before it is renamed into place (GH-868).
 *
 * Every atomic write stages its output in a temp file that `Files.createTempFile` makes owner-only
 * (0600), and a rename keeps that mode, so outputs used to end up 0600 whatever the umask. The
 * staged file stays owner-only while it is written; [[prepare]] then gives it the mode the output
 * should have:
 *   - an output that already exists keeps its own mode and group (a private file stays private).
 *     The staged file belongs to whatever group the writer's new files get, so the original group
 *     is set on it too; where that is not allowed (the writer is not in that group) the group bits
 *     are dropped instead, so no group the original did not grant can open the output;
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
      try modeFor(target, view).foreach(view.setPermissions)
      catch case _: IOException | _: UnsupportedOperationException | _: SecurityException => ()
    }

  private def modeFor(
    target: Path,
    staged: PosixFileAttributeView
  ): Option[java.util.Set[PosixFilePermission]] =
    if Files.isRegularFile(target) then
      val original = Files.readAttributes(target, classOf[PosixFileAttributes])
      Some(narrowed(original.permissions, takesGroup(staged, original.group)))
    else if Files.exists(target) then None // not a file (a directory): no mode to carry over
    else freshMode(Option(target.toAbsolutePath.getParent).getOrElse(Path.of(".")))

  /** Whether `staged` now belongs to `group`, setting it if needed and allowed. */
  private def takesGroup(staged: PosixFileAttributeView, group: GroupPrincipal): Boolean =
    staged.readAttributes.group == group ||
      (try
        staged.setGroup(group)
        true
      catch case _: IOException | _: SecurityException => false)

  /**
   * `permissions`, without the group bits when the staged file could not take the original's group:
   * they would otherwise grant access to a group the original never named.
   */
  private[ooxml] def narrowed(
    permissions: java.util.Set[PosixFilePermission],
    groupKept: Boolean
  ): java.util.Set[PosixFilePermission] =
    if groupKept then permissions
    else
      permissions.asScala
        .filterNot(p =>
          p == PosixFilePermission.GROUP_READ || p == PosixFilePermission.GROUP_WRITE ||
            p == PosixFilePermission.GROUP_EXECUTE
        )
        .asJava

  /** The mode a file newly created in `dir` gets, read from an empty probe. */
  private def freshMode(dir: Path): Option[java.util.Set[PosixFilePermission]] =
    val probe = dir.resolve(s".xl-mode-${UUID.randomUUID()}.tmp")
    Files.createFile(probe)
    try Some(Files.getPosixFilePermissions(probe))
    finally Files.deleteIfExists(probe): Unit
