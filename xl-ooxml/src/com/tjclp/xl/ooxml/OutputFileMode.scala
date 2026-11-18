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
 *   - an output that already exists keeps its owner and others bits, and its group where the writer
 *     may set it. Its group bits are capped at its others bits (`0664` becomes `0644`, `0640`
 *     becomes `0600`): the group bits cannot be trusted for the staged file, whose group may not be
 *     the original's, and which carries no ACL (on a file with a POSIX ACL the group bits are the
 *     ACL mask, not the group's own entry). The JVM cannot read POSIX ACLs, so the cap is applied
 *     every time. No user or group gains access; a group-shared file loses the group's extra
 *     rights;
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
      keepGroup(staged, original.group)
      Some(groupCappedAtOthers(original.permissions))
    else if Files.exists(target) then None // not a file (a directory): no mode to carry over
    else freshMode(Option(target.toAbsolutePath.getParent).getOrElse(Path.of(".")))

  /** Give `staged` the original's group where allowed; the capped bits make it safe either way. */
  private def keepGroup(staged: PosixFileAttributeView, group: GroupPrincipal): Unit =
    try if staged.readAttributes.group != group then staged.setGroup(group)
    catch case _: IOException | _: SecurityException => ()

  private val groupToOthers: List[(PosixFilePermission, PosixFilePermission)] = List(
    PosixFilePermission.GROUP_READ -> PosixFilePermission.OTHERS_READ,
    PosixFilePermission.GROUP_WRITE -> PosixFilePermission.OTHERS_WRITE,
    PosixFilePermission.GROUP_EXECUTE -> PosixFilePermission.OTHERS_EXECUTE
  )

  /** `permissions` with each group bit kept only where the matching others bit is set. */
  private[ooxml] def groupCappedAtOthers(
    permissions: java.util.Set[PosixFilePermission]
  ): java.util.Set[PosixFilePermission] =
    val dropped = groupToOthers.collect {
      case (g, o) if permissions.contains(g) && !permissions.contains(o) => g
    }.toSet
    permissions.asScala.filterNot(dropped.contains).asJava

  /** The mode a file newly created in `dir` gets, read from an empty probe. */
  private def freshMode(dir: Path): Option[java.util.Set[PosixFilePermission]] =
    val probe = dir.resolve(s".xl-mode-${UUID.randomUUID()}.tmp")
    Files.createFile(probe)
    try Some(Files.getPosixFilePermissions(probe))
    finally Files.deleteIfExists(probe): Unit
