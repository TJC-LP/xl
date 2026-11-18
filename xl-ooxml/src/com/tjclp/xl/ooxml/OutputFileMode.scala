package com.tjclp.xl.ooxml

import java.io.IOException
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.nio.file.attribute.{
  GroupPrincipal,
  PosixFileAttributeView,
  PosixFileAttributes,
  PosixFilePermission
}
import java.util.UUID
import java.util.concurrent.TimeUnit

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
 *     is set on it too. Where that is not allowed (the writer is not in that group), or where the
 *     original's group bits cannot be trusted (an extended ACL, whose mask they are), the group
 *     keeps only what everyone else may do — no group gains access the original did not give it;
 *   - a new output gets the mode any new file gets in its directory — `0666 & ~umask`, narrowed by
 *     a default ACL if the directory has one — read from an empty probe file there, since the JVM
 *     cannot read the umask. The probe holds no data, so it exposes nothing.
 *
 * An extended ACL on a replaced file is not carried over (the JVM cannot read POSIX ACLs), so the
 * users and groups it names lose that access; they never gain any.
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
      val perms = original.permissions
      // ACL check last: it runs a process, and only matters when the group bits could widen
      val groupTrusted = takesGroup(staged, original.group) &&
        (groupWithinOthers(perms) || !mayHaveExtendedAcl(target))
      Some(narrowed(perms, groupTrusted))
    else if Files.exists(target) then None // not a file (a directory): no mode to carry over
    else freshMode(Option(target.toAbsolutePath.getParent).getOrElse(Path.of(".")))

  /** Whether `staged` now belongs to `group`, setting it if needed and allowed. */
  private def takesGroup(staged: PosixFileAttributeView, group: GroupPrincipal): Boolean =
    staged.readAttributes.group == group ||
      (try
        staged.setGroup(group)
        true
      catch case _: IOException | _: SecurityException => false)

  private val groupToOthers: List[(PosixFilePermission, PosixFilePermission)] = List(
    PosixFilePermission.GROUP_READ -> PosixFilePermission.OTHERS_READ,
    PosixFilePermission.GROUP_WRITE -> PosixFilePermission.OTHERS_WRITE,
    PosixFilePermission.GROUP_EXECUTE -> PosixFilePermission.OTHERS_EXECUTE
  )

  /** Whether the group bits grant nothing the others bits do not. */
  private def groupWithinOthers(permissions: java.util.Set[PosixFilePermission]): Boolean =
    groupToOthers.forall((g, o) => !permissions.contains(g) || permissions.contains(o))

  /**
   * `permissions`, with the group bits cut down to the others bits when they cannot be trusted for
   * the staged file's group (`0660` becomes `0600`, `0644` stays `0644`): the group then gets no
   * more than anyone else already had.
   */
  private[ooxml] def narrowed(
    permissions: java.util.Set[PosixFilePermission],
    groupTrusted: Boolean
  ): java.util.Set[PosixFilePermission] =
    if groupTrusted then permissions
    else
      val dropped = groupToOthers.collect {
        case (g, o) if permissions.contains(g) && !permissions.contains(o) => g
      }.toSet
      permissions.asScala.filterNot(dropped.contains).asJava

  /**
   * Whether `target` may carry an extended ACL. On such a file the group bits are the ACL mask, not
   * the owning group's own entry, so copying them could widen the group's access. The JVM cannot
   * read POSIX ACLs, so `ls -ld` is asked (its `+` marker); when it cannot answer, assume an ACL.
   */
  private def mayHaveExtendedAcl(target: Path): Boolean =
    try
      val proc = new ProcessBuilder("ls", "-ld", "--", target.toAbsolutePath.toString)
        .redirectErrorStream(true)
        .start()
      proc.getOutputStream.close()
      val out = new String(proc.getInputStream.readAllBytes(), StandardCharsets.UTF_8)
      if !proc.waitFor(5, TimeUnit.SECONDS) then
        proc.destroyForcibly(): Unit
        true
      else proc.exitValue != 0 || out.lift(10).contains('+')
    catch
      case _: InterruptedException =>
        Thread.currentThread().interrupt()
        true
      case _: IOException | _: SecurityException => true

  /** The mode a file newly created in `dir` gets, read from an empty probe. */
  private def freshMode(dir: Path): Option[java.util.Set[PosixFilePermission]] =
    val probe = dir.resolve(s".xl-mode-${UUID.randomUUID()}.tmp")
    Files.createFile(probe)
    try Some(Files.getPosixFilePermissions(probe))
    finally Files.deleteIfExists(probe): Unit
