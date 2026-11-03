package com.tjclp.xl.formula

import java.util.concurrent.atomic.AtomicReference

/**
 * Run a body on a fresh platform thread with a chosen stack size rather than the test runner's: a
 * `StackOverflowError` fails the test by name (it is caught, not swallowed), and so does a body
 * still running when the budget runs out. Shared by the specs that pin stack independence (#829).
 */
object SmallStack:
  val BudgetMs: Long = 60000L

  def run[A](stackBytes: Long, name: String = "small-stack")(body: => A): A =
    val out = new AtomicReference[Either[Throwable, A]](Left(new IllegalStateException("not run")))
    val thread = Thread
      .ofPlatform()
      .name(name)
      .daemon(true)
      .stackSize(stackBytes)
      .start(() =>
        out.set(
          try Right(body)
          catch case e: Throwable => Left(e)
        )
      )
    if !thread.join(java.time.Duration.ofMillis(BudgetMs)) then
      thread.interrupt()
      throw new AssertionError(s"$name did not finish within ${BudgetMs}ms")
    out.get().fold(e => throw e, identity)
