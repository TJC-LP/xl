package com.tjclp.xl.testing

import java.util.concurrent.TimeoutException
import java.util.concurrent.atomic.AtomicReference

/**
 * Run a body on a fresh platform thread with a chosen stack size rather than the test runner's.
 * [[attempt]] returns what the body threw (a `StackOverflowError` included: `Try` would not catch
 * it) or, when the body is still running after [[BudgetMs]], a `TimeoutException`; [[run]] rethrows
 * either, so a test fails by name. Shared by the specs that pin stack independence (#829) and the
 * lint oracle's small-stack contract (#681).
 *
 * The thread is a daemon and is interrupted on timeout, but a CPU-bound body that never checks the
 * interrupt flag keeps running until the JVM exits: the failure is reported, the core is not freed.
 */
object SmallStack:
  val BudgetMs: Long = 60000L

  def attempt[A](stackBytes: Long, name: String = "small-stack")(body: => A): Either[Throwable, A] =
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
    if thread.join(java.time.Duration.ofMillis(BudgetMs)) then out.get()
    else
      thread.interrupt()
      Left(new TimeoutException(s"$name did not finish within ${BudgetMs}ms"))

  def run[A](stackBytes: Long, name: String = "small-stack")(body: => A): A =
    attempt(stackBytes, name)(body).fold(e => throw e, identity)
