

final class build$_ {
def args = build_sc.args$
def scriptPath = """build.sc"""
/*<script>*/
import mill._
import mill.scalalib._
import mill.scalalib.publish._

val scala3Version = "3.7.4"
val organization = "com.tjclp"

trait XLModule extends ScalaModule with PublishModule {
  def scalaVersion = scala3Version

  def publishVersion = "0.1.0-SNAPSHOT"

  def pomSettings = PomSettings(
    description = "Pure Scala 3.7 Excel (OOXML) Library",
    organization = organization,
    url = "https://github.com/tjclp/xl",
    licenses = Seq(License.MIT),
    versionControl = VersionControl.github("tjclp", "xl"),
    developers = Seq()
  )

  def scalacOptions = Seq(
    "-encoding", "utf-8",
    "-feature",
    "-language:implicitConversions",
    "-language:higherKinds",
    "-unchecked",
    "-deprecation",
    "-Xkind-projector"
  )
}

object `xl-core` extends XLModule {
  def ivyDeps = Agg(
    ivy"org.typelevel::cats-core:2.12.0",
    ivy"org.typelevel::cats-laws:2.12.0"
  )
}

object `xl-ooxml` extends XLModule {
  def moduleDeps = Seq(`xl-core`)

  def ivyDeps = Agg(
    ivy"org.scala-lang.modules::scala-xml:2.3.0"
  )
}

object `xl-cats-effect` extends XLModule {
  def moduleDeps = Seq(`xl-core`, `xl-ooxml`)

  def ivyDeps = Agg(
    ivy"org.typelevel::cats-effect:3.5.7",
    ivy"co.fs2::fs2-core:3.11.0",
    ivy"co.fs2::fs2-io:3.11.0"
  )
}

object `xl-evaluator` extends XLModule {
  def moduleDeps = Seq(`xl-core`)

  def ivyDeps = Agg(
    ivy"org.typelevel::cats-core:2.12.0"
  )
}

object `xl-testkit` extends XLModule {
  def moduleDeps = Seq(`xl-core`, `xl-ooxml`)

  def ivyDeps = Agg(
    ivy"org.typelevel::cats-laws:2.12.0",
    ivy"org.scalacheck::scalacheck:1.18.1",
    ivy"org.scalameta::munit:1.0.3",
    ivy"org.typelevel::munit-cats-effect:2.0.0",
    ivy"org.typelevel::scalacheck-effect:1.0.4"
  )
}

// Test modules for each project
object tests extends Module {
  trait TestModule extends XLModule with TestModule.ScalaTest {
    def ivyDeps = Agg(
      ivy"org.scalameta::munit:1.0.3",
      ivy"org.typelevel::munit-cats-effect:2.0.0",
      ivy"org.scalacheck::scalacheck:1.18.1"
    )

    def testFramework = "munit.Framework"
  }

  object core extends TestModule {
    def moduleDeps = Seq(`xl-core`, `xl-testkit`)
  }

  object ooxml extends TestModule {
    def moduleDeps = Seq(`xl-ooxml`, `xl-testkit`)
  }

  object `cats-effect` extends TestModule {
    def moduleDeps = Seq(`xl-cats-effect`, `xl-testkit`)
  }

  object evaluator extends TestModule {
    def moduleDeps = Seq(`xl-evaluator`, `xl-testkit`)
  }
}
/*</script>*/ /*<generated>*//*</generated>*/
}

object build_sc {
  private var args$opt0 = Option.empty[Array[String]]
  def args$set(args: Array[String]): Unit = {
    args$opt0 = Some(args)
  }
  def args$opt: Option[Array[String]] = args$opt0
  def args$: Array[String] = args$opt.getOrElse {
    sys.error("No arguments passed to this script")
  }

  lazy val script = new build$_

  def main(args: Array[String]): Unit = {
    args$set(args)
    val _ = script.hashCode() // hashCode to clear scalac warning about pure expression in statement position
  }
}

export build_sc.script as `build`

