import mill._
import mill.scalalib._
import mill.scalalib.publish._

val scala3Version = "3.7.3"
val organization = "com.tjclp"

trait XLModule extends ScalaModule with PublishModule {
  override def scalaVersion = scala3Version

  override def publishVersion = "0.1.0-SNAPSHOT"

  override def pomSettings = PomSettings(
    description = "Pure Scala 3.7 Excel (OOXML) Library",
    organization = organization,
    url = "https://github.com/TJC-LP/xl",
    licenses = Seq(License.MIT),
    versionControl = VersionControl.github("tjclp", "xl"),
    developers = Seq()
  )

  override def scalacOptions = Seq(
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
  override def mvnDeps = Seq(
    mvn"org.typelevel::cats-core:2.12.0",
    mvn"org.typelevel::cats-laws:2.12.0"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}

object `xl-macros` extends XLModule {
  override def moduleDeps = Seq(`xl-core`)

  override def mvnDeps = Seq(
    mvn"org.typelevel::cats-core:2.12.0"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}

object `xl-ooxml` extends XLModule {
  override def moduleDeps = Seq(`xl-core`)

  override def mvnDeps = Seq(
    mvn"org.scala-lang.modules::scala-xml:2.3.0"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}

object `xl-cats-effect` extends XLModule {
  override def moduleDeps = Seq(`xl-core`, `xl-ooxml`)

  override def mvnDeps = Seq(
    mvn"org.typelevel::cats-effect:3.5.7",
    mvn"co.fs2::fs2-core:3.11.0",
    mvn"co.fs2::fs2-io:3.11.0"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.typelevel::munit-cats-effect:2.0.0",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}

object `xl-evaluator` extends XLModule {
  override def moduleDeps = Seq(`xl-core`)

  override def mvnDeps = Seq(
    mvn"org.typelevel::cats-core:2.12.0"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}

object `xl-testkit` extends XLModule {
  override def moduleDeps = Seq(`xl-core`, `xl-ooxml`)

  override def mvnDeps = Seq(
    mvn"org.typelevel::cats-laws:2.12.0",
    mvn"org.scalacheck::scalacheck:1.18.1",
    mvn"org.scalameta::munit:1.0.3",
    mvn"org.typelevel::munit-cats-effect:2.0.0",
    mvn"org.typelevel::scalacheck-effect:1.0.4"
  )

  object test extends ScalaTests with TestModule.Munit {
    override def mvnDeps = Seq(
      mvn"org.scalameta::munit:1.0.3",
      mvn"org.scalacheck::scalacheck:1.18.1"
    )
  }
}
