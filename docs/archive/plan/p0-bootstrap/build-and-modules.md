
# Build & Modules — sbt, CI, Flags

## Modules
```
xl-core
xl-ooxml        (pure printers/parsers)
xl-cats-effect  (IO interpreters + streaming)
xl-evaluator    (optional pure evaluator)
xl-testkit      (laws, generators, goldens)
docs            (mdoc site, optional)
```

## Scala & flags
- Scala 3.7.x
- `-deprecation -feature -unchecked -Xfatal-warnings`
- `-Yretain-trees` (for better macro errors), optimize at release.

## CI
- JVM matrix (temurin 17/21), `sbt +test`, scoverage, artifacts.
- (Later) MiMa once API stabilizes (≥ 0.2.x).
