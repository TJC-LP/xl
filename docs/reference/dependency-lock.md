# Dependency lock

`deps.lock.json` records the existing Mill/Coursier resolution and SHA-256 of each
Maven artifact. It does not change dependency selection or upgrade any library.
CI's build job runs the checker before formatting, compilation and tests:

```sh
python3 -B scripts/test-dependency-lock.py
python3 -B scripts/dependency_lock.py check
```

Python 3.9+ and the normal `./mill` prerequisites are required. Run from any
directory: the script locates the repository relative to its own file. Mill uses
the pinned toolchain and normal Coursier cache. Checks hash the actual downloaded
bytes with SHA-256, including on a warm cache; they do not trust Mill's cached
PathRef hash or Maven's published checksum.

## Coverage

The script dynamically selects all modules, including nested tests and the
evaluator's perf module, with these Mill 1.1.5 targets:

- `__.resolvedMvnDeps`: compile/provided dependencies, including transitives.
- `__.resolvedRunMvnDeps`: runtime resolution, kept separate from compile resolution.
- `__.scalaCompilerClasspath`, `__.scalacPluginClasspath`,
  `__.scalaDocClasspath`, `__.scalaDocPluginClasspath`, and
  `__.annotationProcessorsResolvedMvnDeps`: compiler, documentation and plugin artifacts.
- `__.prepareOffline`: additional worker/bridge artifacts exposed by Mill's offline
  preparation. The native-image module also returns XL's compiled classes and
  resource directories; those first-party directories are excluded. This target
  can compile XL as part of preparing its native-image inputs.
- `__.generatorDeps`: JMH bytecode generator dependencies.

The first target group also runs at `--meta-level 1` to include dependencies of
the build definition, including its JMH plugin. Target membership and classpath
order are recorded, so moving a dependency between configurations or changing a
single module's version cannot hide behind a repository-wide union. Newly
discovered or removed targets also change the lock.

Artifact identities include group, cross-built artifact name, version, and full
filename (including classifier/extension). The script reads the adjacent POM's
group and checks it against Coursier's Maven cache layout. Cache roots and
repository URLs are omitted from the lock; identical artifacts in different
HTTP(S) repository caches must have identical bytes. Unrecognised references,
missing files/POMs, resolution failures and conflicting bytes fail the check.

This is a CI comparison gate, not an offline resolver or forced-version policy.
An ordinary `./mill` invocation does not enforce the lock automatically.
Release jobs currently resolve their own artifacts without running this check,
so this gate does not verify the dependency bytes used by native release builds,
assembly bundles or Maven publication. Release enforcement is a separate follow-up. The
gate runs after Mill bootstrap, so it is not a sandbox against executing an
untrusted build or compiler plugin. It does not lock the Mill launcher, JDK
download, GitHub Actions, external scala-cli/example toolchains, or artifacts
downloaded internally by tools rather than exposed through the selected Mill
targets (for example Scalafmt's dynamically loaded formatter). POM/metadata bytes
are not checksummed; changes to their resulting selected artifact tree are
detected. New tooling/configurations must have their resolver targets added
explicitly if they are not covered above. Launcher/action pinning and updater
policy are separate security tasks (#834–#839).

## Deliberate updates

First investigate a failing check. Do not regenerate the lock merely to turn CI
green. When changing a dependency intentionally:

1. Edit the appropriate `mvnDeps` declaration.
2. Run `python3 -B scripts/dependency_lock.py update`.
3. Review the lock diff: intended versions, transitive changes, target membership
   and any changed bytes at an unchanged coordinate. An unchanged coordinate with
   different bytes requires investigation, not routine acceptance.
4. Run `check` and the repository's relevant build/test gates. Commit the dependency
   declaration and reviewed lock together.

`update` only records what the existing build resolves. It does not search for
new releases, edit declarations, or force a shared version across modules.
Failed resolution leaves the existing lock intact. A project lock does not force
downstream consumers of XL's published Maven libraries to use the same graph.

Official references: [Mill dependency configuration and resolved targets](https://mill-build.org/mill/1.1.x/fundamentals/library-deps.html),
[Mill 1.1.5 JavaModule source](https://github.com/com-lihaoyi/mill/blob/1.1.5/libs/javalib/src/mill/javalib/JavaModule.scala),
[Mill 1.1.5 ScalaModule source](https://github.com/com-lihaoyi/mill/blob/1.1.5/libs/scalalib/src/mill/scalalib/ScalaModule.scala).
