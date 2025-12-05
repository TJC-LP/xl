# Release Preparation

Prepare the codebase for release version: $ARGUMENTS

## Instructions

Update all version references from current SNAPSHOT/version to the new release version.

### Files to Update

1. **`build.mill`** (~line 55)
   - Find: `sys.env.getOrElse("PUBLISH_VERSION", "...")`
   - Update the fallback version string

2. **`examples/project.scala`** (lines 5-8)
   - Update all 4 module dependencies: xl-core, xl-ooxml, xl-cats-effect, xl-evaluator

3. **`xl-core/src/com/tjclp/xl/workbooks/WorkbookMetadata.scala`** (~line 12)
   - Update `appVersion` default value

4. **`README.md`**
   - Update `//> using dep` lines (~lines 11-12)
   - Update `ivy"..."` dependencies (~lines 51-53)

5. **`docs/QUICK-START.md`**
   - Update Mill ivyDeps (~lines 21-23)
   - Update sbt libraryDependencies (~lines 33-35)

6. **`examples/README.md`** (~line 126)
   - Update example `//> using dep` line

### Files to SKIP

- **`.claude/skills/xl-cli/SKILL.md`** - Uses `__XL_VERSION__` placeholders that CI replaces at build time
- **`docs/RELEASING.md`** - Contains example version strings for documentation

### Verification Steps

After updating all files:

1. Run `./mill __.compile` to verify compilation
2. Run `./mill __.test` to verify tests pass
3. Run this to verify no SNAPSHOT refs remain:
   ```bash
   grep -r "SNAPSHOT" --include="*.scala" --include="*.mill" --include="*.md" . | grep -v RELEASING.md | grep -v ".scala-build" | grep -v "out/"
   ```

### Commit

When complete, stage and commit with message:
```
chore(release): Bump version to $ARGUMENTS
```
