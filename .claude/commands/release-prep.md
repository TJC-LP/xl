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

- **`plugin/skills/xl-cli/SKILL.md`** - Auto-detects latest release from GitHub API (no version to update)
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

### Tagging

After committing, create an **annotated tag** with release notes from CHANGELOG.md:

```bash
# Step 1: Extract and preview release notes (run separately to avoid zsh parse issues)
awk '/^## \[0.5.0-RC1\]/{flag=1; next} /^## \[/{flag=0} flag' CHANGELOG.md | sed '/^$/d' | head -50

# Step 2: Create annotated tag with heredoc (replace version and paste notes)
git tag -a "v0.5.0-RC1" -m "$(cat <<'EOF'
<paste release notes here>
EOF
)"

# Step 3: Verify it's annotated (should print "tag", not "commit")
git cat-file -t "v0.5.0-RC1"
```

**Note**: The awk command with `$VERSION` variable substitution causes zsh parse errors. Use literal version strings in each command instead of variable chaining.

**Important**: Do NOT use `git tag v$ARGUMENTS` (without `-a`) - this creates a lightweight tag with no release notes, causing GitHub releases to show the commit message instead.

### Push

Push the commit and tag to trigger the release workflow:

```bash
git push origin main
git push origin "v$VERSION"
```

The release workflow will:
1. Build native binaries for all platforms
2. Publish to Maven Central
3. Create GitHub Release with the tag message as release notes
