# Release Preparation

Prepare the codebase for release version: $ARGUMENTS

## Instructions

Update all version references from current SNAPSHOT/version to the new release version.

### Files to Update

> **Line numbers drift between releases.** Always re-confirm with grep before editing — the
> authoritative list of every version location is:
> ```bash
> grep -rln "<OLD_VERSION>" --include="*.scala" --include="*.mill" --include="*.json" --include="*.md" . \
>   | grep -v "out/" | grep -v ".scala-build" | grep -v CHANGELOG.md
> ```

1. **`build.mill`** (line 17)
   - Find: `val version: String = sys.env.getOrElse("PUBLISH_VERSION", "...")`
   - Update the fallback version string

2. **`xl-core/src/com/tjclp/xl/workbooks/WorkbookMetadata.scala`** (line 19)
   - Update `appVersion` default value: `appVersion: Option[String] = Some("...")`

3. **`plugin/.claude-plugin/plugin.json`** (the `"version"` field)
   - Update the plugin marketplace version. **Do not skip this** — it has drifted in past
     releases (was stale at 0.7.0) because it was missing from this list.

4. **`examples/project.scala`** (line 5)
   - Update the umbrella `//> using dep com.tjclp::xl:...` line (single dependency, not 4)

5. **`README.md`**
   - Update `//> using dep` line (~line 11)
   - Update `ivyDeps` line (~line 58) and the commented per-module `ivy"..."` examples (~lines 61-64)

6. **`docs/QUICK-START.md`**
   - Update Mill `ivyDeps` (~line 15), sbt `libraryDependencies` (~line 22), and `//> using dep` (~line 27)

7. **`examples/README.md`** (~line 144)
   - Update the `com.tjclp::xl:...` reference in prose

8. **`plugin/skills/xl-scripting/SKILL.md` + `plugin/skills/xl-scripting/reference/RECIPES.md`**
   - Update every `//> using dep com.tjclp::xl:...` line (the recipe headers are intentionally
     byte-identical, so a single find-replace covers all of them). The release workflow has a
     gate that **fails the release** if these pins don't match the tag.
   - After bumping, run `./scripts/verify-skill-snippets.sh --local` — this also catches "new
     version breaks documented patterns" before tagging.

### Files to SKIP

- **`plugin/skills/xl-cli/SKILL.md`** - Auto-detects latest release from GitHub API (no version to update)
- **`docs/RELEASING.md`** - Contains example version strings for documentation
- **`CHANGELOG.md`** - Historical version headings; managed separately during release notes

### Verification Steps

After updating all files:

1. Run `./mill __.compile` to verify compilation
2. Run `./mill __.test` to verify tests pass
3. Run `./scripts/test-examples.sh` (version drift guard + examples against the local build)
4. Run `./scripts/verify-skill-snippets.sh --local` (xl-scripting skill snippets compile)
5. Run this to verify no SNAPSHOT refs remain:
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
