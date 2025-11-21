# Pull Request

## Description
<!-- Brief description of changes -->

## Type of Change
- [ ] Bug fix (non-breaking change fixing an issue)
- [ ] New feature (non-breaking change adding functionality)
- [ ] Performance improvement (faster/smaller with no API changes)
- [ ] Breaking change (fix or feature changing existing API)
- [ ] Refactoring (code restructure with no behavioral changes)
- [ ] Documentation only

## Testing
- [ ] Unit tests pass: `./mill <module>.test`
- [ ] All tests pass: `./mill __.test`
- [ ] Code formatted: `./mill __.reformat`
- [ ] Zero WartRemover warnings

## Documentation Updates

**If this PR implements a work item from roadmap.md, complete this checklist**:

- [ ] Work item(s) marked ✅ in plan doc (`docs/plan/*.md`)
- [ ] `roadmap.md` updated:
  - [ ] Work Items Table status changed to ✅ Complete
  - [ ] Mermaid DAG node class changed to `completed`
  - [ ] TL;DR updated if new work unblocked
  - [ ] Downstream dependencies recalculated (blocked → available)
- [ ] `STATUS.md` updated if new capability added:
  - [ ] Added to "What Works" section
  - [ ] Test count updated
  - [ ] Performance metrics updated (if applicable)
- [ ] `LIMITATIONS.md` updated if limitation removed
- [ ] `CLAUDE.md` updated if API surface changed
- [ ] Examples added to `docs/reference/examples.md` (if user-facing feature)

**If this PR completes an entire phase/plan**:

- [ ] Plan doc archived (moved to git history)
- [ ] Historical note added to roadmap.md

## Related Issues

<!-- Link to issues, work items, or plan docs -->

- Closes #<issue>
- Implements work item: `<WI-XX>` from `docs/plan/roadmap.md`
- Related to: `docs/plan/<plan-file>.md`

## Additional Context

<!-- Any additional information reviewers should know -->
