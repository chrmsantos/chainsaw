# Repository Maintenance Guide

## Commit History Management

### Current Status (2025-09-24)
- **Total commits:** 2
- **Oldest commit:** 2025-09-23 (2 days ago)
- **Latest commit:** 2025-09-24 (1 day ago)
- **Commits older than 30 days:** None

### Commit History Cleanup

This repository currently has no commits older than 30 days. All commits are recent and should be preserved.

#### Future Maintenance Script

When this repository accumulates commits older than 30 days, you can use the following approach:

**⚠️ WARNING: These operations are destructive and will permanently remove commit history. Always create backups before proceeding.**

1. **Check for old commits:**
   ```bash
   # Find commits older than 30 days
   git log --until="$(date -d '30 days ago' --iso-8601)" --oneline
   ```

2. **Create a backup:**
   ```bash
   # Create a backup branch
   git branch backup-$(date +%Y%m%d) HEAD
   ```

3. **If you need to squash or remove old commits:**
   ```bash
   # Find the first commit to keep (30 days ago)
   CUTOFF_DATE=$(date -d '30 days ago' --iso-8601)
   FIRST_COMMIT_TO_KEEP=$(git log --since="$CUTOFF_DATE" --format="%H" | tail -1)
   
   # Interactive rebase to squash old commits (if any exist)
   git rebase -i $FIRST_COMMIT_TO_KEEP^
   ```

### Repository Statistics
- Created: 2025-09-23
- Language: VBA (Microsoft Word macros)
- Purpose: Document standardization system for legislative documents
- License: Apache 2.0 (modified)

### Notes
- This is a small, focused project with minimal commit history
- Regular cleanup may not be necessary due to the project's scope
- Consider the impact on collaborators before removing commit history
- GitHub Issues and Pull Requests reference commits by SHA - removing commits may break these references