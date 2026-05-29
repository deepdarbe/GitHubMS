---
description: Authenticate to GitHub and verify access to this repository
argument-hint: "[org/repo]"
allowed-tools: Bash, Read
---

Help the user authenticate to GitHub and confirm access.

Steps:

1. Check for the GitHub CLI: `gh --version`. If missing, point the user to
   https://cli.github.com/ and stop.
2. Show current auth status: `gh auth status`. If not logged in, run
   `gh auth login` and explain the interactive prompts (HTTPS vs SSH,
   browser vs token).
3. Verify the active identity: `gh api user --jq .login`.
4. Confirm access to the target repository (default to the repo derived from
   `git remote get-url origin` if "$ARGUMENTS" is empty):
   `gh repo view "$ARGUMENTS"`.
5. Report the authenticated user and whether they can read/write the repo.

If `gh` is unavailable, fall back to verifying the git remote and credentials
with `git remote -v` and `git ls-remote origin -h HEAD`.
