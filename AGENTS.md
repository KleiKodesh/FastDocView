# FastDocView: instructions for agents

Read [CLAUDE.md](CLAUDE.md) first: it holds every rule for working in this repo (windowed data,
verifying against sources and LibreOffice, QA before each commit, worktrees, commits).

## Keep the feature lists current

Every change that adds, completes, removes or limits a feature updates both lists in the same
commit:

- [docs/SUPPORTED-FEATURES.md](docs/SUPPORTED-FEATURES.md): what is shown, per format.
- [docs/UNSUPPORTED-FEATURES.md](docs/UNSUPPORTED-FEATURES.md): what is not, with the reason and
  its status (planned, deferred or won't do).

Move the item from one list to the other, fix its per-format cells, and keep the reason for
anything still unsupported. A feature commit that leaves the lists stale is not finished.
