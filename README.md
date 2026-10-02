# Colossus Plugins

Plugins for [Colossus](https://github.com/obscuritylabs/Colossus), using
[Agent Plugins v1](https://agent-plugins.org/specification) and OCI distribution
through **GitHub Container Registry**. Licensed under [Apache-2.0](LICENSE).

The first plugin is **Outlook Classic**, a self-contained Windows x64 C# MCP
server with 14 bounded MCP tools for listing, searching, reading, moving, archiving,
soft deletion, read-state changes, and unsent drafts. It attaches to classic Outlook already
running in the same logged-in Windows user session. The target machine does not
need Docker, PowerShell, or a separately installed .NET runtime.

This is an alpha. All 24 live checks passed locally in the normal user session,
including writes confined to synthetic PSTs and bounded reads of real Inbox messages.
The tested Colossus `windows_job`
AppContainer could not attach. Broader mailbox and Office compatibility remains
unvalidated. See [Outlook Classic](plugins/outlook-classic/README.md)
and the [local test report](docs/local-build.md).

## Distribution

- [GitHub prereleases](https://github.com/obscuritylabs/colossus-plugins/releases): portable ZIP, verified OCI layout, checksums, and signatures.
- GHCR repository: `ghcr.io/obscuritylabs/colossus-plugin-outlook-classic`.
- [Catalog](catalog.json): published versions pinned to immutable OCI manifest digests.
- [Install and verify](docs/releases.md): registry retrieval, signer identity, and CI/CD behavior.

Pull requests run catalog, Windows component, MCP protocol, and packaging checks.
Pushes to `main` publish a new preview version, sign it with GitHub OIDC, verify a
registry round trip with Colossus, and update the catalog. Catalog-only commits do
not trigger another release. Publication does not auto-update installed plugins.
Publication also requires [passing live Outlook evidence](validation/outlook-classic-live.json)
matching the current runtime and test sources; a code change invalidates stale evidence.

## Repository structure

```text
catalog.json                          # Generated release discovery index
schemas/catalog.schema.json           # Catalog contract
plugins/outlook-classic/
  package/plugin.json                 # Portable Agent Plugins manifest
  package/mcp.json                     # MCP transport
  package/skills/mail/SKILL.md          # Agent instructions
  src/                                # Windows COM MCP implementation
  tests/                              # Component, protocol, and opt-in live checks
scripts/                              # Build, package, publish, catalog tooling
validation/outlook-classic-live.json   # Sanitized local acceptance evidence
.github/workflows/plugins.yml         # PR validation and main publication
docs/                                 # Architecture, release, and test details
```

Each future plugin gets its own source/package/tests directory and independently
versioned OCI repository. The workflow currently builds the first Windows plugin;
additional plugins need their build job and verified release metadata added.

The repository catalog is a discovery contract. Colossus's current installed-plugin
catalog does not yet load this remote index; that consumer is separate work.

## Local build

Install the SDK from `global.json` and Node.js, then run:

```powershell
npm ci --ignore-scripts
npm test
npm run validate
./scripts/Build-OutlookClassic.ps1
```

The build prints package and executable paths and records them in
`.local/last-outlook-build.json`. Keep the entire published `bin/` directory
together. `--probe` checks the Outlook connection without reading messages;
`--stdio` starts the MCP server.

See [architecture](docs/architecture.md) for package, catalog, trust, and runtime responsibilities.
