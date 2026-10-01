# Repository, catalog, and release architecture

The repository owns plugin source, portable packages, OCI publication, and a
discovery catalog. Colossus owns installation, trust policy, activation, MCP tool
permissions, and execution isolation. The CI packager/verifier is pinned to
Colossus 0.11.4.

## Package boundaries

Each plugin lives under `plugins/<name>/`, with `package/`, `src/`, and `tests/`
directories. The portable manifest is **root `plugin.json`**, with root `mcp.json`
and `skills/<name>/SKILL.md` discovered by convention. This is Agent Plugins v1,
not the Codex `.codex-plugin/` format.

The build stages an allowlisted package under ignored `dist/`, adding the
self-contained runtime, Apache license, upstream notices, and dependency lock.
The manifest supplies the base version; CI derives a unique preview version and
uses it in both the staged manifest and executable. The complete directory is
input to `colossus plugins validate` and `colossus plugins package`.

```text
outlook-classic/
  plugin.json
  mcp.json
  skills/mail/SKILL.md
  bin/outlook-classic-mcp.exe           # Keep adjacent runtime DLLs
  README.md
  LICENSE
  NOTICE
  THIRD-PARTY-NOTICES.txt
  licenses/
  dependencies.lock.json
```

## OCI profile and signatures

| Field | Value |
| --- | --- |
| Manifest | Standard OCI image manifest; one plugin/platform |
| Artifact type | `application/vnd.colossus.agent-plugin.v1` |
| Config type | `application/vnd.colossus.agent-plugin.config.v1+json` |
| Single content layer | `application/vnd.colossus.agent-plugin.content.v1.tar+gzip` |
| Archive root | Exactly `<plugin-name>/` |
| Identity | SHA-256 of the OCI manifest bytes |

Colossus packages sorted paths with normalized archive metadata and deterministic
gzip. CI repeats packaging to check that the same staged bytes produce the same
digest. This does not claim independent compiler reproducibility. Indexes,
multiple payload layers, links, and traversal are not plugin payloads.

ORAS transfers this layout to `ghcr.io/obscuritylabs/colossus-plugin-<name>`.
The plugin runs as a Windows process, not inside a container. A tag identifies a
preview; catalog references always pin the manifest digest.

Colossus 0.11.4 verifies standard Sigstore bundles over the exact manifest bytes.
CI therefore uses `cosign sign-blob`, then attaches the bundle as an OCI referrer.
It recursively retrieves the package and signing evidence with ORAS, then uses a
`required` Colossus trust profile to verify the downloaded layout. ORAS supports
both referrers API and tag fallback; do not assume Colossus's native registry pull
supports both. The documented installation route preserves the evidence in a
local OCI layout. See [release instructions](releases.md).

The GitHub OIDC signer is restricted to this repository's exact main-branch
workflow identity. OCI signing does not provide Windows Authenticode signing.
The current executable is not Authenticode-signed.

## Catalog contract

[`catalog.json`](../catalog.json) is the discovery index, validated against
[`schemas/catalog.schema.json`](../schemas/catalog.schema.json) and semantic
checks in `scripts/catalog.mjs`.

| Record | Meaning |
| --- | --- |
| Catalog | Schema version, stable ID, repository, monotonic revision |
| Plugin | Portable name, display name, description, source path |
| Release | SemVer, channel, timestamp, source commit, minimum tested Colossus version, yanked status |
| Artifact | Immutable OCI reference, execution platform, explicit prerequisites and limitations |

Names and release versions must be unique. Platform pairs must be unique per
release. Stable releases cannot carry prerelease suffixes. Published version to
digest mappings never change. Yanked entries remain for audit. Preview consumers
select the greatest compatible non-yanked SemVer; there is no mutable `latest`
field. A fresh empty catalog is valid until the first verified publication.

The release workflow serializes catalog writes. It adds an entry only after build
checks, signing, registry retrieval, and Colossus trust verification pass. A failed
check leaves the previous catalog intact. Catalog changes do not grant signer
trust, credentials, origins, tool permissions, or activation authority.

The index is available over GitHub HTTPS and attached to releases. Catalog
signatures and freshness/rollback policy remain work for a future discovery
loader; a revision number alone does not authenticate discovery.

## Runtime integration and alpha scope

Colossus's existing internal catalog is a snapshot of installed plugins. A remote
loader for this repository's index has not been implemented. That future loader
must validate bounded discovery data, show compatibility, retain provenance,
resolve to operator-configured trust/registry profiles, and delegate installation
and activation to existing lifecycle APIs.

The Outlook plugin targets Windows 11 x64 and classic Outlook in an interactive
user session. Real local connection probing passed; direct COM inside the tested
`windows_job` AppContainer failed. An authenticated user-session bridge or reviewed
host integration is separate future work. Publishing this alpha does not claim
AppContainer compatibility or full mailbox integration coverage.

Hosted Windows CI exercises fixtures and the real stdio protocol without opening
Outlook. Full COM testing needs a dedicated synthetic mailbox and an interactive
Office session. New Outlook/Graph and mailbox write operations are outside scope.
