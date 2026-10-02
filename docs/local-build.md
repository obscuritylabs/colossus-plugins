# Local Outlook build and existing alternatives

## Live acceptance on 2026-10-02

The `0.1.0-alpha.4` Windows x64 build passed all 24 local live checks against
classic Outlook `16.0.0.20326` in the interactive user session. The
[sanitized report](../validation/outlook-classic-live.json) binds the tested
source and runtime and contains no personal message content. The MCP protocol
test confirmed all 14 tool names, read/write and destructive annotations,
argument validation, and the absence of any send tool.

The real Outlook suite listed mail directly, resolved standard folders, created
and edited one unsent draft in a synthetic PST, changed its read state, moved it,
archived it, and moved it to Deleted Items. It rejected a wrong source folder,
cross-store move, draft update outside Drafts, and permanent deletion. The runner
then removed only that exact test-created item from the synthetic PST. Independent
COM snapshots showed the original fixture counts, bodies, unread flags, and
modification times unchanged. Both test PSTs detached cleanly. The suite also
performed bounded reads of two real Inbox messages without changing their unread
state; no personal mail was written or logged.

The real Colossus 0.11.4 / `gpt-6-luna` agent also tested the new source skill and
the local alpha.4 MCP executable against the synthetic PST. It used `list_messages`
for a plain listing, reported all three expected fixture subjects, read a bounded
body, and left the fixture unchanged. The run used a test-only explicit MCP
registration and the already established interactive-user test policy; it did not
install a signed alpha.4 package or test automatic plugin MCP discovery.

The tested Colossus `windows_job` AppContainer still cannot attach to
Outlook COM, and automatic plugin MCP registration in Colossus 0.11.4 is still
blocked by the host issue described in [installation](releases.md).

## Live acceptance on 2026-10-01

The `0.1.0-alpha.3` Windows x64 build passed all 18 live acceptance checks against
classic Outlook `16.0.0.20326` in the interactive user session. The
[sanitized report](../validation/outlook-classic-live.json) records the source
fingerprint, complete runtime hash, machine/Office versions, and each result.

The suite exercised all six MCP tools through the actual executable: two local
synthetic PSTs, 502 bulk messages, nested and empty folders, non-mail items,
pagination, literal filters, Unicode text and attachment names, bounded bodies,
invalid handles, and process restart. It also read bounded content from two real
unread Inbox messages; both stayed unread. Personal message content was not logged
or saved in the report. Independent COM snapshots confirmed unchanged fixture
counts, body hashes, unread flags, and modification times. Test PSTs were detached,
and Outlook remained running with its original 16 stores. The corrected fixture
builder was also rerun from scratch; all 18 checks and cleanup passed again.

Live testing found a UTF-16 truncation defect in alpha.1: a character limit could
split an emoji and return a replacement character. Alpha.2 preserves surrogate
pairs and passes both a component regression and the live case.

A real Colossus 0.11.4 / `gpt-6-luna` agent also exercised the signed alpha.2
distribution. Automatic plugin MCP discovery failed because Colossus injects
reserved plugin environment variables and then rejects them. The selected skill
worked; explicit MCP registration reached Outlook and searched synthetic mail.
The model altered the long base64 handles twice, preventing completion. Alpha.3
replaces them with short stateless store fingerprints and item identifiers plus a
checksum. Component tests and the repeated live suite cover altered handles,
unknown stores, cross-store resolution, and restart. These checks do not fix the
host's automatic-discovery defect or establish AppContainer compatibility.

During initial fixture setup, Outlook placed 506 synthetic unsent messages in the
default Drafts folder. All 506 were identified and moved into the local test PSTs;
a fresh check found none remaining in Drafts. None were sent. The fixture builder
now moves each generated item to the requested PST folder and verifies its parent
before saving. These fixture writes are separate from the read-only MCP server.

See [re-running the acceptance suite](releases.md#re-run-the-local-outlook-acceptance-suite).
The CI publication gate rejects stale evidence when runtime, dependency, build,
or test sources change. It does not claim hosted CI has a real Outlook profile.

The alpha.2 binary was also retested through Colossus 0.11.4's `windows_job`
AppContainer on October 1 ([isolation result](../validation/outlook-classic-isolation.json)).
The process launched but COM attachment still failed
with `0x800401E3`, without a timeout or output truncation. That deployment remains
unsupported; the passing interactive suite does not override this failed gate.

## Earlier connection and packaging verification

Historical verification from 2026-09-29. The digest below identifies an
unsigned development build, not a published release or a completed Colossus
deployment. Current publication and installation details are in [releases.md](releases.md).

## What is available on this machine

- Classic Outlook x64 is installed and running. Office reports build
  `16.0.20326.20158`; the Outlook object model reports `16.0.0.20326`.
- The system had .NET 8 runtime 8.0.5 but no SDK. Microsoft .NET SDK `10.0.401`
  was downloaded into the ignored `.local/toolchains/dotnet/` directory and its
  SHA-512 checked against Microsoft's release metadata. No system .NET installation
  or PATH change was needed.
- The server uses the official `ModelContextProtocol` C# SDK `2.2.0`, .NET hosting
  `10.0.12`, and the Windows x64 self-contained runtime. Package lock files are
  committed with the source for reproducible dependency resolution.

## Results

| Check | Result |
| --- | --- |
| Release compilation | Passed; zero warnings/errors |
| Handle/input checks | Passed; malformed, oversized, wrong-kind handles and out-of-range pagination rejected |
| COM dispatch checks without Outlook | Passed; one STA, active message pump, concurrent-call rejection, cancellation prevents reuse |
| Published EXE MCP protocol | Passed; initialization, ping, six read-only tools, required arguments, Unicode invalid input, bounded inputs, no send tool |
| Real Outlook connection in normal user session | Passed; STA and profile store collection accessible; no message content requested |
| Same connection in Codex's restricted process | Failed to attach, HRESULT `0x800401E3` |
| Published EXE through Colossus `windows_job` | Process launched; connection failed with `0x800401E3`; no timeout or truncation |
| Colossus package validation | Passed with an empty diagnostics list |
| OCI packaging | Passed using the adjacent Colossus CLI |

The historical read-only build exposed `get_status`, `list_stores`, `list_folders`, `search_messages`,
`get_message`, and `list_attachments`. It has no send, draft, move, delete, attachment
download, or mark-read implementation. It does not start or quit Outlook, modify
Office security settings, or write mailbox content to logs.

The self-contained output includes runtime DLLs: retain the whole `bin/` directory
when copying the EXE. Build locations are printed by `scripts/Build-OutlookClassic.ps1`
and recorded in `.local/last-outlook-build.json`. The OCI layout sits beside the
staged `outlook-classic/` directory after packaging. Use the digest in `index.json`
as the artifact identity; this original build had no tag or catalog entry.

Initial local build `afe66f96d7d4` OCI manifest digest:

```text
sha256:e8b61dc3c4e9468cbc9fae72a1d7a8132ad9dd772c9f94508ebffd279b9e4af2
```

Its portable ZIP is 52,758,273 bytes. This identifies this particular unsigned
build; later compiler or dependency changes can produce a different digest.

Colossus validation used a new `ColossusPluginValidation-*` home beneath
`LocalAppData`, recorded in `.local/colossus-validation-home-path.txt`. Workspace
and Temp ACLs include Codex access grants and were rejected by Colossus's private
storage checks. The separate validation home avoided altering existing Colossus
settings or those ACLs. The AppContainer probe used its own configuration in
`.local/outlook-appcontainer-probe.yaml`.

## Existing implementations

An existing downloadable binary is available from
[schirkan/outlook-mcp-server v0.1.0][existing-release]: `outlook-mcp-server.zip`,
13,693,111 bytes, published 2026-07-20. GitHub's release metadata reports SHA-256
`ab30751425e4839199f75ebd24e66b1acadcfab9799a8b3f11803ae7e3357f4d`.
The project is C#/.NET with Outlook COM interop and an MIT license. It exposes
both reading and mailbox/calendar writes. The release metadata and source were
inspected; its binary was not downloaded, executed, or incorporated here.

Its adapter uses a semaphore to serialize COM calls. Serialization alone does not
establish STA thread affinity, so that implementation would need further review
before reuse. Our alpha uses the official MCP SDK but contains its own small,
read-only Outlook adapter and an explicit STA message pump.

[Astral0/outlook-com-mcp][python-com] is another COM-based option, distributed from
source using Python and pywin32. [Redemption][redemption] is an existing native COM /
Extended MAPI library; it is a component for building integrations, not a drop-in
Colossus plugin or MCP server.

## Remaining acceptance work

The October 1 suite supersedes the earlier connection-only result for the covered
read operations. Shared-mailbox permission models, offline/uncached mail,
protected items, Object Model Guard prompts, missing profiles, and pagination while
mail moves still need separate acceptance scenarios. Other Office builds and
32-bit Office are unvalidated. Passing on this machine is not production approval
for every legacy environment.

The direct stdio deployment gate failed under the tested AppContainer policy.
An isolated Colossus deployment needs an explicit, authenticated user-session bridge
or another reviewed host integration. This has not been implemented, and the
plugin cannot grant itself access through manifest metadata. The catalog should
not advertise AppContainer compatibility.

The subsequent release pipeline uses Apache-2.0, includes dependency notices,
records the execution limitations, and validates with Colossus 0.11.4. It signs
OCI manifests and requires a registry round trip before adding a catalog entry.
Those distribution checks do not replace the remaining compatibility tests above.

[existing-release]: https://github.com/schirkan/outlook-mcp-server/releases/tag/v0.1.0
[python-com]: https://github.com/Astral0/outlook-com-mcp
[redemption]: https://dimastr.com/redemption/home.htm
