# Local Outlook build and existing alternatives

Historical local verification from 2026-09-29. The digest below identifies an
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

The alpha exposes `get_status`, `list_stores`, `list_folders`, `search_messages`,
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

The normal-session connection test does not validate all mailbox operations. Test
those against a synthetic mailbox: message and folder variants, shared stores,
offline data, Object Model Guard prompts, pagination while mail moves, restart,
and unread-state preservation. No personal message content was used in these tests.

The direct stdio deployment gate failed under the tested AppContainer policy.
An isolated Colossus deployment needs an explicit, authenticated user-session bridge
or another reviewed host integration. This has not been implemented, and the
plugin cannot grant itself access through manifest metadata. The catalog should
not advertise AppContainer compatibility.

The subsequent release pipeline uses Apache-2.0, includes dependency notices,
records the execution limitations, and validates with Colossus 0.11.4. It signs
OCI manifests and requires a registry round trip before adding a catalog entry.
Those distribution checks do not replace the remaining real-mailbox tests above.

[existing-release]: https://github.com/schirkan/outlook-mcp-server/releases/tag/v0.1.0
[python-com]: https://github.com/Astral0/outlook-com-mcp
[redemption]: https://dimastr.com/redemption/home.htm
