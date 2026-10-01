# Outlook Classic plugin

Status: read-only alpha, base version `0.1.0-alpha.2`; CI appends a unique preview build suffix. Target the legacy environment: Windows 11 with
**classic Outlook for Windows**, a configured mail profile, and a signed-in
interactive user. The proposed plugin ID is `outlook-classic`, with MCP server ID
`outlook-classic/mail` and skill ID `outlook-classic/mail`.

Microsoft lists the Outlook Object Model as available in classic Outlook and
unsupported in new Outlook ([feature comparison][outlook-comparison]). New Outlook
and Microsoft Graph are outside this plugin's initial scope. A compiled binary
does not make COM available in new Outlook.

## Build and run

From the repository root, run `./scripts/Build-OutlookClassic.ps1`. It uses the
pinned .NET SDK, runs component and MCP protocol checks, and publishes a
self-contained Windows x64 executable with its runtime files. Node.js is needed
for the protocol check, not to run the published plugin. Keep the entire `bin/`
directory together; this is not a single-file distribution.

The script prints the executable and staged package paths and saves a local
receipt at `.local/last-outlook-build.json`:

```powershell
$build = Get-Content -Raw .local/last-outlook-build.json | ConvertFrom-Json
& $build.executable --probe    # Version, connection, apartment, and store count only
& $build.executable --stdio    # MCP transport; do not type ordinary text here
```

All six read-only tools in the table below are implemented. Discovery never
connects to Outlook; only a tool call or `--probe` attaches. Outlook must already
be open, and the process must run in an execution context that can access its
running COM object. The server never starts Outlook or calls `Quit`.

The connection probe passed on this machine in the normal user session. It failed
inside the tested Colossus `windows_job` boundary with HRESULT `0x800401E3`.
Consequently direct stdio inside that boundary is not currently a supported
deployment. Do not disable isolation as an installation workaround. See the
[build report](../../docs/local-build.md) for test scope and the bridge follow-up.

Main-branch CI publishes signed OCI previews to GHCR and adds verified digests to
the catalog. The Windows executable is not Authenticode-signed. All 18 local live
checks passed on 2026-10-01, covering two synthetic PSTs and bounded reads of two
real unread Inbox messages. Publication requires evidence matching the current
sources. Broader Office and mailbox compatibility remains unvalidated. See the
[release and installation guide](../../docs/releases.md).

With an existing configured Colossus CLI, package the staged directory:

```powershell
colossus plugins validate $build.packageRoot
$layout = Join-Path (Split-Path -Parent $build.packageRoot) 'outlook-classic.oci'
colossus plugins package $build.packageRoot --output $layout
```

The portable ZIP is also produced beside the staged directory. An OCI layout is
an additional distribution format; neither packaging nor unzipping enables tools.

## Implementation rationale

PowerShell can be used for read-only feasibility experiments. It can instantiate
`Outlook.Application` through COM; a compiled binary is not required merely to
access the object model. Run explicitly in STA mode and return structured output.
The experiment should inspect Outlook/profile availability and enumerate a small
amount of synthetic test data, without sending or modifying mail.

The distributable MCP server is a self-contained C#/.NET executable:
typed request validation, bounded JSON output, explicit error handling, and one
controlled COM apartment are easier to maintain than a large PowerShell protocol
server. The build pins the SDK and dependency lock files. It uses normal
COM interop; NativeAOT, trimming, Extended MAPI, and third-party COM libraries are
not initial requirements. Ship every required runtime file, whether publishing
as one executable or a directory.

All Outlook operations must execute on the designated STA thread with a message
pump, including COM object creation and release. Marshal requests onto that thread;
do not send Outlook COM objects through thread-pool continuations. Microsoft
documents STA-only object-model access and no Windows Service support
([API selection guidance][outlook-api]). The process runs as the same interactive
user as Outlook. A container or session-0 service is not the deployment target.

## First gate: prove the Colossus execution boundary

Colossus's `windows_job` backend combines AppContainer and Job Object isolation.
Its Windows process implementation currently supplies no AppContainer capabilities.
Being able to use COM in an ordinary PowerShell window does not establish that
a plugin process can reach Outlook, its COM activation endpoints, or its profile.
See the inspected [sandbox contract][colossus-sandbox].

Run the same small probe under the intended Colossus policy and test:

- Outlook already running versus not running; avoid starting it as an unobserved
  side effect of discovery. The initial supported flow can require it to be open.
- COM activation/attachment, profile access, and bounded folder enumeration.
- Timeout, modal prompt, process cleanup, and user-owned Outlook process survival.
- The actual Office build and architecture. Begin with Windows x64; do not infer
  either support or incompatibility for 32-bit Office without an interop test.

If direct stdio execution cannot work within the approved boundary, explicitly
design a user-session bridge and its Colossus integration. A possible bridge is a
separately managed local Streamable HTTP MCP service, with authenticated loopback
access and constrained mailbox operations. Its installer, credential provisioning,
origin/Host checks, session ownership, lifecycle, and sandbox network reachability
would need their own implementation and tests. It must not expose generic shell
or COM execution. Do not silently weaken isolation or claim the current plugin
manifest can grant COM access. This fallback is a design option, not an existing
Colossus capability proven here.

## Direct-stdio package

The staged package's `mcp.json` contains:

```json
{
  "$schema": "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json",
  "mcpServers": {
    "mail": {
      "type": "stdio",
      "command": "./bin/outlook-classic-mcp.exe",
      "args": ["--stdio"],
      "cwd": "${PLUGIN_ROOT}"
    }
  }
}
```

The source manifest is under `package/`; the build adds the executable to the
staged package. This alpha persists no mailbox state or attachments. Plugin files
remain immutable; future deliberate state belongs in `${PLUGIN_DATA}`.
Never bundle a PST/OST, mailbox credentials,
personal mail samples, or the user's Outlook profile.

The current Colossus MCP adapter initializes a fresh transport for each discovery
page and tool call. Do not rely on a permanently running process, in-memory paging
state, or a previous call's COM object. The implementation reattaches per operation,
uses opaque store/item handles, and exposes bounded numeric offsets. Those offsets
are best-effort positions, not snapshot cursors. Keep stdout
strictly MCP protocol data and send redacted diagnostics to stderr.

## Initial tool surface

Start with a small read-only release and add writes in a separately reviewed
release. Every call has bounded input, time, result count, and output size.

| Implemented tool | Behavior |
| --- | --- |
| `get_status` | Report client/profile/session compatibility and actionable failures |
| `list_stores` | Enumerate mail stores visible in the current profile |
| `list_folders` | Enumerate bounded folder metadata in an explicit store |
| `search_messages` | Search an explicit folder with structured filters, bounded results, and pagination |
| `get_message` | Return selected fields and bounded plain-text content for an explicit message |
| `list_attachments` | Return metadata; do not save or open files implicitly |

The implementation uses `StoreID` plus `EntryID` identifiers inside opaque handles. Resolve
them afresh and report missing or moved items; do not promise identity survives a
move. Restrict initial content operations to mail items and preserve unread state.
Escape structured filter values instead of accepting arbitrary Outlook query or
PowerShell expressions. Do not return COM objects or unrestricted property bags.
Show only data the current Outlook profile already permits, including shared
stores where available. A local adapter does not guarantee offline completeness:
uncached mail and server searches can depend on Outlook's existing connection.

Later phases may add bounded attachment export, draft creation/update, and
explicit send/reply/move operations. Draft creation already writes to the mailbox.
Before enabling writes, establish concrete recipient/content approval through
Colossus policy, tested outcome handling, and a durable duplicate-prevention design.
A model-supplied `approved: true` argument is not authorization. After a timeout or
uncertain send outcome, report uncertainty and reconcile; never blindly retry send.

Colossus must separately enable the installed plugin digest and configure
`plugins.mcpServers["outlook-classic/mail"]` with `enabled: true` and an exact
`allowedTools` list. MCP annotations and skill instructions describe behavior;
they do not grant authority or replace server-side validation.

## Legacy-environment behavior

Object Model Guard can display prompts for protected data or operations. Preserve
those controls and report blocked access; neither an executable nor code signing
guarantees prompt-free automation ([Microsoft security guidance][outlook-security]).
Never change Trust Center, antivirus, or registry security settings automatically.

Treat mail and attachment content as untrusted data. Do not follow instructions
embedded in messages. Do not render active HTML, fetch tracking resources, launch
attachments, or include message bodies in diagnostic logs. Any future attachment
export needs a validated filename and an explicitly permitted destination.

COM calls can stall behind Outlook UI. A timed-out request must fail clearly;
cancelling an MCP request does not prove Outlook cancelled the underlying operation.
Release owned COM references without calling `Quit` on the user's Outlook instance.
The integration tests must check that Colossus's process cleanup does not terminate
the user's existing Outlook session.

## Validation and production-readiness gates

1. Record actual Windows/Office versions, bitness, profile type, and effective
   Colossus policy; demonstrate safe COM access under that policy.
2. Exercise MCP initialization, tools/list, tools/call, malformed input, Unicode,
   result limits, timeout, cancellation, and process restart against fixtures.
3. Test real COM against a dedicated synthetic mailbox in an interactive session:
   multiple stores, missing profile, denied access, offline/uncached items,
   non-mail items, and large folders. Verify no unintended mailbox mutation.
4. Validate/package the staged plugin with Colossus; pull and verify its signed
   OCI artifact; install disabled, then explicitly enable the digest and tools.
5. CI publishes only verified alpha artifacts, with the unvalidated mailbox and
   isolation limits recorded in the catalog. Complete the real-mailbox and
   execution-policy gates before promoting to production. Do not claim support
   for untested Office architectures or new Outlook.

[outlook-comparison]: https://support.microsoft.com/en-us/outlook/getstarted/feature-comparison-between-new-outlook-and-classic-outlook
[outlook-api]: https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook
[outlook-security]: https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/outlook-object-model-security-warnings
[colossus-sandbox]: https://github.com/obscuritylabs/Colossus/blob/d0cb54fa03396d45f93222447414474658abf6db/docs/admin/sandbox.md
