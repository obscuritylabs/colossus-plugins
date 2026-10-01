# Releases and installation

## What runs where

Run the MCP executable on the Windows machine with classic Outlook, as the same
logged-in user, with Outlook already open. Keep the entire portable ZIP together.
The package includes .NET; a target-side SDK, PowerShell, and Docker are unnecessary.
The executable supports `--version`, `--probe`, and `--stdio`.

This read-only alpha does not support new Outlook, Windows services/session 0, or
direct COM access from the tested Colossus `windows_job` AppContainer. Its OCI
signature verifies publisher/content integrity; it does not grant COM access.
Do not disable isolation as a workaround. Broader mailbox compatibility needs
synthetic-mailbox testing; see [local evidence](local-build.md).

## Retrieve and verify a release

Select a version from [catalog.json](../catalog.json) or
[GitHub releases](https://github.com/obscuritylabs/colossus-plugins/releases).
Use its **complete digest-pinned reference**, not a mutable tag. With ORAS installed:

```powershell
# Replace the placeholder with the full reference from the catalog.
$reference = 'ghcr.io/obscuritylabs/colossus-plugin-outlook-classic@sha256:<manifest-digest>'
$digest = $reference.Split('@')[1]
oras cp --recursive $reference --to-oci-layout './outlook-classic.oci:plugin'
```

The recursive copy retains signatures whether GHCR exposes them through the OCI
referrers API or its tag fallback. Colossus 0.11.4's native registry pull does not
implement the same fallback, so use this ORAS-to-local-layout route. Alternatively,
extract the release's `outlook-classic-oci.tar.gz` into a fresh directory; it
contains the already downloaded OCI layout and signing evidence.

Merge this trust profile into an existing initialized Colossus configuration:

```yaml
plugins:
  trustProfiles:
    obscuritylabs-github:
      mode: required
      identities:
        - issuer: https://token.actions.githubusercontent.com
          subject: https://github.com/obscuritylabs/colossus-plugins/.github/workflows/plugins.yml@refs/heads/main
```

Then verify and install **disabled**:

```powershell
colossus plugins verify ./outlook-classic.oci --digest $digest --trust-profile obscuritylabs-github
colossus plugins install --layout ./outlook-classic.oci --digest $digest --trust-profile obscuritylabs-github
```

Activation is an explicit operator decision under a compatible execution policy.
Enable only an accepted digest, then enable `outlook-classic/mail` with an explicit
tool allowlist. The repository catalog cannot perform these actions automatically.

The portable ZIP is a convenience distribution. Verify its checksum file with
Cosign before comparing the ZIP's SHA-256:

```powershell
cosign verify-blob --bundle SHA256SUMS.sigstore.json --certificate-identity 'https://github.com/obscuritylabs/colossus-plugins/.github/workflows/plugins.yml@refs/heads/main' --certificate-oidc-issuer 'https://token.actions.githubusercontent.com' SHA256SUMS
Get-FileHash ./outlook-classic-windows-amd64.zip -Algorithm SHA256
```

OCI and checksum signatures use GitHub OIDC/Sigstore. The Windows executable is
not Authenticode-signed. For disconnected verification, provision the operator's
trusted Sigstore root and retain the bundles; the OCI layout includes them.

## CI/CD

The [`plugins.yml`](../.github/workflows/plugins.yml) workflow:

1. On pull requests and main pushes, validates catalog/schema invariants, runs STA
   component and MCP protocol checks, and builds Windows x64 with pinned .NET and
   locked NuGet dependencies. It stages licenses and dependency notices.
2. Uses checksum-pinned Colossus 0.11.4 to validate and package the plugin twice,
   requiring identical manifest digests.
3. On `main` only, pushes to GHCR using `GITHUB_TOKEN` with `packages: write`.
   GitHub OIDC signs the exact manifest bytes; ORAS attaches the standard bundle.
4. Pulls the package and signature into a fresh layout and verifies the exact
   digest using Colossus's required signer profile. It signs release checksums too.
5. Adds the verified digest to the catalog and publishes a GitHub prerelease with
   portable ZIP, signed OCI layout, checksums, signatures, and catalog snapshot.

Actions are pinned by commit. No Docker Hub credential or long-lived signing key
is needed. Publishing needs repository Contents write, Packages write, and OIDC
permissions. New GHCR packages may initially be private: an organization/package
administrator must set the package visibility to public for anonymous downloads.

The source manifest supplies the base version. Each build appends
`.ci.<run-number>.<attempt>` to a prerelease (or `-ci...` to a stable base), keeping
every automatic build in the **preview** channel. Full workflow reruns allocate a
new attempt/version; never overwrite existing version tags. Use **Re-run all
jobs**, not just a failed publishing job. Manual dispatch on `main` also publishes.
Promotion to a stable release is separate future work.

Main publications are serialized and catalog-only commits do not recurse. If a
human changes main during the final catalog commit, a non-fast-forward push fails
instead of overwriting their work. Retry the whole workflow after resolving the
change. Consumers remain on their explicitly installed version until they choose
an update.

To add another plugin, give it its own package/src/tests directory and GHCR
repository, add its platform build job, and extend the catalog generator with its
compatibility requirements. Share the verified-publication pattern; do not publish
unbuilt catalog entries.
