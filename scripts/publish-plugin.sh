#!/usr/bin/env bash
set -euo pipefail

release_dir="$PWD/dist/release"
layout="$release_dir/outlook-classic.oci"
repository='ghcr.io/obscuritylabs/colossus-plugin-outlook-classic'
identity='https://github.com/obscuritylabs/colossus-plugins/.github/workflows/plugins.yml@refs/heads/main'
issuer='https://token.actions.githubusercontent.com'
version=$(jq -er .version "$release_dir/release.json")
digest=$(jq -er .digest "$release_dir/release.json")
test "$(jq -r .sourceCommit "$release_dir/release.json")" = "$GITHUB_SHA"
test "$(jq -r .checksPassed "$release_dir/release.json")" = true
test "$(jq -r .version "$release_dir/plugin.json")" = "$version"
[[ "$digest" =~ ^sha256:[a-f0-9]{64}$ ]]
tag="$repository:$version-windows-amd64"
reference="$repository@$digest"
printf '%s' "$GH_TOKEN" | oras login ghcr.io --username "$GITHUB_ACTOR" --password-stdin
# Version contains unique Actions run number + attempt. Never overwrite a release tag.
if oras manifest fetch "$tag" >/dev/null 2>&1; then
  echo 'Release tag already exists. Re-run the full workflow to allocate a new version.' >&2
  exit 1
fi
oras cp --from-oci-layout "$layout@$digest" "$tag"
oras manifest fetch "$reference" --output "$release_dir/manifest.json"
echo "${digest#sha256:}  $release_dir/manifest.json" | sha256sum --check

# Colossus verifies the manifest bytes directly, not Cosign's image payload envelope.
cosign sign-blob --yes --bundle "$release_dir/manifest.sigstore.json" "$release_dir/manifest.json"
cosign verify-blob --bundle "$release_dir/manifest.sigstore.json" \
  --certificate-identity "$identity" --certificate-oidc-issuer "$issuer" "$release_dir/manifest.json"
(
  cd "$release_dir"
  oras attach --artifact-type application/vnd.dev.sigstore.bundle.v0.3+json "$reference" \
    manifest.sigstore.json:application/vnd.dev.sigstore.bundle.v0.3+json
)

# Recursive copy includes signatures whether the registry uses referrers API or tag fallback.
verified="$release_dir/verified.oci"
oras cp --recursive "$reference" --to-oci-layout "$verified:plugin"
COLOSSUS_HOME=$(mktemp -d "$HOME/colossus-plugin-verification.XXXXXXXX")
export COLOSSUS_HOME
cp .github/colossus-release.yaml "$COLOSSUS_HOME/config.yaml"
chmod 600 "$COLOSSUS_HOME/config.yaml"
colossus plugins verify "$verified" --digest "$digest" --trust-profile github

# Retain the verified layout for disconnected installation, with signing evidence.
tar -czf "$release_dir/outlook-classic-oci.tar.gz" -C "$verified" .
jq --arg repository "$repository" --arg runUrl "$GITHUB_SERVER_URL/$GITHUB_REPOSITORY/actions/runs/$GITHUB_RUN_ID" \
  '. + {repository: $repository, verified: true, workflowRun: $runUrl}' \
  "$release_dir/release.json" > "$release_dir/release.tmp.json"
mv "$release_dir/release.tmp.json" "$release_dir/release.json"
cp validation/outlook-classic-live.json "$release_dir/outlook-classic-live.json"
(
  cd "$release_dir"
  sha256sum outlook-classic-windows-amd64.zip outlook-classic-oci.tar.gz manifest.json manifest.sigstore.json release.json outlook-classic-live.json > SHA256SUMS
  cosign sign-blob --yes --bundle SHA256SUMS.sigstore.json SHA256SUMS
)
cat > "$release_dir/release-notes.md" <<EOF
Classic Outlook mail MCP alpha for Windows 11 x64, packaged as Agent Plugins v1.

OCI artifact: \`$reference\`

Runs beside classic Outlook in the same logged-in user session. Includes 14 bounded tools for mail listing, reading, read-state changes, same-store move/archive, move to Deleted Items, and unsent draft creation/update. No mail sending, permanent deletion, or attachment downloads. New Outlook is outside scope.

The OCI manifest and checksums are signed with this repository's GitHub Actions identity. CI built the self-contained Windows package, exercised component/MCP/catalog checks, and verified the signed artifact after pulling it from GHCR with Colossus 0.11.4. The Windows EXE is not Authenticode-signed.

Publication required passing local live Outlook evidence matching these sources. The attached outlook-classic-live.json records 24 checks against two synthetic PSTs, synthetic-only writes, and bounded real Inbox reads; no personal message content is included. CI rebuilds the tested sources with this preview version; the local report identifies the separately tested local runtime.

Known alpha limits: Colossus 0.11.4 rejects its injected plugin environment variables, blocking automatic plugin MCP discovery. The installation guide describes explicit operator MCP registration as a separate route. Direct COM attachment does not work inside the tested Colossus windows_job/AppContainer boundary. Hosted CI does not exercise a real mailbox. Full mailbox/Office compatibility remains unvalidated. Do not disable isolation as an installation workaround.

See [installation and verification](https://github.com/obscuritylabs/colossus-plugins/blob/$GITHUB_SHA/docs/releases.md). Keep every file in the portable ZIP together. The OCI archive includes the plugin and its signing evidence for disconnected verification.
EOF
cat "$release_dir/release-notes.md" >> "$GITHUB_STEP_SUMMARY"
