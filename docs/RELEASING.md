# Release pipeline

Pushing an annotated `vX.Y.Z` tag starts `.github/workflows/release.yml`.
The tag version must match `package.json`, `package-lock.json`, `Cargo.toml`,
and `Cargo.lock`.

Ordinary branch pushes and pull requests do not start automated builds. The
full `.github/workflows/ci.yml` gate remains available through manual dispatch;
automatic compilation is reserved for `v*` release tags.

The workflow:

1. runs the JavaScript, package, Rust, and native release gates;
2. builds the Apache-2.0 SDK, public Viewer package, Pages demo and
   official-site bundle from the same commit, then uploads them to a draft release;
3. dispatches `publish.yml` to verify checksums, publish `@docviewkit/viewer`,
   deploy Pages and publish the GitHub Release;
4. optionally deploys the exact official-site bundle when
   `DOCVIEWKIT_DEPLOY_WEBSITE=true`.

## Source repository configuration

The canonical source and release repository is `docviewkit/viewer`. Its
`GITHUB_TOKEN` creates one draft containing all artifacts and triggers the
existing `publish.yml` with `repository_dispatch` (`release-ready`),
which retains the npm Trusted Publishing identity. No cross-repository token
or GitHub App is needed. Release jobs run only in `docviewkit/viewer`;
keep version tags attached to the actual source commit.

Commit `.github/workflows/ci.yml`, `release.yml` and `publish.yml` together
with the source on the default `main` branch before pushing a new version tag.
`repository_dispatch` uses the default branch's publishing workflow. Import
source through a normal commit or merge that preserves the public repository's
history. Existing binary-only release tags must not be pushed again or moved.

Repository Settings must keep Actions enabled and Pages configured to deploy
from GitHub Actions. Keep the `npm-production` and `github-pages` environments.
Workflow files grant the required write permissions explicitly; the repository's
default token permission can remain read-only. Old `DOCVIEWKIT_APP_ID`,
`DOCVIEWKIT_APP_PRIVATE_KEY` and `DOCVIEWKIT_RELEASE_TOKEN` secrets are no longer
used by this pipeline.

## Optional official website deployment

Create `website-production` in `docviewkit/viewer` and add its Actions secrets:

- `DOCVIEWKIT_WEBSITE_SSH_HOST`
- `DOCVIEWKIT_WEBSITE_SSH_USER`
- `DOCVIEWKIT_WEBSITE_SSH_PRIVATE_KEY`
- `DOCVIEWKIT_WEBSITE_SSH_KNOWN_HOSTS`

GitHub environment secrets are not copied when source moves between repositories.

Set the repository variable `DOCVIEWKIT_DEPLOY_WEBSITE=true` only when
the website-production environment and credentials are configured. The
website job is skipped otherwise; npm, Pages and Releases remain independent.

The production SSH key must be restricted to the website hosting account. The
workflow verifies the release checksum, uploads the exact official-site archive,
switches `~/apps/docviewkit/current`, and preserves archived customer data
under `~/apps/docviewkit/data`. After deployment it verifies:

`https://docviewkit.com/sdk/version.json`

The production canary must also verify that `/sdk/office-viewer-xps.wasm`,
`/sdk/office-viewer-ofd.wasm`, and their versioned equivalents return HTTP 200.

The endpoint must return:

```json
{"name":"@docviewkit/viewer","version":"<released version>"}
```

The official demo imports `/sdk/v<version>/viewer.js`, so a successful release
uses the exact deployed version rather than a mutable build-time dependency.

Protect the `website-production` environment and require approval if desired.

The Namecheap cPanel application uses Node.js 24 in production mode with:

- application root: `apps/docviewkit`
- application URL: `docviewkit.com/`
- startup file: `server.js`

The website needs no account database or authentication configuration. Preserve any existing customer data as a private archive; the website does not open or migrate it.

SSH runs on shared-hosting port `21098`. The deployment script retains every
versioned release for rollback and restarts Passenger through `tmp/restart.txt`.

## npm and Pages configuration

The public workflow is
`docviewkit/viewer/.github/workflows/publish.yml`.

The existing `@docviewkit/viewer` package uses npm Trusted Publishing. Preserve
and verify its configured identity:

- Organization or user: `docviewkit`
- Repository: `viewer`
- Workflow: `publish.yml`
- Environment: `npm-production`
- Allowed action: direct `npm publish`

Keep `publish.yml` and `npm-production` named exactly as configured. OIDC needs
`id-token: write`, a GitHub-hosted runner, npm >=11.5.1 and Node >=22.14.0;
the workflow uses Node 24. It continues to publish the generated
`@docviewkit/viewer` tarball rather than the private root workspace package.
Source migration requires no new npm token. Only initial publication of a new
package needs the optional `NPM_TOKEN` fallback before setting up OIDC.
See [npm Trusted Publishing](https://docs.npmjs.com/trusted-publishers/) and
[GitHub workflow triggers](https://docs.github.com/en/actions/how-tos/write-workflows/choose-when-workflows-run/trigger-a-workflow).

## Open source license

Viewer, Engine and project source use Apache-2.0. Every package, Pages runtime
and official-site bundle includes LICENSE, NOTICE and the third-party licenses.
The public package preserves its Viewer entry and adds `/engine`, `/viewer`
and `/accuracy` exports. Runtime licenses and mandatory branding are removed.
Support, customization and enterprise delivery are arranged by email;
the website has no customer accounts or database and signing keys are never bundled.
Root `private: true` prevents accidental npm publication of the build workspace;
it does not restrict rights under Apache-2.0.

Before importing source into the existing public repository, review Git history
for credentials and confidential content, and review external test documents
for redistribution rights. Preserve existing public Releases and Issues; do not
force-push over public history or move old version tags.

## Retry

If a downstream service fails after artifacts have been built, run the source
workflow manually with the existing source tag. Draft assets can be replaced
until publication completes, and an already published npm version is not
republished. Once the GitHub Release is public, a retry leaves its notes and
attachments intact. The website deployment script is checked out from the same
release tag as its archive.
