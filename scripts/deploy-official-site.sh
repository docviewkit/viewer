#!/usr/bin/env bash
set -euo pipefail

tag=${1:?release tag is required}
version=${2:?release version is required}
archive_name=${3:?archive name is required}

[[ "$tag" == "v$version" ]]
[[ "$version" =~ ^[0-9]+\.[0-9]+\.[0-9]+([+-][0-9A-Za-z.-]+)?$ ]]
[[ "$archive_name" == "docviewkit-official-site-$version.tar.gz" ]]

deploy_root=${DOCVIEWKIT_DEPLOY_ROOT:-"$HOME/apps/docviewkit"}
case "$deploy_root" in
  "$HOME"/*) ;;
  *) echo "Deploy root must be inside the account home directory" >&2; exit 1 ;;
esac

archive="$deploy_root/uploads/$archive_name"
release="$deploy_root/releases/$tag"
staging="$deploy_root/releases/.$tag-staging-$$"

test -f "$archive"
mkdir -p "$deploy_root/releases" "$deploy_root/tmp"

cleanup() { rm -rf "$staging"; }
trap cleanup EXIT

if [[ ! -d "$release" ]]; then
  mkdir "$staging"
  tar -xzf "$archive" -C "$staging"
  test -f "$staging/commercial/src/server.mjs"
  test -f "$staging/commercial/package.json"
  test -f "$staging/dist/version.json"
  actual_version=$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "$staging/dist/version.json" | head -1)
  [[ "$actual_version" == "$version" ]]
  mv "$staging" "$release"
fi

actual_version=$(sed -n 's/.*"version"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' "$release/dist/version.json" | head -1)
[[ "$actual_version" == "$version" ]]
if [[ -e "$deploy_root/current" && ! -L "$deploy_root/current" ]]; then
  echo "Current release path must be a symlink" >&2
  exit 1
fi

printf '%s\n' 'import("./current/commercial/src/server.mjs");' > "$deploy_root/server.js.next"
mv -f "$deploy_root/server.js.next" "$deploy_root/server.js"
ln -sfn "releases/$tag" "$deploy_root/current"
touch "$deploy_root/tmp/restart.txt"
rm -f "$archive"

printf 'Deployed DocViewKit %s\n' "$version"
