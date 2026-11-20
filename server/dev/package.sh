#!/usr/bin/env bash
# Packages the release (server/public_html, server/grid_app, dist from npm run build:cloud) as a gzip'd ustar
# archive with a sha256 manifest, and builds the one-file DirectAdmin deployer.
#
#   GRID_DEPLOY_KEY=<key> bash server/dev/package.sh            → small deployer that downloads the release
#                                                                    from GitHub (MODE=remote, the default)
#   MODE=payload bash server/dev/package.sh                      → deploy/grid-cloud.tar.gz + .sha256 (CI)
#   MODE=embed GRID_DEPLOY_KEY=<key> bash server/dev/package.sh → one big deployer with everything inside
#
# The deploy key is never stored in the repository.
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT="${OUT:-$ROOT/deploy/deployment.php}"
MODE="${MODE:-remote}"
REMOTE_URL="${GRID_REMOTE_URL:-https://github.com/clerdoapp/clerdogrid/releases/download/cloud-latest/grid-cloud.tar.gz}"
TEMPLATE="$ROOT/dev/deployer-template.php"
DIST="${GRID_DIST:-$(dirname "$ROOT")/dist}"   # GRID_DIST overrides for testing only

die() { echo "package.sh: $*" >&2; exit 1; }

[ -f "$TEMPLATE" ] || die "missing $TEMPLATE"
case "$MODE" in remote|payload|embed) ;; *) die "MODE must be remote, payload or embed" ;; esac
if [ "$MODE" != remote ]; then
[ -f "$DIST/index.html" ] || die "dist/index.html not found — run 'npm run build:cloud' first"
grep -q "Grid" "$DIST/index.html" 2>/dev/null || echo "warning: dist/index.html does not mention Grid; is this the right build?" >&2
if grep -q "dev placeholder" "$DIST/index.html"; then
  die "dist/index.html is the dev placeholder written by dev/start.sh — build the real SPA first"
fi
grep -q "clerdo-cloud-build" "$DIST/index.html" || die "dist/ is not the cloud build — run npm run build:cloud"
fi
if [ "$MODE" != payload ]; then
[ -n "${GRID_DEPLOY_KEY:-}" ] || die "set GRID_DEPLOY_KEY (the ?key= value for deployment.php)"
[[ "$GRID_DEPLOY_KEY" =~ ^[A-Za-z0-9_-]{6,128}$ ]] || die "GRID_DEPLOY_KEY must be 6-128 letters, digits, - or _"
fi
for f in "$ROOT"/grid_app/*.php "$ROOT"/grid_app/lib/*.php "$ROOT"/public_html/*.php "$TEMPLATE"; do
  php -l "$f" >/dev/null || die "php -l failed: $f"
done

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
SHA=""; FILES=0; RAW=0; GZ=0
if [ "$MODE" != remote ]; then
mkdir -p "$STAGE/pkg/public_html" "$STAGE/pkg/grid_app"

# public_html: SPA build + entry point + .htaccess (the deployer itself is never shipped).
cp -R "$DIST"/. "$STAGE/pkg/public_html/"
cp "$ROOT/public_html/.htaccess" "$ROOT/public_html/api.php" "$STAGE/pkg/public_html/"
rm -f "$STAGE/pkg/public_html/deployment.php"
# Which commit this is: https://<app>.clerdo.app/build.txt
printf 'commit %s\nbuilt %s\n' "${GITHUB_SHA:-$(git -C "$ROOT/.." rev-parse HEAD 2>/dev/null || echo unknown)}" "$(date -u +%Y-%m-%dT%H:%M:%SZ)" > "$STAGE/pkg/public_html/build.txt"
find "$STAGE/pkg/public_html" -name '*.map' -delete

# grid_app: everything except the live config and local junk.
(cd "$ROOT/grid_app" && find . -type f \
    ! -name 'config.php' ! -name '.DS_Store' ! -path '*/.git/*' ! -name '*.sqlite*' \
    -exec cp --parents {} "$STAGE/pkg/grid_app/" \;)
[ -f "$STAGE/pkg/grid_app/lib/Core.php" ] || die "grid_app/lib/Core.php missing from stage"
printf 'Require all denied\n' > "$STAGE/pkg/grid_app/.htaccess"

# Manifest: sha256 of every file, paths relative to the package root.
(cd "$STAGE/pkg" && find . -type f ! -name 'manifest.sha256' | sed 's#^\./##' | LC_ALL=C sort | xargs -d '\n' sha256sum > manifest.sha256)
FILES=$(wc -l < "$STAGE/pkg/manifest.sha256")

# ustar/gnu tar, deterministic order, no owner info; the PHP side has a ~60-line reader.
tar --format=gnu --owner=0 --group=0 --numeric-owner --mtime='2000-01-01 00:00:00Z' \
    -C "$STAGE/pkg" -cf "$STAGE/payload.tar" manifest.sha256 public_html grid_app
gzip -9 -n < "$STAGE/payload.tar" > "$STAGE/payload.tar.gz"
# Verify the archive round-trips before embedding.
tar -tzf "$STAGE/payload.tar.gz" > "$STAGE/list.txt" || die "payload tar is unreadable"
grep -qx 'manifest.sha256' "$STAGE/list.txt" || die "payload tar lacks manifest.sha256"

SHA=$(sha256sum "$STAGE/payload.tar.gz" | cut -d' ' -f1)
RAW=$(stat -c %s "$STAGE/payload.tar")
GZ=$(stat -c %s "$STAGE/payload.tar.gz")
fi
BUILD="$(date -u +%Y%m%d-%H%M%S)-$(printf '%s' "${SHA:-remote00}" | cut -c1-8)"

if [ "$MODE" = payload ]; then
  mkdir -p "$ROOT/deploy"
  cp "$STAGE/payload.tar.gz" "$ROOT/deploy/grid-cloud.tar.gz"
  echo "$SHA  grid-cloud.tar.gz" > "$ROOT/deploy/grid-cloud.tar.gz.sha256"
  echo "release archive: $ROOT/deploy/grid-cloud.tar.gz ($GZ bytes, $FILES files, sha256 $SHA)"
  exit 0
fi

mkdir -p "$(dirname "$OUT")"
# Substitute placeholders with PHP (the payload is far too big for sed).
php -d memory_limit=1G -r '
    [$tpl, $gz, $out, $sha, $build, $key, $remote] = array_slice($argv, 1);
    $src = file_get_contents($tpl);
    foreach (["__PAYLOAD__", "__PAYLOAD_SHA256__", "__BUILD_ID__", "__DEPLOY_KEY__", "__REMOTE_URL__"] as $ph) {
        if (substr_count($src, $ph) !== 1) { fwrite(STDERR, "template placeholder $ph not found exactly once\n"); exit(1); }
    }
    $b64 = $remote === "" ? base64_encode(file_get_contents($gz)) : "";
    $src = str_replace(["__PAYLOAD_SHA256__", "__BUILD_ID__", "__DEPLOY_KEY__", "__REMOTE_URL__"], [$sha, $build, $key, $remote], $src);
    $src = str_replace("__PAYLOAD__", $b64, $src);
    $tmp = $out . ".tmp";
    if (file_put_contents($tmp, $src) !== strlen($src)) { fwrite(STDERR, "write failed\n"); exit(1); }
    rename($tmp, $out);
' "$TEMPLATE" "$STAGE/payload.tar.gz" "$OUT" "$SHA" "$BUILD" "$GRID_DEPLOY_KEY" "$([ "$MODE" = remote ] && echo "$REMOTE_URL")"

php -l "$OUT" >/dev/null || die "built deployer fails php -l"

# Self-check: pull the embedded payload back out and confirm the checksum and a key file.
[ "$MODE" = embed ] && php -d memory_limit=1G -r '
    $src = file_get_contents($argv[1]);
    preg_match("/const PAYLOAD = \x27([^\x27]*)\x27;/", $src, $m) || exit(2);
    preg_match("/const PAYLOAD_SHA256 = \x27([a-f0-9]{64})\x27;/", $src, $h) || exit(3);
    $gz = base64_decode($m[1], true);
    if ($gz === false || hash("sha256", $gz) !== $h[1]) exit(4);
    $tar = gzdecode($gz);
    if ($tar === false || !str_contains($tar, "grid_app/lib/Core.php") || !str_contains($tar, "public_html/index.html")) exit(5);
' "$OUT" || { [ "$MODE" = remote ] || die "self-check of embedded payload failed"; }

OUT_SZ=$(stat -c %s "$OUT")
echo "deployment.php built"
echo "  build:    $BUILD"
if [ "$MODE" = remote ]; then
echo "  release:  downloaded by the server from $REMOTE_URL"
else
echo "  files:    $FILES (public_html + grid_app, config.php excluded)"
echo "  payload:  $RAW bytes tar → $GZ bytes gzip (sha256 $SHA)"
fi
echo "  output:   $OUT ($OUT_SZ bytes)"
echo
echo "Upload to domains/grid.clerdo.app/public_html/deployment.php, then open:"
echo "  https://grid.clerdo.app/deployment.php?key=<GRID_DEPLOY_KEY>"
echo "  &go=1 → &config=1 → &client=1 → &migrate=1 → &probe=1 → &delete=1"
