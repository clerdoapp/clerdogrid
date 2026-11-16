#!/usr/bin/env bash
# Builds deploy/deployment.php: the one-file DirectAdmin deployer with the whole
# release (server/public_html, server/grid_app, dist from npm run build:cloud) embedded as a gzip'd
# ustar payload + sha256 manifest, base64 inlined at __PAYLOAD__.
#
#   GRID_DEPLOY_KEY=<key> bash server/dev/package.sh   → server/deploy/deployment.php
#   OUT=/path/x.php bash dev/package.sh
#
# dist/ must exist (npm run build:cloud first). The deploy key is never stored in the repository.
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT="${OUT:-$ROOT/deploy/deployment.php}"
TEMPLATE="$ROOT/dev/deployer-template.php"
DIST="${GRID_DIST:-$(dirname "$ROOT")/dist}"   # GRID_DIST overrides for testing only

die() { echo "package.sh: $*" >&2; exit 1; }

[ -f "$TEMPLATE" ] || die "missing $TEMPLATE"
[ -f "$DIST/index.html" ] || die "dist/index.html not found — run 'npm run build:cloud' first"
grep -q "Grid" "$DIST/index.html" 2>/dev/null || echo "warning: dist/index.html does not mention Grid; is this the right build?" >&2
if grep -q "dev placeholder" "$DIST/index.html"; then
  die "dist/index.html is the dev placeholder written by dev/start.sh — build the real SPA first"
fi
[ -n "${GRID_DEPLOY_KEY:-}" ] || die "set GRID_DEPLOY_KEY (the ?key= value for deployment.php)"
[[ "$GRID_DEPLOY_KEY" =~ ^[A-Za-z0-9_-]{6,128}$ ]] || die "GRID_DEPLOY_KEY must be 6-128 letters, digits, - or _"
grep -q "clerdo-cloud-build" "$DIST/index.html" || die "dist/ is not the cloud build — run npm run build:cloud"
for f in "$ROOT"/grid_app/*.php "$ROOT"/grid_app/lib/*.php "$ROOT"/public_html/*.php "$TEMPLATE"; do
  php -l "$f" >/dev/null || die "php -l failed: $f"
done

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
mkdir -p "$STAGE/pkg/public_html" "$STAGE/pkg/grid_app"

# public_html: SPA build + entry point + .htaccess (the deployer itself is never shipped).
cp -R "$DIST"/. "$STAGE/pkg/public_html/"
cp "$ROOT/public_html/.htaccess" "$ROOT/public_html/api.php" "$STAGE/pkg/public_html/"
rm -f "$STAGE/pkg/public_html/deployment.php"
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
BUILD="$(date -u +%Y%m%d-%H%M%S)-$(printf '%s' "$SHA" | cut -c1-8)"
RAW=$(stat -c %s "$STAGE/payload.tar")
GZ=$(stat -c %s "$STAGE/payload.tar.gz")

mkdir -p "$(dirname "$OUT")"
# Substitute placeholders with PHP (the payload is far too big for sed).
php -d memory_limit=1G -r '
    [$tpl, $gz, $out, $sha, $build, $key] = array_slice($argv, 1);
    $src = file_get_contents($tpl);
    foreach (["__PAYLOAD__", "__PAYLOAD_SHA256__", "__BUILD_ID__", "__DEPLOY_KEY__"] as $ph) {
        if (substr_count($src, $ph) !== 1) { fwrite(STDERR, "template placeholder $ph not found exactly once\n"); exit(1); }
    }
    $b64 = base64_encode(file_get_contents($gz));
    $src = str_replace(["__PAYLOAD_SHA256__", "__BUILD_ID__", "__DEPLOY_KEY__"], [$sha, $build, $key], $src);
    $src = str_replace("__PAYLOAD__", $b64, $src);
    $tmp = $out . ".tmp";
    if (file_put_contents($tmp, $src) !== strlen($src)) { fwrite(STDERR, "write failed\n"); exit(1); }
    rename($tmp, $out);
' "$TEMPLATE" "$STAGE/payload.tar.gz" "$OUT" "$SHA" "$BUILD" "$GRID_DEPLOY_KEY"

php -l "$OUT" >/dev/null || die "built deployer fails php -l"

# Self-check: pull the embedded payload back out and confirm the checksum and a key file.
php -d memory_limit=1G -r '
    $src = file_get_contents($argv[1]);
    preg_match("/const PAYLOAD = \x27([^\x27]*)\x27;/", $src, $m) || exit(2);
    preg_match("/const PAYLOAD_SHA256 = \x27([a-f0-9]{64})\x27;/", $src, $h) || exit(3);
    $gz = base64_decode($m[1], true);
    if ($gz === false || hash("sha256", $gz) !== $h[1]) exit(4);
    $tar = gzdecode($gz);
    if ($tar === false || !str_contains($tar, "grid_app/lib/Core.php") || !str_contains($tar, "public_html/index.html")) exit(5);
' "$OUT" || die "self-check of embedded payload failed"

OUT_SZ=$(stat -c %s "$OUT")
echo "deployment.php built"
echo "  build:    $BUILD"
echo "  files:    $FILES (public_html + grid_app, config.php excluded)"
echo "  payload:  $RAW bytes tar → $GZ bytes gzip (sha256 $SHA)"
echo "  output:   $OUT ($OUT_SZ bytes)"
echo
echo "Upload to domains/grid.clerdo.app/public_html/deployment.php, then open:"
echo "  https://grid.clerdo.app/deployment.php?key=<GRID_DEPLOY_KEY>"
echo "  &go=1 → &config=1 → &client=1 → &migrate=1 → &probe=1 → &delete=1"
