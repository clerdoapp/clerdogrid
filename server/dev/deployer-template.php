<?php
declare(strict_types=1);

/**
 * Grid one-file deployer for DirectAdmin — upload to domains/grid.clerdo.app/public_html/deployment.php
 * and open https://grid.clerdo.app/deployment.php?key=...
 *
 * Actions (all GET, all idempotent): go=1 (install files), config=1 (write config.php), client=1
 * (register OAuth clients at auth.clerdo.app), migrate=1, probe=1, delete=1.
 * The payload below is a gzip'd ustar archive, base64 encoded, built by dev/package.sh.
 */

const DEPLOY_KEY = '__DEPLOY_KEY__';
const PAYLOAD = '__PAYLOAD__';
const PAYLOAD_SHA256 = '__PAYLOAD_SHA256__';
const BUILD_ID = '__BUILD_ID__';

@ini_set('display_errors', '1');
@ini_set('memory_limit', '512M');
error_reporting(E_ALL);
set_time_limit(300);
header('Cache-Control: no-store');
header('X-Content-Type-Options: nosniff');
header('Content-Type: text/html; charset=utf-8');
header("Content-Security-Policy: default-src 'none'; style-src 'unsafe-inline'; form-action 'self'");

$key = (string)($_GET['key'] ?? '');
if ($key === '' || !hash_equals(DEPLOY_KEY, $key)) {
    http_response_code(403);
    echo '<!doctype html><title>403</title><p>Forbidden.</p>';
    exit;
}

$PUBLIC = __DIR__;
$APP = dirname(__DIR__) . '/grid_app';
$DOMAINS = dirname(__DIR__, 2); // .../domains
$SELF = __FILE__;

function h(string $s): string
{
    return htmlspecialchars($s, ENT_QUOTES, 'UTF-8');
}
function link_(string $action, string $label): string
{
    return '<a href="?key=' . h(DEPLOY_KEY) . '&' . $action . '">' . h($label) . '</a>';
}
function page(string $title, string $body): never
{
    echo '<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>' . h($title) . '</title>';
    echo '<style>body{font:15px/1.5 -apple-system,Segoe UI,Inter,sans-serif;max-width:900px;margin:40px auto;padding:0 16px;color:#111}code,pre{font:13px/1.4 ui-monospace,Menlo,monospace;background:#f3f4f8;padding:2px 5px;border-radius:4px}pre{padding:12px;overflow:auto}table{border-collapse:collapse}td,th{padding:4px 10px;border-bottom:1px solid #e3e5ec;text-align:left}.ok{color:#137333}.bad{color:#b3261e}input{font:inherit;padding:6px 8px;border:1px solid #c8ccd8;border-radius:6px;width:360px}label{display:block;margin:10px 0 2px}button{font:inherit;padding:8px 16px;border-radius:8px;border:0;background:#2347FF;color:#fff}</style>';
    echo '<h1>' . h($title) . '</h1>' . $body;
    echo '<p style="margin-top:40px"><a href="?key=' . h(DEPLOY_KEY) . '">← menu</a></p>';
    exit;
}

/** Atomic write: tmp + rename, with directory creation. */
function writeAtomic(string $path, string $data, int $mode = 0644): void
{
    $dir = dirname($path);
    if (!is_dir($dir) && !mkdir($dir, 0755, true) && !is_dir($dir)) {
        throw new RuntimeException("cannot create $dir");
    }
    $tmp = $dir . '/.' . basename($path) . '.' . bin2hex(random_bytes(4)) . '.tmp';
    if (file_put_contents($tmp, $data) !== strlen($data)) {
        @unlink($tmp);
        throw new RuntimeException("cannot write $path");
    }
    chmod($tmp, $mode);
    if (!rename($tmp, $path)) {
        @unlink($tmp);
        throw new RuntimeException("cannot rename into $path");
    }
}

/** Minimal ustar reader: returns [path => [content, mode]] for regular files. */
function tarEntries(string $tar): array
{
    $out = [];
    $len = strlen($tar);
    $off = 0;
    $longName = null;
    while ($off + 512 <= $len) {
        $hdr = substr($tar, $off, 512);
        $off += 512;
        if (trim($hdr, "\0") === '') {
            break;
        }
        $name = rtrim(substr($hdr, 0, 100), "\0");
        $mode = octdec(trim(substr($hdr, 100, 8), "\0 ")) ?: 0644;
        $size = (int)octdec(trim(substr($hdr, 124, 12), "\0 "));
        $type = substr($hdr, 156, 1);
        $prefix = rtrim(substr($hdr, 345, 155), "\0");
        if ($prefix !== '' && substr($hdr, 257, 5) === 'ustar') {
            $name = $prefix . '/' . $name;
        }
        $data = substr($tar, $off, $size);
        $off += (int)(ceil($size / 512) * 512);
        if ($type === 'L') { // GNU long name
            $longName = rtrim($data, "\0");
            continue;
        }
        if ($longName !== null) {
            $name = $longName;
            $longName = null;
        }
        if ($type === '0' || $type === "\0" || $type === '') {
            $name = ltrim($name, './');
            if ($name !== '' && !str_contains($name, '..')) {
                $out[$name] = [$data, (int)$mode];
            }
        }
    }
    return $out;
}

function payloadFiles(): array
{
    if (PAYLOAD === '' || str_starts_with(PAYLOAD, '__')) {
        throw new RuntimeException('This deployer has no payload (template). Run dev/package.sh.');
    }
    $gz = base64_decode(PAYLOAD, true);
    if ($gz === false) {
        throw new RuntimeException('payload base64 invalid');
    }
    if (!hash_equals(PAYLOAD_SHA256, hash('sha256', $gz))) {
        throw new RuntimeException('payload checksum mismatch — upload was corrupted');
    }
    $tar = gzdecode($gz);
    if ($tar === false) {
        throw new RuntimeException('payload gunzip failed');
    }
    return tarEntries($tar);
}

function envInfo(string $PUBLIC, string $APP): string
{
    $srv = $_SERVER['SERVER_SOFTWARE'] ?? 'unknown';
    $rows = [
        'PHP' => PHP_VERSION,
        'Server' => $srv . (stripos($srv, 'litespeed') !== false ? ' (LiteSpeed)' : ''),
        'Build' => BUILD_ID,
    ];
    foreach (['pdo_mysql', 'zip', 'openssl', 'sodium', 'mbstring', 'curl', 'json', 'zlib'] as $ext) {
        $rows['ext ' . $ext] = extension_loaded($ext) ? '<span class=ok>yes</span>' : '<span class=bad>missing</span>';
    }
    $rows['public_html writable'] = is_writable($PUBLIC) ? '<span class=ok>yes</span>' : '<span class=bad>no</span>';
    $rows['parent writable (for ../grid_app)'] = is_writable(dirname($PUBLIC)) ? '<span class=ok>yes</span>' : '<span class=bad>no</span>';
    $rows['../grid_app exists'] = is_dir($APP) ? '<span class=ok>yes</span>' : 'no';
    $rows['config.php exists'] = is_file($APP . '/config.php') ? '<span class=ok>yes</span>' : '<span class=bad>no</span>';
    $rows['index.html exists'] = is_file($PUBLIC . '/index.html') ? '<span class=ok>yes</span>' : 'no';
    $t = '<table>';
    foreach ($rows as $k => $v) {
        $t .= '<tr><th>' . h($k) . '</th><td>' . (str_starts_with($v, '<span') ? $v : h($v)) . '</td></tr>';
    }
    return $t . '</table>';
}

function loadConfig(string $APP): ?array
{
    $f = $APP . '/config.php';
    if (!is_file($f)) {
        return null;
    }
    if (function_exists('opcache_invalidate')) {
        @opcache_invalidate($f, true); // never read a stale compiled copy after a rewrite
    }
    $c = include $f;
    return is_array($c) ? $c : null;
}

/** Boot the real app library against the live config (for ensureSchema / probe). */
function bootApp(string $APP): void
{
    if (!is_file($APP . '/lib/Core.php')) {
        throw new RuntimeException('grid_app not installed yet — run go=1 first');
    }
    spl_autoload_register(static function (string $class) use ($APP): void {
        if (!str_starts_with($class, 'Grid\\')) {
            return;
        }
        $short = substr($class, 5);
        $core = ['HttpError', 'Db', 'Config', 'Json', 'Crypto', 'Http', 'Id', 'Clock', 'RateLimit'];
        $file = $APP . '/lib/' . (in_array($short, $core, true) ? 'Core' : $short) . '.php';
        if (is_file($file)) {
            require_once $file;
        }
    });
    require_once $APP . '/lib/Core.php';
    $cfg = loadConfig($APP);
    if ($cfg === null) {
        throw new RuntimeException('config.php missing — run config=1 first');
    }
    if (!defined('Grid\\GRID_VERSION')) {
        define('Grid\\GRID_VERSION', 'deployer');
    }
    Grid\Config::load($cfg);
}

/** Look for a sibling Clerdo module's config to prefill DB credentials. */
function siblingDb(string $DOMAINS): array
{
    foreach ([
        $DOMAINS . '/frame.clerdo.app/grid_app/config.php',
        $DOMAINS . '/zoetra.clerdo.app/zoetra_app/config.php',
        $DOMAINS . '/frame.clerdo.app/frame_app/config.php',
        $DOMAINS . '/graphix.clerdo.app/graphix_app/config.php',
    ] as $f) {
        if (is_readable($f)) {
            $c = @include $f;
            if (is_array($c) && isset($c['db']) && is_array($c['db'])) {
                return ['source' => $f, 'host' => (string)($c['db']['host'] ?? '127.0.0.1'), 'name' => (string)($c['db']['name'] ?? ''), 'user' => (string)($c['db']['user'] ?? ''), 'pass' => (string)($c['db']['pass'] ?? '')];
            }
        }
    }
    return ['source' => null, 'host' => 'localhost', 'name' => '', 'user' => '', 'pass' => ''];
}

try {
    // ------------------------------------------------------------------ go=1
    if (isset($_GET['go'])) {
        $files = payloadFiles();
        if (!isset($files['manifest.sha256'])) {
            throw new RuntimeException('payload has no manifest');
        }
        $manifest = [];
        foreach (explode("\n", trim($files['manifest.sha256'][0])) as $line) {
            if (preg_match('/^([a-f0-9]{64})\s+\*?(.+)$/', trim($line), $m)) {
                $manifest[$m[2]] = $m[1];
            }
        }
        $written = [];
        $skipped = [];
        $total = 0;
        $bad = [];
        $newAssets = [];
        foreach ($files as $rel => [$data, $mode]) {
            if ($rel === 'manifest.sha256') {
                continue;
            }
            if (!isset($manifest[$rel])) {
                $bad[] = $rel . ' (not in manifest)';
                continue;
            }
            if (!hash_equals($manifest[$rel], hash('sha256', $data))) {
                $bad[] = $rel . ' (checksum mismatch)';
                continue;
            }
            if (str_starts_with($rel, 'public_html/')) {
                $dest = $PUBLIC . '/' . substr($rel, 12);
            } elseif (str_starts_with($rel, 'grid_app/')) {
                $dest = $APP . '/' . substr($rel, 9);
            } else {
                $bad[] = $rel . ' (unexpected root)';
                continue;
            }
            if (basename($dest) === 'config.php' && is_file($dest)) {
                $skipped[] = $rel . ' (kept live config)';
                continue;
            }
            if (realpath($dest) === realpath($SELF)) {
                $skipped[] = $rel . ' (deployer itself)';
                continue;
            }
            if (str_starts_with($rel, 'public_html/assets/')) {
                $newAssets[basename($dest)] = true;
            }
            writeAtomic($dest, $data, str_ends_with($dest, '.php') && str_contains($dest, '/grid_app/') ? 0640 : 0644);
            $written[] = [$rel, strlen($data), hash('sha256', $data) === $manifest[$rel]];
            $total += strlen($data);
        }
        if ($bad !== []) {
            throw new RuntimeException('Refusing to install, payload integrity failed: ' . implode(', ', $bad));
        }
        // Remove stale assets from previous builds.
        $removed = [];
        if (is_dir($PUBLIC . '/assets') && $newAssets !== []) {
            foreach (scandir($PUBLIC . '/assets') ?: [] as $f) {
                if ($f[0] !== '.' && !isset($newAssets[$f]) && is_file($PUBLIC . '/assets/' . $f)) {
                    @unlink($PUBLIC . '/assets/' . $f);
                    $removed[] = $f;
                }
            }
        }
        // Guard grid_app against direct web access should it ever end up under a docroot.
        if (!is_file($APP . '/.htaccess')) {
            writeAtomic($APP . '/.htaccess', "Require all denied\n");
        }
        if (function_exists('opcache_reset')) {
            @opcache_reset();
        }
        $b = '<p class=ok>Installed ' . count($written) . ' files, ' . number_format($total) . ' bytes. All sha256 verified.</p><table><tr><th>file</th><th>bytes</th><th>sha256</th></tr>';
        foreach ($written as [$rel, $n, $okk]) {
            $b .= '<tr><td><code>' . h($rel) . '</code></td><td>' . number_format($n) . '</td><td class=' . ($okk ? 'ok>ok' : 'bad>BAD') . '</td></tr>';
        }
        $b .= '</table>';
        if ($skipped) {
            $b .= '<p>Skipped: <code>' . h(implode('</code>, <code>', $skipped)) . '</code></p>';
        }
        if ($removed) {
            $b .= '<p>Removed stale assets: <code>' . h(implode('</code>, <code>', $removed)) . '</code></p>';
        }
        $b .= is_file($APP . '/config.php')
            ? '<p>Next: ' . link_('migrate=1', 'run migrations') . ' then ' . link_('probe=1', 'probe') . '.</p>'
            : '<p>Next: ' . link_('config=1', 'create config.php') . '.</p>';
        page('Grid — files installed', $b);
    }

    // -------------------------------------------------------------- config=1
    if (isset($_GET['config'])) {
        $host = $_SERVER['HTTP_HOST'] ?? 'grid.clerdo.app';
        $https = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') || ($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? '') === 'https';
        $base = ($https ? 'https' : 'http') . '://' . $host;
        $existing = loadConfig($APP);
        // The form POSTs so the DB password never appears in a URL or access log.
        if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['save'])) {
            $db = ['driver' => 'mysql', 'host' => trim((string)($_POST['db_host'] ?? '')), 'port' => 3306, 'name' => trim((string)($_POST['db_name'] ?? '')), 'user' => trim((string)($_POST['db_user'] ?? '')), 'pass' => (string)($_POST['db_pass'] ?? '')];
            if ($db['name'] === '' || $db['user'] === '') {
                throw new RuntimeException('DB name and user are required');
            }
            if ($db['pass'] === '') {
                // Blank = keep the live password, or the sibling module's when prefilled from it.
                $fallback = $existing['db'] ?? siblingDb($DOMAINS);
                if (($fallback['name'] ?? '') === $db['name'] && ($fallback['user'] ?? '') === $db['user']) {
                    $db['pass'] = (string)($fallback['pass'] ?? '');
                }
            }
            $encKey = $existing['encryption_key'] ?? base64_encode(random_bytes(32));
            $installKey = $existing['install_key'] ?? rtrim(strtr(base64_encode(random_bytes(18)), '+/', '-_'), '=');
            $issuer = rtrim(trim((string)($_POST['issuer'] ?? 'https://auth.clerdo.app')), '/');
            if (!preg_match('#^https://[a-z0-9.-]+$#i', $issuer)) {
                throw new RuntimeException('issuer must be an https origin');
            }
            $cfg = [
                'db' => $db,
                'encryption_key' => $encKey,
                'install_key' => $installKey,
                'auth' => [
                    'issuer' => $issuer,
                    'client_id' => 'grid-web',
                    'redirect_uri' => $base . '/auth/callback',
                    'authorize_url' => $issuer . '/authorize',
                    'token_url' => $issuer . '/token',
                    'userinfo_url' => $issuer . '/userinfo',
                    'revoke_url' => $issuer . '/revoke',
                    'scope' => 'openid profile email',
                ],
                'pro' => ['everyone_has_pro' => (($_POST['everyone_pro'] ?? '1') === '1')],
                'limits' => ['free_projects' => 5, 'free_doc_bytes' => 16 * 1024 * 1024, 'pro_doc_bytes' => 25 * 1024 * 1024, 'share_per_minute' => 60, 'auth_per_minute' => 30],
                'app' => ['base_url' => $base],
            ];
            $php = "<?php\n// Generated by deployment.php on " . gmdate('c') . ". Keep private (0600).\nreturn " . var_export($cfg, true) . ";\n";
            writeAtomic($APP . '/config.php', $php, 0600);
            // Test DB + schema.
            $b = '<p class=ok>config.php written (0600). Secrets are not shown.</p>';
            try {
                bootApp($APP);
                Grid\Db::pdo();
                $b .= '<p class=ok>DB connection OK.</p>';
                Grid\Projects::ensureSchema();
                $b .= '<p class=ok>Schema ensured (grid_* tables).</p>';
            } catch (Throwable $e) {
                $b .= '<p class=bad>DB check failed: ' . h($e->getMessage()) . '</p>';
            }
            $b .= '<p>Next: ' . link_('client=1', 'register OAuth clients') . ' · ' . link_('probe=1', 'probe') . '</p>';
            page('Grid — config saved', $b);
        }
        $sib = siblingDb($DOMAINS);
        $pre = $existing['db'] ?? $sib;
        $b = '<p>' . ($existing ? '<b>A config.php already exists.</b> Saving will keep its encryption_key and install_key but replace everything else.' : 'No config.php yet.') . '</p>';
        if ($sib['source'] && !$existing) {
            $b .= '<p>Prefilled from <code>' . h($sib['source']) . '</code>.</p>';
        }
        // Password is never echoed back: an existing/sibling one is used when the field is left blank.
        $hasPass = (string)($pre['pass'] ?? '') !== '';
        $b .= '<form method="post" action="?key=' . h(DEPLOY_KEY) . '&config=1"><input type=hidden name=save value=1>';
        $b .= '<label>DB host</label><input name=db_host value="' . h((string)($pre['host'] ?? 'localhost')) . '">';
        $b .= '<label>DB name</label><input name=db_name value="' . h((string)($pre['name'] ?? '')) . '">';
        $b .= '<label>DB user</label><input name=db_user value="' . h((string)($pre['user'] ?? '')) . '">';
        $b .= '<label>DB password</label><input name=db_pass type=password value="" placeholder="' . ($hasPass ? 'leave blank to keep the existing password' : '') . '">';
        $b .= '<label>Auth issuer</label><input name=issuer value="' . h((string)($existing['auth']['issuer'] ?? 'https://auth.clerdo.app')) . '">';
        $b .= '<label>Everyone has Pro</label><input name=everyone_pro value="' . ((($existing['pro']['everyone_has_pro'] ?? true) ? '1' : '0')) . '"> (1 or 0)';
        $b .= '<p>Redirect URI will be <code>' . h($base) . '/auth/callback</code>; client_id <code>grid-web</code>. encryption_key and install_key are generated.</p>';
        $b .= '<p><button type=submit>Write config.php</button></p></form>';
        page('Grid — configuration', $b);
    }

    // -------------------------------------------------------------- client=1
    if (isset($_GET['client'])) {
        $clientsFile = $DOMAINS . '/auth.clerdo.app/public_html/app/Clients.php';
        $host = $_SERVER['HTTP_HOST'] ?? 'grid.clerdo.app';
        $webRedirect = 'https://' . $host . '/auth/callback';
        $wanted = [
            'grid-web' => $webRedirect,
            'grid-ios' => 'grid://auth/callback',
        ];
        $snippet = '';
        foreach ($wanted as $id => $redir) {
            $snippet .= "    '" . $id . "' => [\n        'name' => 'Grid',\n        'public' => true,\n        'pkce' => true,\n        'redirect_uris' => ['" . $redir . "'],\n        'scopes' => ['openid', 'profile', 'email'],\n    ],\n";
        }
        if (!is_file($clientsFile)) {
            page('Grid — OAuth clients', '<p class=bad>Not found: <code>' . h($clientsFile) . '</code>. Add this to the auth server\'s client registry manually:</p><pre>' . h($snippet) . '</pre>');
        }
        if (!is_writable($clientsFile)) {
            page('Grid — OAuth clients', '<p class=bad><code>' . h($clientsFile) . '</code> is not writable. Add manually:</p><pre>' . h($snippet) . '</pre>');
        }
        $src = file_get_contents($clientsFile);
        $report = [];
        $changed = false;
        foreach ($wanted as $id => $redir) {
            if (preg_match("/['\"]" . preg_quote($id, '/') . "['\"]\s*=>/", $src)) {
                $report[] = '<li><code>' . h($id) . '</code> already registered.</li>';
                continue;
            }
            // Find an existing sibling entry: 'frame-web' => [ ... ],  (balanced brackets)
            $template = null;
            $templateId = null;
            $candidates = str_ends_with($id, '-ios')
                ? ['frame-ios', 'zoetra-ios', 'graphix-ios', 'frame-web', 'zoetra-web', 'graphix-web']
                : ['frame-web', 'zoetra-web', 'graphix-web'];
            foreach ($candidates as $sib) {
                if (preg_match("/(^[ \t]*)(['\"])" . preg_quote($sib, '/') . "\\2\s*=>\s*\[/m", $src, $m, PREG_OFFSET_CAPTURE)) {
                    $start = $m[0][1];
                    $indent = $m[1][0];
                    // Walk to the matching close bracket.
                    $pos = $start + strlen($m[0][0]);
                    $depth = 1;
                    $n = strlen($src);
                    $inStr = null;
                    while ($pos < $n && $depth > 0) {
                        $ch = $src[$pos];
                        if ($inStr !== null) {
                            if ($ch === '\\') {
                                $pos++;
                            } elseif ($ch === $inStr) {
                                $inStr = null;
                            }
                        } elseif ($ch === "'" || $ch === '"') {
                            $inStr = $ch;
                        } elseif ($ch === '[') {
                            $depth++;
                        } elseif ($ch === ']') {
                            $depth--;
                        }
                        $pos++;
                    }
                    if ($depth !== 0) {
                        continue;
                    }
                    // include trailing comma
                    $end = $pos;
                    if (isset($src[$end]) && $src[$end] === ',') {
                        $end++;
                    }
                    $template = substr($src, $start, $end - $start);
                    $templateId = $sib;
                    break;
                }
            }
            if ($template === null) {
                $report[] = '<li class=bad>Could not find a <code>frame-web</code>/<code>zoetra-web</code>/<code>graphix-web</code> entry to clone for <code>' . h($id) . '</code>. Add manually (see snippet below).</li>';
                continue;
            }
            $entry = $template;
            $entry = preg_replace("/(['\"])" . preg_quote($templateId, '/') . "\\1(\s*=>)/", '$1' . $id . '$1$2', $entry, 1);
            // Replace every redirect URI in the cloned entry with ours.
            $entry = preg_replace("/(['\"])https?:\/\/[^'\"]*\/auth\/callback[^'\"]*\\1/", "'" . $redir . "'", $entry) ?? $entry;
            $entry = preg_replace("/(['\"])[a-z][a-z0-9+.-]*:\/\/[^'\"]*\\1/i", "'" . $redir . "'", $entry) ?? $entry;
            // Human name: Frame/Zoetra/Graphix → Grid
            $entry = preg_replace("/(['\"]name['\"]\s*=>\s*['\"])[^'\"]*(['\"])/", '${1}Grid' . ($id === 'grid-ios' ? ' iOS' : '') . '$2', $entry) ?? $entry;
            if (!str_ends_with(rtrim($entry), ',')) {
                $entry = rtrim($entry) . ',';
            }
            // Insert right after the template entry.
            $tplPos = strpos($src, $template);
            $src = substr($src, 0, $tplPos + strlen($template)) . "\n" . $entry . substr($src, $tplPos + strlen($template));
            $changed = true;
            $report[] = '<li class=ok>Added <code>' . h($id) . '</code> (cloned from <code>' . h($templateId) . '</code>, redirect <code>' . h($redir) . '</code>).</li>';
        }
        if ($changed) {
            // Syntax check before committing.
            $tmp = tempnam(sys_get_temp_dir(), 'clients');
            file_put_contents($tmp, $src);
            $lintOk = true;
            if (function_exists('exec')) {
                exec('php -l ' . escapeshellarg($tmp) . ' 2>&1', $o, $rc);
                $lintOk = $rc === 0;
            }
            @unlink($tmp);
            if (!$lintOk) {
                page('Grid — OAuth clients', '<p class=bad>Generated Clients.php failed <code>php -l</code>; nothing written. Add manually:</p><pre>' . h($snippet) . '</pre>');
            }
            copy($clientsFile, $clientsFile . '.' . gmdate('Ymd-His') . '.bak');
            writeAtomic($clientsFile, $src, fileperms($clientsFile) & 0777 ?: 0644);
            if (function_exists('opcache_invalidate')) {
                @opcache_invalidate($clientsFile, true);
            }
        }
        page('Grid — OAuth clients', '<ul>' . implode('', $report) . '</ul>' . ($changed ? '<p>Backup written next to Clients.php.</p>' : '') . '<details><summary>Manual snippet</summary><pre>' . h($snippet) . '</pre></details>');
    }

    // ------------------------------------------------------------- migrate=1
    if (isset($_GET['migrate'])) {
        bootApp($APP);
        Grid\Projects::ensureSchema();
        $rows = '';
        foreach (array_keys(Grid\Projects::schema()) as $t) {
            $c = Grid\Db::one('SELECT COUNT(*) AS c FROM ' . $t);
            $rows .= '<tr><td><code>' . h($t) . '</code></td><td>' . (int)($c['c'] ?? 0) . '</td></tr>';
        }
        page('Grid — migrate', '<p class=ok>Schema ensured.</p><table><tr><th>table</th><th>rows</th></tr>' . $rows . '</table>');
    }

    // --------------------------------------------------------------- probe=1
    if (isset($_GET['probe'])) {
        $b = envInfo($PUBLIC, $APP);
        $b .= '<h2>Checks</h2><ul>';
        try {
            bootApp($APP);
            Grid\Db::pdo();
            $b .= '<li class=ok>DB connect OK (' . h(Grid\Db::driver()) . ')</li>';
            foreach (array_keys(Grid\Projects::schema()) as $t) {
                if (Grid\Db::tableExists($t)) {
                    $c = Grid\Db::one('SELECT COUNT(*) AS c FROM ' . $t);
                    $b .= '<li class=ok><code>' . h($t) . '</code>: ' . (int)($c['c'] ?? 0) . ' rows</li>';
                } else {
                    $b .= '<li class=bad><code>' . h($t) . '</code> missing (run migrate)</li>';
                }
            }
            $plain = 'probe-' . bin2hex(random_bytes(4));
            $rt = Grid\Crypto::decrypt(Grid\Crypto::encrypt($plain));
            $b .= $rt === $plain ? '<li class=ok>AES-256-GCM encrypt/decrypt OK</li>' : '<li class=bad>encrypt/decrypt mismatch</li>';
            $issuer = (string)Grid\Config::get('auth.issuer');
            foreach ([$issuer . '/.well-known/openid-configuration', (string)Grid\Config::get('auth.authorize_url')] as $u) {
                $ch = curl_init($u);
                curl_setopt_array($ch, [CURLOPT_NOBODY => true, CURLOPT_RETURNTRANSFER => true, CURLOPT_TIMEOUT => 6, CURLOPT_FOLLOWLOCATION => false]);
                curl_exec($ch);
                $code = curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
                $err = curl_error($ch);
                curl_close($ch);
                $b .= ($code > 0 ? '<li class=ok>' : '<li class=bad>') . 'HEAD <code>' . h($u) . '</code> → ' . ($code ?: h($err)) . '</li>';
            }
        } catch (Throwable $e) {
            $b .= '<li class=bad>' . h($e->getMessage()) . '</li>';
        }
        $tmp = sys_get_temp_dir();
        $b .= '<li class=' . (is_writable($tmp) ? 'ok' : 'bad') . '>temp dir <code>' . h($tmp) . '</code> ' . (is_writable($tmp) ? 'writable' : 'NOT writable') . ' (needed for XLSX export)</li>';
        $b .= '<li class=' . (is_file($PUBLIC . '/.htaccess') ? 'ok' : 'bad') . '>.htaccess ' . (is_file($PUBLIC . '/.htaccess') ? 'present' : 'missing') . '</li>';
        $b .= '</ul>';
        page('Grid — probe', $b);
    }

    // -------------------------------------------------------------- delete=1
    if (isset($_GET['delete'])) {
        if (@unlink($SELF)) {
            echo '<!doctype html><meta charset="utf-8"><title>Removed</title><p class=ok>deployment.php deleted. <a href="/">Open Grid</a></p>';
            exit;
        }
        page('Grid — delete', '<p class=bad>Could not delete ' . h(basename($SELF)) . '; remove it via FTP/File Manager.</p>');
    }

    // ------------------------------------------------------------------ menu
    $b = envInfo($PUBLIC, $APP);
    $b .= '<h2>Actions</h2><ol>';
    $b .= '<li>' . link_('go=1', 'Install / update files') . ' — writes public_html/* and ../grid_app/* from the embedded payload (keeps live config.php, prunes old assets).</li>';
    $b .= '<li>' . link_('config=1', 'Configure') . ' — write ../grid_app/config.php (DB prefilled from Frame/Zoetra when readable) and test the DB.</li>';
    $b .= '<li>' . link_('client=1', 'Register OAuth clients') . ' — adds grid-web and grid-ios to auth.clerdo.app Clients.php.</li>';
    $b .= '<li>' . link_('migrate=1', 'Migrate') . ' — create/upgrade grid_* tables.</li>';
    $b .= '<li>' . link_('probe=1', 'Probe') . ' — diagnostics.</li>';
    $b .= '<li>' . link_('delete=1', 'Delete this deployer') . ' — do this when finished.</li>';
    $b .= '</ol>';
    page('Grid deployer', $b);
} catch (Throwable $e) {
    http_response_code(500);
    page('Grid deployer — error', '<p class=bad>' . h($e->getMessage()) . '</p>');
}
