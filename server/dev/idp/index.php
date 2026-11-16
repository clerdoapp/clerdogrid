<?php
declare(strict_types=1);

/**
 * Mock auth.clerdo.app for local development. NOT for production.
 *   php -S 127.0.0.1:8090 dev/idp/index.php
 * Endpoints: /authorize, /token (JSON body, PKCE S256), /userinfo, /revoke, /.well-known/openid-configuration
 * State: /tmp/grid-idp.json (override with GRID_IDP_STATE).
 */

const STATE_FILE_DEFAULT = '/tmp/grid-idp.json';
const ACCESS_TTL = 600;

$users = [
    'u-hamzeh' => ['sub' => 'u-hamzeh', 'name' => 'Hamzeh', 'email' => 'hamzeh@example.com', 'picture' => 'https://example.com/a.png', 'plan' => 'pro'],
    'u-eyas' => ['sub' => 'u-eyas', 'name' => 'Eyas', 'email' => 'eyas@example.com', 'picture' => null, 'plan' => 'free'],
];
define('APP_PORT', getenv('GRID_APP_PORT') ?: '8080');

$clients = [
    'grid-web' => ['redirects' => array_values(array_unique([
        'http://127.0.0.1:' . APP_PORT . '/auth/callback',
        'http://localhost:' . APP_PORT . '/auth/callback',
        'http://127.0.0.1:8080/auth/callback',
        'http://localhost:8080/auth/callback',
    ]))],
    'grid-ios' => ['redirects' => ['grid://auth/callback']],
];

$stateFile = getenv('GRID_IDP_STATE') ?: STATE_FILE_DEFAULT;

function load(string $f): array
{
    $raw = @file_get_contents($f);
    $d = $raw ? json_decode($raw, true) : null;
    return is_array($d) ? $d + ['codes' => [], 'access' => [], 'refresh' => []] : ['codes' => [], 'access' => [], 'refresh' => []];
}
function save(string $f, array $d): void
{
    file_put_contents($f, json_encode($d), LOCK_EX);
}
function out(array $body, int $status = 200): never
{
    http_response_code($status);
    header('Content-Type: application/json');
    header('Cache-Control: no-store');
    echo json_encode($body, JSON_UNESCAPED_SLASHES);
    exit;
}
function tok(): string
{
    return rtrim(strtr(base64_encode(random_bytes(32)), '+/', '-_'), '=');
}
function jsonBody(): array
{
    $raw = file_get_contents('php://input') ?: '';
    $ct = $_SERVER['CONTENT_TYPE'] ?? '';
    if (str_contains($ct, 'json')) {
        $d = json_decode($raw, true);
        return is_array($d) ? $d : [];
    }
    parse_str($raw, $d);
    return $d + $_POST;
}

$path = parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH) ?: '/';
$method = $_SERVER['REQUEST_METHOD'];
$fh = fopen($stateFile . '.lock', 'c');
flock($fh, LOCK_EX);
$st = load($stateFile);

switch (true) {
    case $path === '/.well-known/openid-configuration':
        $base = 'http://' . ($_SERVER['HTTP_HOST'] ?? '127.0.0.1:8090');
        out(['issuer' => $base, 'authorization_endpoint' => $base . '/authorize', 'token_endpoint' => $base . '/token', 'userinfo_endpoint' => $base . '/userinfo', 'revocation_endpoint' => $base . '/revoke', 'code_challenge_methods_supported' => ['S256']]);

    case $path === '/authorize':
        $q = $_GET;
        $client = $clients[$q['client_id'] ?? ''] ?? null;
        if ($client === null) {
            out(['error' => 'invalid_client'], 400);
        }
        $redirect = (string)($q['redirect_uri'] ?? '');
        if (!in_array($redirect, $client['redirects'], true)) {
            out(['error' => 'invalid_redirect_uri'], 400);
        }
        if (($q['response_type'] ?? '') !== 'code' || ($q['code_challenge_method'] ?? '') !== 'S256' || empty($q['code_challenge'])) {
            out(['error' => 'invalid_request'], 400);
        }
        $user = $q['user'] ?? null;
        if ($user === null || !isset($users[$user])) {
            header('Content-Type: text/html; charset=utf-8');
            echo '<!doctype html><title>Mock Clerdo Auth</title><h1>Mock auth.clerdo.app</h1><p>Sign in as:</p>';
            foreach ($users as $id => $u) {
                $link = '/authorize?' . http_build_query($q + ['user' => $id]);
                echo '<p><a href="' . htmlspecialchars($link) . '">' . htmlspecialchars($u['name']) . ' (' . htmlspecialchars($id) . ')</a></p>';
            }
            exit;
        }
        $code = tok();
        $st['codes'][$code] = ['user' => $user, 'client' => $q['client_id'], 'challenge' => $q['code_challenge'], 'redirect' => $redirect, 'exp' => time() + 120];
        save($stateFile, $st);
        $sep = str_contains($redirect, '?') ? '&' : '?';
        header('Location: ' . $redirect . $sep . http_build_query(['code' => $code, 'state' => $q['state'] ?? '']), true, 302);
        exit;

    case $path === '/token' && $method === 'POST':
        $b = jsonBody();
        $grant = $b['grant_type'] ?? '';
        if ($grant === 'authorization_code') {
            $c = $st['codes'][$b['code'] ?? ''] ?? null;
            unset($st['codes'][$b['code'] ?? '']);
            save($stateFile, $st);
            if ($c === null || $c['exp'] < time() || $c['client'] !== ($b['client_id'] ?? '') || $c['redirect'] !== ($b['redirect_uri'] ?? '')) {
                out(['error' => 'invalid_grant'], 400);
            }
            $ver = (string)($b['code_verifier'] ?? '');
            $calc = rtrim(strtr(base64_encode(hash('sha256', $ver, true)), '+/', '-_'), '=');
            if (!hash_equals($c['challenge'], $calc)) {
                out(['error' => 'invalid_grant', 'error_description' => 'PKCE verification failed'], 400);
            }
            $family = tok();
            $user = $c['user'];
        } elseif ($grant === 'refresh_token') {
            $rt = (string)($b['refresh_token'] ?? '');
            $r = $st['refresh'][$rt] ?? null;
            if ($r === null) {
                out(['error' => 'invalid_grant'], 400);
            }
            if (!empty($r['used'])) {
                // Reuse detected → revoke the whole family.
                foreach ($st['refresh'] as $k => $v) {
                    if ($v['family'] === $r['family']) {
                        unset($st['refresh'][$k]);
                    }
                }
                foreach ($st['access'] as $k => $v) {
                    if ($v['family'] === $r['family']) {
                        unset($st['access'][$k]);
                    }
                }
                save($stateFile, $st);
                out(['error' => 'invalid_grant', 'error_description' => 'refresh token reuse'], 400);
            }
            if ($r['client'] !== ($b['client_id'] ?? '')) {
                out(['error' => 'invalid_client'], 400);
            }
            $st['refresh'][$rt]['used'] = true;
            $family = $r['family'];
            $user = $r['user'];
        } else {
            out(['error' => 'unsupported_grant_type'], 400);
        }
        $access = tok();
        $refresh = tok();
        $st['access'][$access] = ['user' => $user, 'family' => $family, 'exp' => time() + ACCESS_TTL];
        $st['refresh'][$refresh] = ['user' => $user, 'family' => $family, 'client' => $b['client_id'], 'used' => false];
        // prune
        foreach ($st['access'] as $k => $v) {
            if ($v['exp'] < time() - 3600) {
                unset($st['access'][$k]);
            }
        }
        save($stateFile, $st);
        out(['access_token' => $access, 'token_type' => 'Bearer', 'expires_in' => ACCESS_TTL, 'refresh_token' => $refresh, 'scope' => 'openid profile email']);

    case $path === '/userinfo':
        $h = $_SERVER['HTTP_AUTHORIZATION'] ?? '';
        if (!preg_match('/^Bearer\s+(\S+)$/', $h, $m)) {
            header('WWW-Authenticate: Bearer');
            out(['error' => 'invalid_token'], 401);
        }
        $a = $st['access'][$m[1]] ?? null;
        if ($a === null || $a['exp'] < time()) {
            out(['error' => 'invalid_token'], 401);
        }
        out($users[$a['user']]);

    case $path === '/revoke' && $method === 'POST':
        $b = jsonBody();
        $t = (string)($b['token'] ?? '');
        unset($st['access'][$t], $st['refresh'][$t]);
        save($stateFile, $st);
        out([]);

    // Test helper: mint an access token directly (dev only) — used by api-test for the Bearer checks.
    case $path === '/_dev/token' && $method === 'POST':
        $b = jsonBody();
        $user = $b['user'] ?? 'u-hamzeh';
        if (!isset($users[$user])) {
            out(['error' => 'no_user'], 400);
        }
        $access = tok();
        $st['access'][$access] = ['user' => $user, 'family' => tok(), 'exp' => time() + ACCESS_TTL];
        save($stateFile, $st);
        out(['access_token' => $access, 'expires_in' => ACCESS_TTL]);

    case $path === '/_dev/stats':
        out(['access' => count($st['access']), 'refresh' => count($st['refresh']), 'refresh_used' => count(array_filter($st['refresh'], static fn($r) => !empty($r['used'])))]);

    default:
        out(['error' => 'not_found'], 404);
}
