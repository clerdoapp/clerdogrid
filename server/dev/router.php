<?php
declare(strict_types=1);

/**
 * php -S router for local dev:  PHP_CLI_SERVER_WORKERS=6 php -S 127.0.0.1:8080 dev/router.php
 * Serves dist/ statically, routes /api /auth /share to the front controller, SPA fallback.
 */

$root = dirname(__DIR__);
$dist = dirname($root) . '/dist';
$path = parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH) ?: '/';

putenv('GRID_CONFIG=' . (getenv('GRID_CONFIG') ?: __DIR__ . '/config.dev.php'));

if (preg_match('#^/(api|auth|share)(/|$)#', $path)) {
    require $root . '/public_html/api.php';
    return true;
}

$file = realpath($dist . $path);
if ($path !== '/' && $file !== false && str_starts_with($file, realpath($dist) ?: $dist) && is_file($file)) {
    $ext = strtolower(pathinfo($file, PATHINFO_EXTENSION));
    $types = ['js' => 'text/javascript', 'mjs' => 'text/javascript', 'css' => 'text/css', 'html' => 'text/html; charset=utf-8', 'json' => 'application/json', 'webmanifest' => 'application/manifest+json', 'svg' => 'image/svg+xml', 'png' => 'image/png', 'jpg' => 'image/jpeg', 'ico' => 'image/x-icon', 'woff2' => 'font/woff2', 'woff' => 'font/woff', 'map' => 'application/json', 'txt' => 'text/plain'];
    header('Content-Type: ' . ($types[$ext] ?? 'application/octet-stream'));
    header('X-Content-Type-Options: nosniff');
    header('Cache-Control: ' . (str_starts_with($path, '/assets/') ? 'public, max-age=31536000, immutable' : 'no-cache'));
    readfile($file);
    return true;
}

$index = $dist . '/index.html';
header('Content-Type: text/html; charset=utf-8');
header('Cache-Control: no-store');
if (is_file($index)) {
    readfile($index);
} else {
    http_response_code(503);
    echo '<!doctype html><title>Grid dev</title><p>dist/index.html not found. Run <code>npm run build:cloud</code>.</p>';
}
return true;
