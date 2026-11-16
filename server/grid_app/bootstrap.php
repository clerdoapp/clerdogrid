<?php
declare(strict_types=1);

/**
 * Grid backend bootstrap: strict errors → JSON, autoload lib/, load config.
 * Required by public_html/api.php. Lives outside the web root.
 */

namespace Grid;

const GRID_APP_DIR = __DIR__;
const GRID_VERSION = '1.0.0';

if (!defined('GRID_MEMORY_SET')) {
    define('GRID_MEMORY_SET', true);
    @ini_set('memory_limit', '512M');
    @ini_set('display_errors', '0');
    @ini_set('log_errors', '1');
    @ini_set('html_errors', '0');
    @ini_set('zend.exception_ignore_args', '1');
}

error_reporting(E_ALL);
date_default_timezone_set('UTC');

spl_autoload_register(static function (string $class): void {
    if (!str_starts_with($class, 'Grid\\')) {
        return;
    }
    $short = substr($class, 5);
    // Several classes live in Core.php; everything else maps 1:1.
    static $coreClasses = ['HttpError', 'Db', 'Config', 'Json', 'Crypto', 'Http', 'Id', 'Clock', 'RateLimit'];
    $file = in_array($short, $coreClasses, true) ? 'Core' : $short;
    $path = __DIR__ . '/lib/' . $file . '.php';
    if (is_file($path)) {
        require_once $path;
    }
});

set_error_handler(static function (int $no, string $str, string $file, int $line): bool {
    if (!(error_reporting() & $no)) {
        return false;
    }
    throw new \ErrorException($str, 0, $no, $file, $line);
});

set_exception_handler(static function (\Throwable $e): void {
    if ($e instanceof HttpError) {
        if (!headers_sent()) {
            Json::error($e->status, $e->errorCode, $e->getMessage(), $e->extra);
        }
        return;
    }
    error_log(sprintf('[grid] %s: %s in %s:%d', get_class($e), $e->getMessage(), $e->getFile(), $e->getLine()));
    if (!headers_sent()) {
        Json::error(500, 'server', 'Internal error');
    }
});

register_shutdown_function(static function (): void {
    $err = error_get_last();
    if ($err !== null && in_array($err['type'], [E_ERROR, E_PARSE, E_CORE_ERROR, E_COMPILE_ERROR], true)) {
        error_log('[grid] fatal: ' . $err['message'] . ' in ' . $err['file'] . ':' . $err['line']);
        if (!headers_sent()) {
            http_response_code(500);
            header('Content-Type: application/json; charset=utf-8');
            echo '{"error":"server","message":"Internal error"}';
        }
    }
});

$configPath = getenv('GRID_CONFIG') ?: __DIR__ . '/config.php';
if (!is_file($configPath)) {
    http_response_code(500);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode([
        'error' => 'server',
        'message' => 'grid_app/config.php is missing. Copy config.sample.php to config.php or run deployment.php?key=...&config=1',
    ], JSON_UNESCAPED_SLASHES);
    exit;
}
$cfg = require $configPath;
if (!is_array($cfg)) {
    http_response_code(500);
    header('Content-Type: application/json; charset=utf-8');
    echo '{"error":"server","message":"config.php must return an array"}';
    exit;
}
Config::load($cfg);
require_once __DIR__ . '/lib/Core.php';
