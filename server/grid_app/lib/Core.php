<?php
declare(strict_types=1);

namespace Grid;

use PDO;
use PDOException;
use RuntimeException;

/**
 * Application-level HTTP error. Maps to the JSON error envelope in API.md.
 */
final class HttpError extends RuntimeException
{
    public function __construct(
        public readonly int $status,
        public readonly string $errorCode,
        string $message = '',
        public readonly array $extra = []
    ) {
        parent::__construct($message !== '' ? $message : $errorCode);
    }
}

/**
 * PDO singleton. MariaDB is the primary dialect; sqlite is supported for the
 * local sandbox. All dialect translation happens in Db::sql().
 */
final class Db
{
    private static ?PDO $pdo = null;
    private static string $driver = 'mysql';

    public static function driver(): string
    {
        self::pdo();
        return self::$driver;
    }

    public static function isSqlite(): bool
    {
        return self::driver() === 'sqlite';
    }

    public static function pdo(): PDO
    {
        if (self::$pdo !== null) {
            return self::$pdo;
        }
        $cfg = Config::get('db');
        $driver = (string)($cfg['driver'] ?? 'mysql');
        $opts = [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES => false,
        ];
        try {
            if ($driver === 'sqlite') {
                $path = (string)($cfg['path'] ?? ':memory:');
                if ($path !== ':memory:') {
                    $dir = dirname($path);
                    if (!is_dir($dir)) {
                        @mkdir($dir, 0700, true);
                    }
                }
                $pdo = new PDO('sqlite:' . $path, null, null, $opts);
                $pdo->exec('PRAGMA journal_mode=WAL');
                $pdo->exec('PRAGMA busy_timeout=10000');
                $pdo->exec('PRAGMA foreign_keys=ON');
                self::$driver = 'sqlite';
            } else {
                $dsn = sprintf(
                    'mysql:host=%s;port=%d;dbname=%s;charset=utf8mb4',
                    $cfg['host'] ?? '127.0.0.1',
                    (int)($cfg['port'] ?? 3306),
                    $cfg['name'] ?? ''
                );
                $pdo = new PDO($dsn, (string)($cfg['user'] ?? ''), (string)($cfg['pass'] ?? ''), $opts);
                $pdo->exec("SET NAMES utf8mb4 COLLATE utf8mb4_unicode_ci");
                $pdo->exec("SET SESSION sql_mode='STRICT_ALL_TABLES,NO_ENGINE_SUBSTITUTION'");
                self::$driver = 'mysql';
            }
        } catch (PDOException $e) {
            error_log('[grid] db connect failed: ' . $e->getMessage());
            throw new HttpError(500, 'server', 'Database unavailable');
        }
        self::$pdo = $pdo;
        return $pdo;
    }

    /**
     * Translate MariaDB-flavoured SQL to the active dialect.
     * Primary SQL is written for MariaDB; this is the one place sqlite differs.
     */
    public static function sql(string $sql): string
    {
        if (!self::isSqlite()) {
            return $sql;
        }
        $map = [
            'BIGINT AUTO_INCREMENT PRIMARY KEY' => 'INTEGER PRIMARY KEY AUTOINCREMENT',
            'INT AUTO_INCREMENT PRIMARY KEY' => 'INTEGER PRIMARY KEY AUTOINCREMENT',
            ' FOR UPDATE' => '',
            'NOW()' => "datetime('now')",
            'UTC_TIMESTAMP()' => "datetime('now')",
            'INSERT IGNORE INTO' => 'INSERT OR IGNORE INTO',
            'LONGTEXT' => 'TEXT',
            'MEDIUMTEXT' => 'TEXT',
            'TINYINT(1)' => 'INTEGER',
            'TINYINT' => 'INTEGER',
            'DATETIME' => 'TEXT',
            'BLOB' => 'BLOB',
            ' ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci' => '',
            ' ENGINE=InnoDB' => '',
        ];
        $sql = str_replace(array_keys($map), array_values($map), $sql);
        // ENUM('a','b') -> TEXT
        $sql = preg_replace("/ENUM\\([^)]*\\)/", 'TEXT', $sql) ?? $sql;
        // Inline INDEX definitions inside CREATE TABLE are not valid in sqlite; strip them.
        $sql = preg_replace('/,\s*(UNIQUE\s+)?INDEX\s+\w+\s*\([^)]*\)/i', '', $sql) ?? $sql;
        // DATE_SUB(NOW(), INTERVAL n SECOND|MINUTE|DAY)
        $sql = preg_replace_callback(
            "/DATE_SUB\\(datetime\\('now'\\), INTERVAL (\\d+|\\?) (SECOND|MINUTE|HOUR|DAY)\\)/i",
            static fn(array $m): string => sprintf("datetime('now', '-' || %s || ' %s')", $m[1] === '?' ? '?' : $m[1], strtolower($m[2])),
            $sql
        ) ?? $sql;
        $sql = preg_replace_callback(
            "/DATE_ADD\\(datetime\\('now'\\), INTERVAL (\\d+|\\?) (SECOND|MINUTE|HOUR|DAY)\\)/i",
            static fn(array $m): string => sprintf("datetime('now', '+' || %s || ' %s')", $m[1] === '?' ? '?' : $m[1], strtolower($m[2])),
            $sql
        ) ?? $sql;
        // ON DUPLICATE KEY UPDATE col=VALUES(col), ...  ->  ON CONFLICT DO UPDATE SET col=excluded.col
        if (preg_match('/\sON DUPLICATE KEY UPDATE\s+(.+)$/is', $sql, $m)) {
            $assign = preg_replace('/VALUES\((\w+)\)/i', 'excluded.$1', $m[1]) ?? $m[1];
            $sql = preg_replace('/\sON DUPLICATE KEY UPDATE\s+.+$/is', ' ON CONFLICT DO UPDATE SET ' . $assign, $sql) ?? $sql;
        }
        return $sql;
    }

    /** Prepare + execute with dialect translation. */
    public static function run(string $sql, array $params = []): \PDOStatement
    {
        $st = self::pdo()->prepare(self::sql($sql));
        foreach ($params as $k => $v) {
            $key = is_int($k) ? $k + 1 : $k;
            if (is_int($v)) {
                $st->bindValue($key, $v, PDO::PARAM_INT);
            } elseif (is_bool($v)) {
                $st->bindValue($key, $v ? 1 : 0, PDO::PARAM_INT);
            } elseif ($v === null) {
                $st->bindValue($key, null, PDO::PARAM_NULL);
            } else {
                $st->bindValue($key, (string)$v, PDO::PARAM_STR);
            }
        }
        $st->execute();
        return $st;
    }

    public static function one(string $sql, array $params = []): ?array
    {
        $row = self::run($sql, $params)->fetch();
        return $row === false ? null : $row;
    }

    public static function all(string $sql, array $params = []): array
    {
        return self::run($sql, $params)->fetchAll();
    }

    public static function exec(string $sql, array $params = []): int
    {
        return self::run($sql, $params)->rowCount();
    }

    public static function lastId(): int
    {
        return (int)self::pdo()->lastInsertId();
    }

    /** Run $fn inside a transaction (sqlite uses IMMEDIATE to serialise writers). */
    public static function tx(callable $fn): mixed
    {
        $pdo = self::pdo();
        if ($pdo->inTransaction()) {
            return $fn();
        }
        if (self::isSqlite()) {
            $pdo->exec('BEGIN IMMEDIATE');
        } else {
            $pdo->beginTransaction();
        }
        try {
            $r = $fn();
            $pdo->inTransaction() ? $pdo->commit() : $pdo->exec('COMMIT');
            return $r;
        } catch (\Throwable $e) {
            try {
                $pdo->inTransaction() ? $pdo->rollBack() : $pdo->exec('ROLLBACK');
            } catch (\Throwable) {
            }
            throw $e;
        }
    }

    /** Does a table exist? */
    public static function tableExists(string $table): bool
    {
        if (self::isSqlite()) {
            return self::one("SELECT name FROM sqlite_master WHERE type='table' AND name=?", [$table]) !== null;
        }
        return self::one('SELECT TABLE_NAME FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?', [$table]) !== null;
    }

    /** Column names of a table. */
    public static function columns(string $table): array
    {
        if (self::isSqlite()) {
            return array_map(static fn($r) => $r['name'], self::all("PRAGMA table_info(" . $table . ")"));
        }
        return array_map(static fn($r) => $r['COLUMN_NAME'], self::all('SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?', [$table]));
    }
}

/** Config access (config.php returns an array). */
final class Config
{
    private static array $cfg = [];

    public static function load(array $cfg): void
    {
        self::$cfg = $cfg;
    }

    public static function get(string $key, mixed $default = null): mixed
    {
        $cur = self::$cfg;
        foreach (explode('.', $key) as $part) {
            if (!is_array($cur) || !array_key_exists($part, $cur)) {
                return $default;
            }
            $cur = $cur[$part];
        }
        return $cur;
    }

    public static function all(): array
    {
        return self::$cfg;
    }
}

final class Json
{
    public const FLAGS = JSON_THROW_ON_ERROR | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE;

    public static function encode(mixed $v): string
    {
        return json_encode($v, self::FLAGS);
    }

    /** Decode to associative arrays. Throws JsonException. */
    public static function decode(string $s, int $depth = 64): mixed
    {
        return json_decode($s, true, $depth, JSON_THROW_ON_ERROR | JSON_BIGINT_AS_STRING);
    }

    /** Decode to stdClass objects (preserves empty {} on re-encode); used for echoing stored documents. */
    public static function decodeObj(string $s, int $depth = 64): mixed
    {
        return json_decode($s, false, $depth, JSON_THROW_ON_ERROR);
    }

    public static function send(mixed $payload, int $status = 200): never
    {
        http_response_code($status);
        header('Content-Type: application/json; charset=utf-8');
        header('Cache-Control: no-store');
        if ($status !== 204) {
            echo self::encode($payload);
        }
        exit;
    }

    public static function error(int $status, string $code, string $message = '', array $extra = []): never
    {
        self::send(array_merge(['error' => $code, 'message' => $message !== '' ? $message : $code], $extra), $status);
    }
}

/** AES-256-GCM with a key id prefix for rotation: "v1:" . base64(iv . tag . ciphertext). */
final class Crypto
{
    private static ?array $keys = null;
    private static string $activeId = 'k1';

    private static function keys(): array
    {
        if (self::$keys !== null) {
            return self::$keys;
        }
        $raw = Config::get('encryption_key');
        $keys = [];
        if (is_array($raw)) {
            // ['k2' => 'base64...', 'k1' => 'base64...'] — first entry is active.
            foreach ($raw as $id => $b64) {
                $keys[(string)$id] = self::decodeKey((string)$b64);
            }
            self::$activeId = (string)array_key_first($raw);
        } else {
            $keys['k1'] = self::decodeKey((string)$raw);
            self::$activeId = 'k1';
        }
        self::$keys = $keys;
        return $keys;
    }

    private static function decodeKey(string $b64): string
    {
        $k = base64_decode($b64, true);
        if ($k === false || strlen($k) !== 32) {
            throw new RuntimeException('encryption_key must be base64 of 32 bytes');
        }
        return $k;
    }

    public static function encrypt(string $plain): string
    {
        $keys = self::keys();
        $key = $keys[self::$activeId];
        $iv = random_bytes(12);
        if (function_exists('sodium_crypto_aead_aes256gcm_is_available') && sodium_crypto_aead_aes256gcm_is_available()) {
            $ct = sodium_crypto_aead_aes256gcm_encrypt($plain, self::$activeId, $iv, $key);
            // libsodium appends the 16-byte tag; normalise to iv.tag.ct
            $tag = substr($ct, -16);
            $ct = substr($ct, 0, -16);
        } else {
            $tag = '';
            $ct = openssl_encrypt($plain, 'aes-256-gcm', $key, OPENSSL_RAW_DATA, $iv, $tag, self::$activeId, 16);
            if ($ct === false) {
                throw new RuntimeException('encrypt failed');
            }
        }
        return self::$activeId . ':' . base64_encode($iv . $tag . $ct);
    }

    public static function decrypt(string $blob): string
    {
        $pos = strpos($blob, ':');
        if ($pos === false) {
            throw new RuntimeException('bad ciphertext');
        }
        $id = substr($blob, 0, $pos);
        $keys = self::keys();
        if (!isset($keys[$id])) {
            throw new RuntimeException('unknown key id');
        }
        $raw = base64_decode(substr($blob, $pos + 1), true);
        if ($raw === false || strlen($raw) < 28) {
            throw new RuntimeException('bad ciphertext');
        }
        $iv = substr($raw, 0, 12);
        $tag = substr($raw, 12, 16);
        $ct = substr($raw, 28);
        $plain = openssl_decrypt($ct, 'aes-256-gcm', $keys[$id], OPENSSL_RAW_DATA, $iv, $tag, $id);
        if ($plain === false) {
            throw new RuntimeException('decrypt failed');
        }
        return $plain;
    }
}

final class Http
{
    public static function method(): string
    {
        return strtoupper($_SERVER['REQUEST_METHOD'] ?? 'GET');
    }

    public static function path(): string
    {
        $uri = $_SERVER['REQUEST_URI'] ?? '/';
        $p = parse_url($uri, PHP_URL_PATH) ?: '/';
        $p = rawurldecode($p);
        return '/' . ltrim($p, '/');
    }

    public static function header(string $name): ?string
    {
        $key = 'HTTP_' . strtoupper(str_replace('-', '_', $name));
        $v = $_SERVER[$key] ?? $_SERVER['REDIRECT_' . $key] ?? null;
        if ($v === null && $key === 'HTTP_AUTHORIZATION' && function_exists('apache_request_headers')) {
            $v = apache_request_headers()['Authorization'] ?? null;
        }
        return $v === null ? null : (string)$v;
    }

    public static function isHttps(): bool
    {
        if (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') {
            return true;
        }
        return ($_SERVER['HTTP_X_FORWARDED_PROTO'] ?? '') === 'https';
    }

    /** scheme://host of the current request (host from config base_url when set, else Host header). */
    public static function origin(): string
    {
        $base = (string)Config::get('app.base_url', '');
        if ($base !== '') {
            return rtrim($base, '/');
        }
        $host = $_SERVER['HTTP_HOST'] ?? 'localhost';
        return (self::isHttps() ? 'https' : 'http') . '://' . $host;
    }

    /** True when Origin (or Sec-Fetch-Site) proves a same-origin request. */
    public static function isSameOrigin(): bool
    {
        $origin = self::header('Origin');
        if ($origin !== null && $origin !== '' && $origin !== 'null') {
            $o = strtolower(rtrim($origin, '/'));
            $allowed = [strtolower(self::origin())];
            $host = $_SERVER['HTTP_HOST'] ?? '';
            if ($host !== '') {
                $allowed[] = strtolower((self::isHttps() ? 'https' : 'http') . '://' . $host);
            }
            return in_array($o, $allowed, true);
        }
        $sfs = self::header('Sec-Fetch-Site');
        return $sfs === 'same-origin';
    }

    public static function clientIp(): string
    {
        // Shared hosting sits behind LiteSpeed directly; trust REMOTE_ADDR.
        $ip = $_SERVER['REMOTE_ADDR'] ?? '0.0.0.0';
        return substr($ip, 0, 45);
    }

    public static function securityHeaders(): void
    {
        header("Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com data:; img-src 'self' data: blob:; connect-src 'self'; frame-ancestors 'self' https://kelvo.clerdo.app; base-uri 'self'; form-action 'self' https://auth.clerdo.app");
        header('Referrer-Policy: strict-origin-when-cross-origin');
        header('X-Content-Type-Options: nosniff');
        header('Permissions-Policy: camera=(), microphone=(), geolocation=()');
    }

    public static function redirect(string $to, int $status = 302): never
    {
        http_response_code($status);
        header('Location: ' . $to);
        header('Cache-Control: no-store');
        exit;
    }

    public static function setCookie(string $name, string $value, int $expires): void
    {
        setcookie($name, $value, [
            'expires' => $expires,
            'path' => '/',
            'secure' => self::isHttps(),
            'httponly' => true,
            'samesite' => 'Lax',
        ]);
    }

    /** Validate a relative "next" path; refuse protocol-relative and absolute URLs. */
    public static function safeNext(?string $next): string
    {
        if ($next === null || $next === '' || $next[0] !== '/' || str_starts_with($next, '//') || str_contains($next, '\\')) {
            return '/';
        }
        return substr($next, 0, 512);
    }
}

final class Id
{
    public static function uuid(): string
    {
        $b = random_bytes(16);
        $b[6] = chr((ord($b[6]) & 0x0f) | 0x40);
        $b[8] = chr((ord($b[8]) & 0x3f) | 0x80);
        $h = bin2hex($b);
        return sprintf('%s-%s-%s-%s-%s', substr($h, 0, 8), substr($h, 8, 4), substr($h, 12, 4), substr($h, 16, 4), substr($h, 20, 12));
    }

    /** URL-safe random token. 32 bytes → 43 chars (base64url, no padding). */
    public static function token(int $bytes = 32): string
    {
        return rtrim(strtr(base64_encode(random_bytes($bytes)), '+/', '-_'), '=');
    }

    public static function hex(int $bytes = 32): string
    {
        return bin2hex(random_bytes($bytes));
    }
}

final class Clock
{
    public static function now(): int
    {
        return time();
    }

    /** 'Y-m-d H:i:s' in UTC — the DATETIME representation used in every table. */
    public static function sql(?int $ts = null): string
    {
        return gmdate('Y-m-d H:i:s', $ts ?? time());
    }

    public static function iso(?string $sql): ?string
    {
        if ($sql === null || $sql === '') {
            return null;
        }
        $ts = strtotime($sql . ' UTC');
        return $ts === false ? null : gmdate('Y-m-d\TH:i:s\Z', $ts);
    }

    public static function fromSql(string $sql): int
    {
        $ts = strtotime($sql . ' UTC');
        return $ts === false ? 0 : $ts;
    }
}

/** Per-key per-minute bucket in grid_ratelimit. */
final class RateLimit
{
    public static function hit(string $key, int $limitPerMinute): void
    {
        $bucket = (int)floor(time() / 60);
        $k = substr($key, 0, 120);
        Db::exec(
            'INSERT INTO grid_ratelimit (rl_key, bucket, hits) VALUES (?, ?, 1) ON DUPLICATE KEY UPDATE hits = hits + 1',
            [$k, $bucket]
        );
        $row = Db::one('SELECT hits FROM grid_ratelimit WHERE rl_key = ? AND bucket = ?', [$k, $bucket]);
        if (mt_rand(1, 50) === 1) {
            Db::exec('DELETE FROM grid_ratelimit WHERE bucket < ?', [$bucket - 5]);
        }
        if ($row !== null && (int)$row['hits'] > $limitPerMinute) {
            header('Retry-After: ' . (60 - (time() % 60)));
            throw new HttpError(429, 'rate_limited', 'Too many requests');
        }
    }
}
