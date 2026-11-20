<?php
declare(strict_types=1);

namespace Grid;

/**
 * Locker client. Grid stores each workbook's .xlsx in the user's "Grid" folder in Locker, calling
 * Locker server-to-server with the user's own Clerdo ID access token, so Locker only ever sees
 * requests made for that user. Off when locker.url is empty.
 */
final class Locker
{
    /** Locker refused the save: someone changed the file there first. */
    public const CONFLICT = 'conflict';
    /** The file is gone from Locker for good. */
    public const GONE = 'gone';

    private static ?string $folder = null;

    public static function enabled(): bool
    {
        return (string)Config::get('locker.url', '') !== '';
    }

    /**
     * Locker's library on this server (domains/locker.clerdo.app/locker_app/library.php, or config locker.path):
     * saves then happen in this same request, with no web call. Shared hosting refuses the extra requests a
     * web call needs ("server busy"), so this is the normal way on DirectAdmin; the web API is the fallback.
     */
    private static function library(): bool
    {
        static $ok = null;
        if ($ok !== null) {
            return $ok;
        }
        $dir = (string)Config::get('locker.path', '') ?: dirname(__DIR__, 3) . '/locker.clerdo.app/locker_app';
        $file = rtrim($dir, '/') . '/library.php';
        $ok = false;
        try {
            // @: a host with open_basedir answers "not allowed" for the other folder; then the web API is used.
            if (@is_file($file) && @is_file(rtrim($dir, '/') . '/lib/Library.php')) {
                require_once $file;
                $ok = class_exists(\Locker\Library::class);
            }
        } catch (\Throwable $e) {
            error_log('[grid] locker library could not load: ' . $e->getMessage());
        }
        return $ok;
    }

    /** Runs a library call, turning Locker's errors into this app's: 507 quota, else storage trouble. */
    private static function direct(callable $fn): mixed
    {
        try {
            return $fn();
        } catch (\Locker\HttpError $e) {
            if ($e->status === 507) {
                throw new HttpError(507, 'quota', 'Your Locker is full. Free up space in Locker to keep saving there.');
            }
            error_log('[grid] locker library: ' . $e->status . ' ' . $e->getMessage());
            throw new HttpError(502, 'storage', 'Locker could not be reached');
        } catch (\Throwable $e) {
            error_log('[grid] locker library: ' . get_class($e) . ': ' . $e->getMessage());
            throw new HttpError(502, 'storage', 'Locker could not be reached');
        }
    }

    private static function uid(array $s): string
    {
        return (string)($s['user_id'] ?? '');
    }

    /** Creates the workbook file. Returns ['id' => ..., 'rev' => ...]. */
    public static function create(array $s, string $name, string $bytes): array
    {
        if (self::library()) {
            return self::direct(static fn(): array => \Locker\Library::create(self::uid($s), 'grid', self::fileName($name), $bytes));
        }
        $folder = self::folder($s);
        $r = self::call($s, 'POST', '/api/files?' . http_build_query(['parent' => $folder, 'name' => self::fileName($name)]), $bytes);
        self::expect($r, [201]);
        return self::idRev($r);
    }

    /** Saves new content on top of $rev. Returns ['id', 'rev'], self::CONFLICT or self::GONE. */
    public static function replace(array $s, string $id, int $rev, string $bytes): array|string
    {
        if (self::library()) {
            try {
                return \Locker\Library::replace(self::uid($s), 'grid', $id, $rev, $bytes);
            } catch (\Locker\HttpError $e) {
                if ($e->status === 409) {
                    return self::CONFLICT;
                }
                if ($e->status === 404) {
                    return self::GONE;
                }
                return self::direct(static fn() => throw $e);
            }
        }
        $r = self::call($s, 'PUT', '/api/items/' . rawurlencode($id) . '/content', $bytes, ['If-Match: "' . $rev . '"', 'Content-Type: application/octet-stream']);
        if ($r['status'] === 409) {
            return self::CONFLICT;
        }
        if ($r['status'] === 404 || ($r['status'] === 422 && str_contains($r['body'], 'Bin'))) {
            return self::GONE;
        }
        self::expect($r, [200]);
        return self::idRev($r);
    }

    /** Current revision of the file, or null when it no longer exists. */
    public static function rev(array $s, string $id): ?int
    {
        if (self::library()) {
            return self::direct(static fn(): ?int => \Locker\Library::rev(self::uid($s), $id));
        }
        $r = self::call($s, 'GET', '/api/items/' . rawurlencode($id), null, [], 6);
        if ($r['status'] === 404) {
            return null;
        }
        self::expect($r, [200]);
        return (int)(self::json($r)['item']['rev'] ?? 0);
    }

    /** Current content and revision. */
    public static function fetch(array $s, string $id): array
    {
        if (self::library()) {
            return self::direct(static fn(): array => \Locker\Library::fetch(self::uid($s), $id));
        }
        $r = self::call($s, 'GET', '/api/items/' . rawurlencode($id) . '/content');
        self::expect($r, [200]);
        return ['bytes' => $r['body'], 'rev' => self::etag($r)];
    }

    /** Content of an earlier version, or null when Locker no longer keeps it. */
    public static function version(array $s, string $id, int $rev): ?string
    {
        if (self::library()) {
            return self::direct(static fn(): ?string => \Locker\Library::version(self::uid($s), $id, $rev));
        }
        $r = self::call($s, 'GET', '/api/items/' . rawurlencode($id) . '/versions/' . $rev . '/content');
        if ($r['status'] === 404) {
            return null;
        }
        self::expect($r, [200]);
        return $r['body'];
    }

    public static function rename(array $s, string $id, string $name): void
    {
        if (self::library()) {
            self::direct(static fn() => \Locker\Library::rename(self::uid($s), $id, self::fileName($name)));
            return;
        }
        $r = self::call($s, 'PATCH', '/api/items/' . rawurlencode($id), Json::encode(['name' => self::fileName($name)]), ['Content-Type: application/json']);
        self::expect($r, [200, 404]);
    }

    /** Sends the file to the Locker Bin, where the user can still get it back for 30 days. */
    public static function trash(array $s, string $id): void
    {
        if (self::library()) {
            self::direct(static fn() => \Locker\Library::trash(self::uid($s), $id));
            return;
        }
        $r = self::call($s, 'POST', '/api/items/' . rawurlencode($id) . '/trash', '');
        self::expect($r, [200, 404]);
    }

    private static function folder(array $s): string
    {
        if (self::$folder === null) {
            $r = self::call($s, 'POST', '/api/apps/grid/folder', '');
            self::expect($r, [200, 201]);
            self::$folder = (string)(self::json($r)['item']['id'] ?? '');
        }
        return self::$folder;
    }

    /** "Q3 / Q4" → "Q3 - Q4.xlsx": Locker names can't hold slashes. */
    private static function fileName(string $name): string
    {
        $n = trim(preg_replace('#[/\\\\]+#', '-', $name) ?? '') ?: 'Book1';
        return mb_substr($n, 0, 240) . '.xlsx';
    }

    private static function idRev(array $r): array
    {
        $item = self::json($r)['item'] ?? [];
        return ['id' => (string)($item['id'] ?? ''), 'rev' => (int)($item['rev'] ?? 0)];
    }

    private static function etag(array $r): int
    {
        return preg_match('/"(\d+)"/', $r['headers']['etag'] ?? '', $m) ? (int)$m[1] : 0;
    }

    private static function json(array $r): array
    {
        try {
            $d = Json::decode($r['body']);
        } catch (\JsonException) {
            return [];
        }
        return is_array($d) ? $d : [];
    }

    /** @param int[] $ok */
    private static function expect(array $r, array $ok): void
    {
        if (in_array($r['status'], $ok, true)) {
            return;
        }
        if ($r['status'] === 507) {
            throw new HttpError(507, 'quota', 'Your Locker is full. Free up space in Locker to keep saving there.');
        }
        error_log('[grid] locker answered ' . $r['status'] . ': ' . substr($r['body'], 0, 200));
        throw new HttpError(502, 'storage', 'Locker could not be reached');
    }

    /**
     * One request to Locker as the user. If Locker answers 401 (the access token expired early or was revoked),
     * the session's token is renewed once and the request repeated.
     *
     * @return array{status:int, body:string, headers:array<string,string>}
     */
    private static function call(array $s, string $method, string $path, ?string $body = null, array $headers = [], int $timeout = 60): array
    {
        $note = self::appNote($s);
        if ($note !== null) {
            return self::send('ClerdoApp ' . $note, $method, $path, $body, $headers, $timeout);
        }
        $token = Auth::accessToken($s);
        $r = self::send($token === null ? null : 'Bearer ' . $token, $method, $path, $body, $headers, $timeout);
        if ($r['status'] === 401) {
            $token = Auth::renewAccessToken($s);
            if ($token !== null) {
                $r = self::send('Bearer ' . $token, $method, $path, $body, $headers, $timeout);
            }
        }
        return $r;
    }

    /**
     * With a key shared with Locker (config locker.app_key, set by Locker's deployment.php?apps=1): a two-minute
     * note signed with it, saying who this request is for. Locker checks it without asking Clerdo ID.
     */
    private static function appNote(array $s): ?string
    {
        $key = (string)Config::get('locker.app_key', '');
        if (strlen($key) < 32 || (string)($s['user_id'] ?? '') === '') {
            return null;
        }
        $b64 = static fn(string $b): string => rtrim(strtr(base64_encode($b), '+/', '-_'), '=');
        $payload = $b64(Json::encode([
            'app' => 'grid',
            'sub' => (string)$s['user_id'],
            'name' => (string)($s['user_name'] ?? ''),
            'email' => (string)($s['user_email'] ?? ''),
            'avatar' => (string)($s['user_avatar'] ?? ''),
            'plan' => Auth::plan($s),
            'exp' => time() + 120,
        ]));
        return $payload . '.' . $b64(hash_hmac('sha256', $payload, $key, true));
    }

    /** @return array{status:int, body:string, headers:array<string,string>} */
    private static function send(?string $authorization, string $method, string $path, ?string $body, array $headers, int $timeout = 60): array
    {
        if ($authorization === null) {
            error_log('[grid] locker: no access token in this session (signed in with an app token, or the session could not be read)');
            throw new HttpError(502, 'storage', 'Locker could not be reached');
        }
        static $ch = null;
        $ch ??= curl_init();
        if ($ch === false) {
            $ch = null;
            throw new HttpError(502, 'storage', 'Locker could not be reached');
        }
        curl_reset($ch);
        $got = [];
        $opts = [
            CURLOPT_URL => rtrim((string)Config::get('locker.url'), '/') . $path,
            CURLOPT_CUSTOMREQUEST => $method,
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_FOLLOWLOCATION => false,
            CURLOPT_DNS_CACHE_TIMEOUT => 600,
            CURLOPT_CONNECTTIMEOUT => min(8, $timeout),
            CURLOPT_TIMEOUT => $timeout,
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_PROXY => '',
            CURLOPT_HTTPHEADER => array_merge([
                'Authorization: ' . $authorization,
                'X-Clerdo-App: grid',
                'Accept: application/json',
                'User-Agent: Grid/' . GRID_VERSION,
            ], $body !== null && $headers === [] ? ['Content-Type: application/octet-stream'] : [], $headers),
            CURLOPT_HEADERFUNCTION => static function ($c, string $line) use (&$got): int {
                $i = strpos($line, ':');
                if ($i !== false) {
                    $got[strtolower(trim(substr($line, 0, $i)))] = trim(substr($line, $i + 1));
                }
                return strlen($line);
            },
        ];
        if ($body !== null) {
            $opts[CURLOPT_POSTFIELDS] = $body;
        }
        curl_setopt_array($ch, $opts);
        @set_time_limit(120);
        $url = (string)$opts[CURLOPT_URL];
        $pinned = DnsCache::options($url);
        if ($pinned !== []) {
            curl_setopt_array($ch, $pinned);
        }
        $out = curl_exec($ch);
        if ($out === false && $pinned !== []) {
            DnsCache::unpin($ch, $url);
            $got = [];
            $out = curl_exec($ch);
        }
        if ($out === false && in_array(curl_errno($ch), [6, 7, 28], true)) {
            $got = [];
            curl_setopt($ch, CURLOPT_IPRESOLVE, CURL_IPRESOLVE_V4);
            $out = curl_exec($ch);
        }
        if ($out === false) {
            error_log('[grid] locker http error: ' . curl_error($ch));
            throw new HttpError(502, 'storage', 'Locker could not be reached');
        }
        DnsCache::remember($ch, $url);
        return ['status' => (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE), 'body' => (string)$out, 'headers' => $got];
    }
}
