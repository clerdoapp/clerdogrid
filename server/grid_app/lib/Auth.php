<?php
declare(strict_types=1);

namespace Grid;

/**
 * BFF auth against auth.clerdo.app: OAuth 2.0 authorization code + PKCE.
 * Tokens are stored AES-256-GCM encrypted in grid_sessions; the browser only
 * holds the opaque grid_sid cookie. Native apps send `Authorization: Bearer`.
 */
final class Auth
{
    public const COOKIE = 'grid_sid';
    public const SESSION_TTL = 30 * 86400;
    public const IDLE_TTL = 14 * 86400;
    public const REFRESH_SKEW = 60;
    public const STATE_TTL = 600;
    public const BEARER_CACHE_TTL = 300;

    /** @var array|null resolved session row for this request (with 'kind' = cookie|bearer) */
    private static ?array $session = null;
    private static bool $resolved = false;

    // ------------------------------------------------------------------ routes

    public static function login(array $p, array $body): void
    {
        RateLimit::hit('auth:' . Http::clientIp(), (int)Config::get('limits.auth_per_minute', 30));
        $next = Http::safeNext($_GET['next'] ?? '/');
        $verifier = Id::token(48);
        $challenge = rtrim(strtr(base64_encode(hash('sha256', $verifier, true)), '+/', '-_'), '=');
        $state = Id::token(32);

        Db::exec('DELETE FROM grid_auth_states WHERE expires_at < ?', [Clock::sql()]);
        Db::exec(
            'INSERT INTO grid_auth_states (state, verifier, next_path, created_at, expires_at) VALUES (?, ?, ?, ?, ?)',
            [$state, $verifier, $next, Clock::sql(), Clock::sql(time() + self::STATE_TTL)]
        );

        $q = http_build_query([
            'client_id' => Config::get('auth.client_id', 'grid-web'),
            'response_type' => 'code',
            'code_challenge' => $challenge,
            'code_challenge_method' => 'S256',
            'redirect_uri' => Config::get('auth.redirect_uri'),
            'state' => $state,
            'scope' => Config::get('auth.scope', 'openid profile email'),
        ], '', '&', PHP_QUERY_RFC3986);
        Http::redirect(Config::get('auth.authorize_url') . '?' . $q);
    }

    public static function callback(array $p, array $body): void
    {
        RateLimit::hit('auth:' . Http::clientIp(), (int)Config::get('limits.auth_per_minute', 30));
        $code = (string)($_GET['code'] ?? '');
        $state = (string)($_GET['state'] ?? '');
        if ($code === '' || $state === '' || strlen($state) > 128) {
            throw new HttpError(400, 'validation', 'Missing code or state');
        }
        if (isset($_GET['error'])) {
            throw new HttpError(401, 'unauthenticated', 'Authorization was denied');
        }
        $row = Db::one('SELECT * FROM grid_auth_states WHERE state = ?', [$state]);
        // Single use.
        Db::exec('DELETE FROM grid_auth_states WHERE state = ?', [$state]);
        if ($row === null || Clock::fromSql($row['expires_at']) < time()) {
            throw new HttpError(400, 'validation', 'Login expired, please try again');
        }

        $tok = self::tokenRequest([
            'grant_type' => 'authorization_code',
            'code' => $code,
            'redirect_uri' => Config::get('auth.redirect_uri'),
            'client_id' => Config::get('auth.client_id', 'grid-web'),
            'code_verifier' => $row['verifier'],
        ]);
        if ($tok === null || empty($tok['access_token'])) {
            throw new HttpError(401, 'unauthenticated', 'Token exchange failed');
        }
        $user = self::userinfo((string)$tok['access_token']);
        if ($user === null) {
            throw new HttpError(401, 'unauthenticated', 'Could not load profile');
        }

        $sid = Id::hex(32); // 64 hex chars
        $csrf = Id::hex(32);
        $expiresIn = max(30, (int)($tok['expires_in'] ?? 600));
        Db::exec(
            'INSERT INTO grid_sessions (id, kind, user_id, user_name, user_email, user_avatar, plan, access_enc, refresh_enc, access_expires, csrf, created_at, last_seen, revoked)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0)',
            [
                $sid, 'cookie', $user['id'], $user['name'], $user['email'], $user['avatar'], $user['plan'],
                Crypto::encrypt((string)$tok['access_token']),
                isset($tok['refresh_token']) ? Crypto::encrypt((string)$tok['refresh_token']) : null,
                Clock::sql(time() + $expiresIn), $csrf, Clock::sql(), Clock::sql(),
            ]
        );
        Http::setCookie(self::COOKIE, $sid, time() + self::SESSION_TTL);
        Http::redirect(Http::safeNext($row['next_path']));
    }

    public static function logout(array $p, array $body): void
    {
        $s = self::resolveSession();
        if ($s !== null) {
            if (($s['kind'] ?? 'cookie') === 'cookie') {
                self::requireCsrf($s);
                // Best effort revoke at the IdP.
                foreach (['refresh_enc' => 'refresh_token', 'access_enc' => 'access_token'] as $col => $hint) {
                    if (!empty($s[$col])) {
                        try {
                            self::revokeAtIdp(Crypto::decrypt($s[$col]), $hint);
                        } catch (\Throwable $e) {
                            error_log('[grid] revoke failed: ' . $e->getMessage());
                        }
                    }
                }
            }
            Db::exec('UPDATE grid_sessions SET revoked = 1, access_enc = NULL, refresh_enc = NULL WHERE id = ?', [$s['id']]);
        }
        Http::setCookie(self::COOKIE, '', time() - 3600);
        Json::send(null, 204);
    }

    public static function me(array $p, array $body): void
    {
        $s = self::requireUser();
        $plan = self::plan($s);
        Json::send([
            'user' => [
                'id' => $s['user_id'],
                'name' => $s['user_name'],
                'email' => $s['user_email'],
                'avatar' => $s['user_avatar'],
                'plan' => $plan,
            ],
            'csrf' => ($s['kind'] ?? 'cookie') === 'bearer' ? null : $s['csrf'],
            'limits' => [
                'projects' => $plan === 'pro' ? null : (int)Config::get('limits.free_projects', 5),
                'docBytes' => self::docLimit($plan),
            ],
        ]);
    }

    // --------------------------------------------------------------- sessions

    public static function plan(array $s): string
    {
        if ((bool)Config::get('pro.everyone_has_pro', true)) {
            return 'pro';
        }
        return ($s['plan'] ?? 'free') === 'pro' ? 'pro' : 'free';
    }

    public static function docLimit(string $plan): int
    {
        return $plan === 'pro'
            ? (int)Config::get('limits.pro_doc_bytes', 25 * 1024 * 1024)
            : (int)Config::get('limits.free_doc_bytes', 16 * 1024 * 1024);
    }

    /** Session row or 401. Enforces CSRF + Origin on mutating cookie requests. */
    public static function requireUser(): array
    {
        $s = self::resolveSession();
        if ($s === null) {
            throw new HttpError(401, 'unauthenticated', 'Sign in required');
        }
        if (in_array(Http::method(), ['POST', 'PUT', 'PATCH', 'DELETE'], true) && ($s['kind'] ?? 'cookie') === 'cookie') {
            self::requireCsrf($s);
        }
        return $s;
    }

    private static function requireCsrf(array $s): void
    {
        $token = Http::header('X-CSRF-Token');
        if ($token === null || !hash_equals((string)$s['csrf'], $token) || !Http::isSameOrigin()) {
            throw new HttpError(403, 'csrf', 'Missing or invalid CSRF token / origin');
        }
    }

    /** Resolve the current session from cookie or Bearer. Cached per request. */
    public static function resolveSession(): ?array
    {
        if (self::$resolved) {
            return self::$session;
        }
        self::$resolved = true;
        $sid = $_COOKIE[self::COOKIE] ?? null;
        if (is_string($sid) && preg_match('/^[a-f0-9]{64}$/', $sid)) {
            self::$session = self::resolveCookie($sid);
            return self::$session;
        }
        $authz = Http::header('Authorization');
        if ($authz !== null && preg_match('/^Bearer\s+([A-Za-z0-9\-._~+\/=]{16,2048})$/', $authz, $m)) {
            self::$session = self::resolveBearer($m[1]);
            return self::$session;
        }
        return null;
    }

    private static function resolveCookie(string $sid): ?array
    {
        $s = Db::one('SELECT * FROM grid_sessions WHERE id = ? AND kind = ?', [$sid, 'cookie']);
        if ($s === null || (int)$s['revoked'] === 1) {
            return null;
        }
        $now = time();
        if (Clock::fromSql($s['created_at']) + self::SESSION_TTL < $now || Clock::fromSql($s['last_seen']) + self::IDLE_TTL < $now) {
            Db::exec('UPDATE grid_sessions SET revoked = 1, access_enc = NULL, refresh_enc = NULL WHERE id = ?', [$sid]);
            return null;
        }
        if (Clock::fromSql($s['access_expires']) - self::REFRESH_SKEW <= $now) {
            $s = self::refreshLocked($sid);
            if ($s === null) {
                return null;
            }
        }
        // Touch last_seen at most once a minute.
        if (Clock::fromSql($s['last_seen']) + 60 < $now) {
            Db::exec('UPDATE grid_sessions SET last_seen = ? WHERE id = ?', [Clock::sql(), $sid]);
        }
        $s['kind'] = 'cookie';
        return $s;
    }

    /**
     * Refresh under a row lock so two concurrent requests don't both spend a
     * rotated refresh token. The second waiter re-reads the row and finds a
     * fresh access token.
     */
    private static function refreshLocked(string $sid): ?array
    {
        return Db::tx(static function () use ($sid): ?array {
            $s = Db::one('SELECT * FROM grid_sessions WHERE id = ? FOR UPDATE', [$sid]);
            if ($s === null || (int)$s['revoked'] === 1) {
                return null;
            }
            if (Clock::fromSql($s['access_expires']) - self::REFRESH_SKEW > time()) {
                return $s; // someone else refreshed while we waited
            }
            if (empty($s['refresh_enc'])) {
                Db::exec('UPDATE grid_sessions SET revoked = 1 WHERE id = ?', [$sid]);
                return null;
            }
            $refresh = Crypto::decrypt($s['refresh_enc']);
            $tok = self::tokenRequest([
                'grant_type' => 'refresh_token',
                'refresh_token' => $refresh,
                'client_id' => Config::get('auth.client_id', 'grid-web'),
            ]);
            if ($tok === null || empty($tok['access_token'])) {
                Db::exec('UPDATE grid_sessions SET revoked = 1, access_enc = NULL, refresh_enc = NULL WHERE id = ?', [$sid]);
                return null;
            }
            $expiresIn = max(30, (int)($tok['expires_in'] ?? 600));
            $newRefresh = isset($tok['refresh_token']) ? Crypto::encrypt((string)$tok['refresh_token']) : $s['refresh_enc'];
            $accessEnc = Crypto::encrypt((string)$tok['access_token']);
            $exp = Clock::sql(time() + $expiresIn);
            Db::exec(
                'UPDATE grid_sessions SET access_enc = ?, refresh_enc = ?, access_expires = ?, last_seen = ? WHERE id = ?',
                [$accessEnc, $newRefresh, $exp, Clock::sql(), $sid]
            );
            $s['access_enc'] = $accessEnc;
            $s['refresh_enc'] = $newRefresh;
            $s['access_expires'] = $exp;
            return $s;
        });
    }

    /**
     * Native clients: validate the bearer token at the IdP's userinfo endpoint and
     * cache the identity for 5 minutes keyed by sha256(token). Nothing decryptable
     * is stored for bearer sessions.
     */
    private static function resolveBearer(string $token): ?array
    {
        $key = hash('sha256', $token);
        $s = Db::one('SELECT * FROM grid_sessions WHERE id = ? AND kind = ?', [$key, 'bearer']);
        $now = time();
        if ($s !== null && (int)$s['revoked'] === 0 && Clock::fromSql($s['access_expires']) > $now) {
            $s['kind'] = 'bearer';
            return $s;
        }
        $user = self::userinfo($token);
        if ($user === null) {
            if ($s !== null) {
                Db::exec('DELETE FROM grid_sessions WHERE id = ?', [$key]);
            }
            return null;
        }
        $exp = Clock::sql($now + self::BEARER_CACHE_TTL);
        Db::exec(
            'INSERT INTO grid_sessions (id, kind, user_id, user_name, user_email, user_avatar, plan, access_enc, refresh_enc, access_expires, csrf, created_at, last_seen, revoked)
             VALUES (?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?, 0)
             ON DUPLICATE KEY UPDATE user_id = VALUES(user_id), user_name = VALUES(user_name), user_email = VALUES(user_email), user_avatar = VALUES(user_avatar), plan = VALUES(plan), access_expires = VALUES(access_expires), last_seen = VALUES(last_seen), revoked = 0',
            [$key, 'bearer', $user['id'], $user['name'], $user['email'], $user['avatar'], $user['plan'], $exp, '', Clock::sql(), Clock::sql()]
        );
        if (mt_rand(1, 20) === 1) {
            Db::exec('DELETE FROM grid_sessions WHERE kind = ? AND access_expires < ?', ['bearer', Clock::sql($now - 86400)]);
        }
        $s = Db::one('SELECT * FROM grid_sessions WHERE id = ?', [$key]);
        if ($s === null) {
            return null;
        }
        $s['kind'] = 'bearer';
        return $s;
    }

    // ------------------------------------------------------------- IdP client

    /** POST JSON to the token endpoint. Returns decoded body or null on failure. */
    private static function tokenRequest(array $fields): ?array
    {
        $res = self::http('POST', (string)Config::get('auth.token_url'), Json::encode($fields), ['Content-Type: application/json', 'Accept: application/json']);
        if ($res === null || $res['status'] < 200 || $res['status'] >= 300) {
            error_log('[grid] token endpoint status ' . ($res['status'] ?? 'n/a'));
            return null;
        }
        try {
            $data = Json::decode($res['body']);
        } catch (\JsonException) {
            return null;
        }
        return is_array($data) ? $data : null;
    }

    /** @return array{id:string,name:string,email:string,avatar:?string,plan:string}|null */
    private static function userinfo(string $accessToken): ?array
    {
        $res = self::http('GET', (string)Config::get('auth.userinfo_url'), null, ['Authorization: Bearer ' . $accessToken, 'Accept: application/json']);
        if ($res === null || $res['status'] !== 200) {
            return null;
        }
        try {
            $u = Json::decode($res['body']);
        } catch (\JsonException) {
            return null;
        }
        if (!is_array($u) || empty($u['sub'])) {
            return null;
        }
        $avatar = isset($u['picture']) && is_string($u['picture']) && preg_match('#^https://#', $u['picture']) ? substr($u['picture'], 0, 500) : null;
        return [
            'id' => substr((string)$u['sub'], 0, 64),
            'name' => substr((string)($u['name'] ?? ''), 0, 120),
            'email' => substr((string)($u['email'] ?? ''), 0, 190),
            'avatar' => $avatar,
            'plan' => (($u['plan'] ?? 'free') === 'pro') ? 'pro' : 'free',
        ];
    }

    private static function revokeAtIdp(string $token, string $hint): void
    {
        $url = (string)Config::get('auth.revoke_url', '');
        if ($url === '') {
            return;
        }
        self::http('POST', $url, Json::encode([
            'token' => $token,
            'token_type_hint' => $hint,
            'client_id' => Config::get('auth.client_id', 'grid-web'),
        ]), ['Content-Type: application/json'], 4);
    }

    /** @return array{status:int,body:string}|null */
    private static function http(string $method, string $url, ?string $body, array $headers, int $timeout = 10): ?array
    {
        if ($url === '' || !preg_match('#^https?://#', $url)) {
            return null;
        }
        $ch = curl_init($url);
        if ($ch === false) {
            return null;
        }
        $opts = [
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_FOLLOWLOCATION => false,
            CURLOPT_CONNECTTIMEOUT => 5,
            CURLOPT_TIMEOUT => $timeout,
            CURLOPT_HTTPHEADER => array_merge($headers, ['User-Agent: Grid/' . GRID_VERSION]),
            CURLOPT_SSL_VERIFYPEER => true,
            CURLOPT_SSL_VERIFYHOST => 2,
            CURLOPT_PROXY => '', // never route IdP calls through an ambient proxy
        ];
        if ($method === 'POST') {
            $opts[CURLOPT_POST] = true;
            $opts[CURLOPT_POSTFIELDS] = $body ?? '';
        }
        curl_setopt_array($ch, $opts);
        $out = curl_exec($ch);
        if ($out === false) {
            error_log('[grid] idp http error: ' . curl_error($ch));
            curl_close($ch);
            return null;
        }
        $status = (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
        curl_close($ch);
        return ['status' => $status, 'body' => (string)$out];
    }
}
