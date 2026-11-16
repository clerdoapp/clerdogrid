<?php
declare(strict_types=1);

namespace Grid;

/**
 * Route table + dispatch. Handlers receive ($params, $body) and return nothing;
 * they emit JSON via Json::send().
 */
final class Router
{
    public const MAX_BODY = 36 * 1024 * 1024;

    /** @var array<int, array{0:string,1:string,2:callable}> */
    private static function routes(): array
    {
        return [
            // Auth (BFF)
            ['GET', '/auth/login', [Auth::class, 'login']],
            ['GET', '/auth/callback', [Auth::class, 'callback']],
            ['POST', '/auth/logout', [Auth::class, 'logout']],
            ['GET', '/api/me', [Auth::class, 'me']],

            // Projects
            ['GET', '/api/projects', [Projects::class, 'list']],
            ['POST', '/api/projects', [Projects::class, 'create']],
            ['GET', '/api/projects/{id}', [Projects::class, 'get']],
            ['PUT', '/api/projects/{id}', [Projects::class, 'save']],
            ['PATCH', '/api/projects/{id}', [Projects::class, 'patch']],
            ['DELETE', '/api/projects/{id}', [Projects::class, 'delete']],
            ['POST', '/api/projects/{id}/duplicate', [Projects::class, 'duplicate']],
            ['GET', '/api/projects/{id}/revisions', [Projects::class, 'revisions']],
            ['POST', '/api/projects/{id}/revisions/{rev}/restore', [Projects::class, 'restore']],

            // Share
            ['POST', '/api/projects/{id}/share', [Share::class, 'create']],
            ['DELETE', '/api/projects/{id}/share', [Share::class, 'revoke']],
            ['GET', '/api/share/{token}', [Share::class, 'fetch']],
            ['GET', '/share/{token}', [Share::class, 'viewer']],

        ];
    }

    public static function dispatch(): void
    {
        Http::securityHeaders();
        $method = Http::method();
        $path = rtrim(Http::path(), '/') ?: '/';

        if ($method === 'OPTIONS') {
            // No CORS: same-origin only.
            http_response_code(204);
            exit;
        }
        if ($method === 'HEAD') {
            $method = 'GET';
        }

        $allowed = [];
        foreach (self::routes() as [$m, $pattern, $handler]) {
            $params = self::match($pattern, $path);
            if ($params === null) {
                continue;
            }
            if ($m !== $method) {
                $allowed[] = $m;
                continue;
            }
            Projects::ensureSchema();
            $body = in_array($method, ['POST', 'PUT', 'PATCH', 'DELETE'], true) ? self::readBody() : [];
            $handler($params, $body);
            return;
        }
        if ($allowed !== []) {
            header('Allow: ' . implode(', ', array_unique($allowed)));
            Json::error(405, 'method_not_allowed', 'Method not allowed');
        }
        Json::error(404, 'not_found', 'No such endpoint');
    }

    /** @return array<string,string>|null */
    private static function match(string $pattern, string $path): ?array
    {
        $regex = '#^' . preg_replace('#\{(\w+)\}#', '(?P<$1>[^/]+)', str_replace('.', '\.', $pattern)) . '$#';
        if (!preg_match($regex, $path, $m)) {
            return null;
        }
        $out = [];
        foreach ($m as $k => $v) {
            if (is_string($k)) {
                $out[$k] = $v;
            }
        }
        return $out;
    }

    /** JSON body as an associative array (empty array when no body). */
    private static function readBody(): array
    {
        $len = (int)($_SERVER['CONTENT_LENGTH'] ?? 0);
        if ($len > self::MAX_BODY) {
            throw new HttpError(413, 'too_large', 'Request body too large');
        }
        $raw = file_get_contents('php://input', false, null, 0, self::MAX_BODY + 1);
        if ($raw === false || $raw === '') {
            return [];
        }
        if (strlen($raw) > self::MAX_BODY) {
            throw new HttpError(413, 'too_large', 'Request body too large');
        }
        $ct = strtolower((string)($_SERVER['CONTENT_TYPE'] ?? ''));
        if ($ct !== '' && !str_contains($ct, 'application/json') && !str_contains($ct, 'text/plain')) {
            throw new HttpError(415, 'validation', 'Expected application/json');
        }
        try {
            $data = Json::decode($raw, 128);
        } catch (\JsonException) {
            throw new HttpError(422, 'validation', 'Malformed JSON body');
        }
        if (!is_array($data)) {
            throw new HttpError(422, 'validation', 'Body must be a JSON object');
        }
        $data['__bytes'] = strlen($raw);
        return $data;
    }
}
