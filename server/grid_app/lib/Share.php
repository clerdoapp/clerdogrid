<?php
declare(strict_types=1);

namespace Grid;

/** Public read-only share links. */
final class Share
{
    public static function create(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = Projects::ownRow($s, $p['id']);
        $mode = $body['mode'] ?? 'view';
        if (!in_array($mode, ['view', 'comment'], true)) {
            throw new HttpError(422, 'validation', 'mode must be view|comment');
        }
        $existing = Db::one('SELECT token FROM grid_shares WHERE project_id = ?', [$row['id']]);
        $token = Id::token(32); // 43 chars
        if ($existing !== null) {
            Db::exec('UPDATE grid_shares SET token = ?, mode = ?, created_at = ?, revoked = 0 WHERE project_id = ?', [$token, $mode, Clock::sql(), $row['id']]);
        } else {
            Db::exec('INSERT INTO grid_shares (token, project_id, mode, created_at, revoked) VALUES (?, ?, ?, ?, 0)', [$token, $row['id'], $mode, Clock::sql()]);
        }
        Json::send(['token' => $token, 'url' => Http::origin() . '/?s=' . $token], 201);
    }

    public static function revoke(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = Projects::ownRow($s, $p['id']);
        Db::exec('DELETE FROM grid_shares WHERE project_id = ?', [$row['id']]);
        Json::send(null, 204);
    }

    public static function fetch(array $p, array $body): void
    {
        RateLimit::hit('share:' . Http::clientIp(), (int)Config::get('limits.share_per_minute', 60));
        $token = $p['token'];
        if (!preg_match('/^[A-Za-z0-9_-]{43}$/', $token)) {
            throw new HttpError(404, 'not_found', 'Share link not found');
        }
        $sh = Db::one('SELECT * FROM grid_shares WHERE token = ? AND revoked = 0', [$token]);
        if ($sh === null) {
            throw new HttpError(404, 'not_found', 'Share link not found');
        }
        $row = Db::one('SELECT name, document, format FROM grid_projects WHERE id = ?', [$sh['project_id']]);
        if ($row === null) {
            throw new HttpError(404, 'not_found', 'Share link not found');
        }
        Json::send(['workbook' => Projects::workbookOf($row), 'name' => $row['name'], 'mode' => $sh['mode']]);
    }

    /** /share/:token links from the first Grid web app open the app in share mode. */
    public static function viewer(array $p, array $body): void
    {
        $token = $p['token'];
        if (!preg_match('/^[A-Za-z0-9_-]{43}$/', $token)) {
            throw new HttpError(404, 'not_found', 'Share link not found');
        }
        Http::redirect('/?s=' . rawurlencode($token));
    }
}
