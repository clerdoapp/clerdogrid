<?php
declare(strict_types=1);

namespace Grid;

/**
 * Cloud workbooks: schema, CRUD and revisions.
 *
 * A workbook is stored as its .xlsx bytes (base64 in `document`, `format` = 'xlsx'). Rows written by the
 * first Grid web app hold its JSON document (`format` = 'json'); they are converted to .xlsx on read
 * with Xlsx::build and become .xlsx on their next save.
 */
final class Projects
{
    public const KEEP_REVISIONS = 5;
    public const MAX_SHEETS = 64;
    public const MAX_THUMB = 200 * 1024;

    private static bool $schemaChecked = false;

    // ----------------------------------------------------------------- schema

    /** Idempotent: creates tables and ALTER-adds columns missing after an upgrade. */
    public static function ensureSchema(): void
    {
        if (self::$schemaChecked) {
            return;
        }
        self::$schemaChecked = true;
        foreach (self::schema() as $name => [$create, $columns]) {
            if (!Db::tableExists($name)) {
                try {
                    Db::exec($create);
                } catch (\PDOException $e) {
                    if (!Db::tableExists($name)) {
                        throw $e;
                    }
                }
                continue;
            }
            $have = array_map('strtolower', Db::columns($name));
            foreach ($columns as $col => $ddl) {
                if (!in_array(strtolower($col), $have, true)) {
                    try {
                        Db::exec("ALTER TABLE {$name} ADD COLUMN {$col} {$ddl}");
                    } catch (\PDOException $e) {
                        error_log('[grid] alter failed ' . $name . '.' . $col . ': ' . $e->getMessage());
                    }
                }
            }
        }
        if (Db::isSqlite()) {
            foreach ([
                'CREATE INDEX IF NOT EXISTS ix_grid_projects_user ON grid_projects(user_id, archived, updated_at)',
                'CREATE INDEX IF NOT EXISTS ix_grid_revisions_project ON grid_project_revisions(project_id, revision)',
                'CREATE INDEX IF NOT EXISTS ix_grid_sessions_user ON grid_sessions(user_id)',
            ] as $sql) {
                Db::exec($sql);
            }
        }
    }

    /** @return array<string, array{0:string,1:array<string,string>}> table => [CREATE, columns for ALTER-add] */
    public static function schema(): array
    {
        $tail = ' ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci';
        return [
            'grid_sessions' => [
                "CREATE TABLE grid_sessions (
                    id CHAR(64) NOT NULL PRIMARY KEY,
                    kind VARCHAR(10) NOT NULL DEFAULT 'cookie',
                    user_id VARCHAR(64) NOT NULL,
                    user_name VARCHAR(120) NOT NULL DEFAULT '',
                    user_email VARCHAR(190) NOT NULL DEFAULT '',
                    user_avatar VARCHAR(500) NULL,
                    plan VARCHAR(10) NOT NULL DEFAULT 'free',
                    access_enc BLOB NULL,
                    refresh_enc BLOB NULL,
                    access_expires DATETIME NOT NULL,
                    csrf CHAR(64) NOT NULL,
                    created_at DATETIME NOT NULL,
                    last_seen DATETIME NOT NULL,
                    revoked TINYINT(1) NOT NULL DEFAULT 0,
                    INDEX ix_grid_sessions_user (user_id)
                )" . $tail,
                [
                    'kind' => "VARCHAR(10) NOT NULL DEFAULT 'cookie'",
                    'user_avatar' => 'VARCHAR(500) NULL',
                    'plan' => "VARCHAR(10) NOT NULL DEFAULT 'free'",
                    'revoked' => 'TINYINT(1) NOT NULL DEFAULT 0',
                ],
            ],
            'grid_auth_states' => [
                "CREATE TABLE grid_auth_states (
                    state VARCHAR(128) NOT NULL PRIMARY KEY,
                    verifier VARCHAR(128) NOT NULL,
                    next_path VARCHAR(512) NOT NULL DEFAULT '/',
                    created_at DATETIME NOT NULL,
                    expires_at DATETIME NOT NULL
                )" . $tail,
                ['next_path' => "VARCHAR(512) NOT NULL DEFAULT '/'"],
            ],
            'grid_projects' => [
                "CREATE TABLE grid_projects (
                    id CHAR(36) NOT NULL PRIMARY KEY,
                    user_id VARCHAR(64) NOT NULL,
                    name VARCHAR(200) NOT NULL,
                    document LONGTEXT NOT NULL,
                    format VARCHAR(8) NOT NULL DEFAULT 'xlsx',
                    revision INT NOT NULL DEFAULT 1,
                    bytes INT NOT NULL DEFAULT 0,
                    sheets INT NOT NULL DEFAULT 1,
                    thumbnail MEDIUMTEXT NULL,
                    starred TINYINT(1) NOT NULL DEFAULT 0,
                    archived TINYINT(1) NOT NULL DEFAULT 0,
                    folder VARCHAR(100) NOT NULL DEFAULT '',
                    color VARCHAR(9) NOT NULL DEFAULT '',
                    tags TEXT NULL,
                    notes TEXT NULL,
                    created_at DATETIME NOT NULL,
                    updated_at DATETIME NOT NULL,
                    opened_at DATETIME NULL,
                    locker_id CHAR(36) NULL,
                    locker_rev INT NULL,
                    locker_pending TINYINT(1) NOT NULL DEFAULT 0,
                    INDEX ix_grid_projects_user (user_id, archived, updated_at)
                )" . $tail,
                [
                    // Added to a table created by the first Grid web app: its rows are JSON documents.
                    'format' => "VARCHAR(8) NOT NULL DEFAULT 'json'",
                    'sheets' => 'INT NOT NULL DEFAULT 1',
                    'thumbnail' => 'MEDIUMTEXT NULL',
                    'starred' => 'TINYINT(1) NOT NULL DEFAULT 0',
                    'archived' => 'TINYINT(1) NOT NULL DEFAULT 0',
                    'folder' => "VARCHAR(100) NOT NULL DEFAULT ''",
                    'color' => "VARCHAR(9) NOT NULL DEFAULT ''",
                    'tags' => 'TEXT NULL',
                    'notes' => 'TEXT NULL',
                    'opened_at' => 'DATETIME NULL',
                    'locker_id' => 'CHAR(36) NULL',
                    'locker_rev' => 'INT NULL',
                    'locker_pending' => 'TINYINT(1) NOT NULL DEFAULT 0',
                ],
            ],
            'grid_project_revisions' => [
                "CREATE TABLE grid_project_revisions (
                    id BIGINT AUTO_INCREMENT PRIMARY KEY,
                    project_id CHAR(36) NOT NULL,
                    revision INT NOT NULL,
                    document LONGTEXT NOT NULL,
                    format VARCHAR(8) NOT NULL DEFAULT 'xlsx',
                    bytes INT NOT NULL DEFAULT 0,
                    saved_at DATETIME NOT NULL,
                    locker_rev INT NULL,
                    INDEX ix_grid_revisions_project (project_id, revision)
                )" . $tail,
                [
                    'format' => "VARCHAR(8) NOT NULL DEFAULT 'json'",
                    'bytes' => 'INT NOT NULL DEFAULT 0',
                    'locker_rev' => 'INT NULL',
                ],
            ],
            'grid_shares' => [
                "CREATE TABLE grid_shares (
                    token CHAR(43) NOT NULL PRIMARY KEY,
                    project_id CHAR(36) NOT NULL,
                    mode ENUM('view','comment') NOT NULL DEFAULT 'view',
                    created_at DATETIME NOT NULL,
                    revoked TINYINT(1) NOT NULL DEFAULT 0,
                    UNIQUE INDEX ux_grid_shares_project (project_id)
                )" . $tail,
                ['revoked' => 'TINYINT(1) NOT NULL DEFAULT 0'],
            ],
            'grid_ratelimit' => [
                "CREATE TABLE grid_ratelimit (
                    rl_key VARCHAR(120) NOT NULL,
                    bucket INT NOT NULL,
                    hits INT NOT NULL DEFAULT 0,
                    PRIMARY KEY (rl_key, bucket)
                )" . $tail,
                [],
            ],
        ];
    }

    // ------------------------------------------------------------------ routes

    public static function list(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $archived = (int)(($_GET['archived'] ?? '0') === '1');
        $folder = isset($_GET['folder']) ? substr((string)$_GET['folder'], 0, 100) : null;
        $sql = 'SELECT ' . self::metaCols() . ' FROM grid_projects WHERE user_id = ? AND archived = ?';
        $params = [$s['user_id'], $archived];
        if ($folder !== null && $folder !== '') {
            $sql .= ' AND folder = ?';
            $params[] = $folder;
        }
        $sql .= ' ORDER BY starred DESC, updated_at DESC LIMIT 1000';
        Json::send(['projects' => array_map([self::class, 'meta'], Db::all($sql, $params))]);
    }

    public static function create(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $plan = Auth::plan($s);
        if ($plan !== 'pro') {
            $cnt = Db::one('SELECT COUNT(*) AS c FROM grid_projects WHERE user_id = ?', [$s['user_id']]);
            $max = (int)Config::get('limits.free_projects', 5);
            if ((int)($cnt['c'] ?? 0) >= $max) {
                throw new HttpError(429, 'limit', 'Your plan allows ' . $max . ' cloud workbooks.');
            }
        }
        $name = self::cleanName($body['name'] ?? null, 'Book1');
        $b64 = self::validateWorkbook($body['workbook'] ?? null, Auth::docLimit($plan));
        $sheets = self::cleanSheets($body['sheets'] ?? null);
        $thumb = self::validateThumbnail($body['thumbnail'] ?? null);
        $id = Id::uuid();
        $now = Clock::sql();
        Db::exec(
            'INSERT INTO grid_projects (id, user_id, name, document, format, revision, bytes, sheets, thumbnail, starred, archived, folder, color, tags, notes, created_at, updated_at, opened_at)
             VALUES (?, ?, ?, ?, \'xlsx\', 1, ?, ?, ?, 0, 0, \'\', \'\', \'[]\', \'\', ?, ?, ?)',
            [$id, $s['user_id'], $name, $b64, self::decodedSize($b64), $sheets, $thumb, $now, $now, $now]
        );
        $row = self::ownRow($s, $id);
        self::setLocker($id, self::pushToLocker($s, $row, base64_decode($b64), $name));
        Json::send(['project' => self::meta(self::ownRow($s, $id))], 201);
    }

    public static function get(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::syncFromLocker($s, self::ownRow($s, $p['id']));
        Db::exec('UPDATE grid_projects SET opened_at = ? WHERE id = ?', [Clock::sql(), $row['id']]);
        Json::send(['project' => self::meta($row), 'workbook' => self::workbookOf($row)]);
    }

    public static function save(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $id = self::cleanId($p['id']);
        if (!isset($body['baseRevision']) || !is_int($body['baseRevision'])) {
            throw new HttpError(422, 'validation', 'baseRevision (integer) required');
        }
        $b64 = self::validateWorkbook($body['workbook'] ?? null, Auth::docLimit(Auth::plan($s)));
        $sheets = self::cleanSheets($body['sheets'] ?? null);
        $hasThumb = array_key_exists('thumbnail', $body);
        $thumb = self::validateThumbnail($body['thumbnail'] ?? null);
        $hasName = array_key_exists('name', $body);
        $base = $body['baseRevision'];

        $result = Db::tx(static function () use ($s, $id, $b64, $sheets, $thumb, $hasThumb, $hasName, $body, $base): array {
            $row = Db::one('SELECT * FROM grid_projects WHERE id = ? AND user_id = ? FOR UPDATE', [$id, $s['user_id']]);
            if ($row === null) {
                throw new HttpError(404, 'not_found', 'Workbook not found');
            }
            if ((int)$row['revision'] !== $base) {
                return ['conflict' => $row];
            }
            $name = $hasName ? self::cleanName($body['name'], $row['name']) : $row['name'];
            $locker = self::pushToLocker($s, $row, base64_decode($b64), $name);
            if (isset($locker['conflict'])) {
                return ['conflict' => self::adoptFromLocker($s, $row)];
            }
            self::keepRevision($row, ($locker['locker_pending'] ?? 1) === 0);
            Db::exec(
                'UPDATE grid_projects SET document = ?, format = \'xlsx\', revision = ?, bytes = ?, sheets = ?, name = ?, updated_at = ?'
                . ($hasThumb ? ', thumbnail = ?' : '') . ' WHERE id = ?',
                array_merge(
                    [$b64, (int)$row['revision'] + 1, self::decodedSize($b64), $sheets, $name, Clock::sql()],
                    $hasThumb ? [$thumb] : [],
                    [$id]
                )
            );
            self::setLocker($id, $locker);
            if ($hasName && $name !== $row['name'] && !empty($row['locker_id']) && !isset($locker['locker_id'])) {
                self::renameInLocker($s, (string)$row['locker_id'], $name);
            }
            return ['saved' => true];
        });

        if (isset($result['conflict'])) {
            $row = $result['conflict'];
            Json::send([
                'error' => 'conflict',
                'message' => 'The workbook was changed somewhere else',
                'project' => self::meta($row),
                'workbook' => self::workbookOf($row),
            ], 409);
        }
        Json::send(['project' => self::meta(self::ownRow($s, $id))]);
    }

    public static function patch(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::ownRow($s, $p['id']);
        $sets = [];
        $vals = [];
        if (array_key_exists('name', $body)) {
            $sets[] = 'name = ?';
            $vals[] = self::cleanName($body['name'], $row['name']);
        }
        foreach (['starred', 'archived'] as $flag) {
            if (array_key_exists($flag, $body)) {
                if (!is_bool($body[$flag])) {
                    throw new HttpError(422, 'validation', $flag . ' must be a boolean');
                }
                $sets[] = $flag . ' = ?';
                $vals[] = $body[$flag] ? 1 : 0;
            }
        }
        if (array_key_exists('folder', $body)) {
            if (!is_string($body['folder'])) {
                throw new HttpError(422, 'validation', 'folder must be a string');
            }
            $sets[] = 'folder = ?';
            $vals[] = mb_substr(trim(self::stripControl($body['folder'])), 0, 100);
        }
        if ($sets !== []) {
            $vals[] = $row['id'];
            Db::exec('UPDATE grid_projects SET ' . implode(', ', $sets) . ' WHERE id = ?', $vals);
        }
        $fresh = self::ownRow($s, $row['id']);
        if ($fresh['name'] !== $row['name'] && !empty($row['locker_id'])) {
            self::renameInLocker($s, (string)$row['locker_id'], $fresh['name']);
        }
        Json::send(['project' => self::meta($fresh)]);
    }

    public static function duplicate(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::ownRow($s, $p['id']);
        if (Auth::plan($s) !== 'pro') {
            $cnt = Db::one('SELECT COUNT(*) AS c FROM grid_projects WHERE user_id = ?', [$s['user_id']]);
            if ((int)($cnt['c'] ?? 0) >= (int)Config::get('limits.free_projects', 5)) {
                throw new HttpError(429, 'limit', 'Cloud workbook limit reached');
            }
        }
        $name = self::cleanName($body['name'] ?? null, mb_substr($row['name'] . ' (copy)', 0, 200));
        $id = Id::uuid();
        $now = Clock::sql();
        Db::exec(
            'INSERT INTO grid_projects (id, user_id, name, document, format, revision, bytes, sheets, thumbnail, starred, archived, folder, color, tags, notes, created_at, updated_at, opened_at)
             VALUES (?, ?, ?, ?, ?, 1, ?, ?, ?, 0, 0, ?, ?, ?, ?, ?, ?, NULL)',
            [$id, $s['user_id'], $name, $row['document'], $row['format'], (int)$row['bytes'], (int)$row['sheets'], $row['thumbnail'], $row['folder'], $row['color'], $row['tags'] ?? '[]', $row['notes'] ?? '', $now, $now]
        );
        Json::send(['project' => self::meta(self::ownRow($s, $id))], 201);
    }

    public static function delete(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::ownRow($s, $p['id']);
        if (!empty($row['locker_id']) && Locker::enabled()) {
            try {
                Locker::trash($s, (string)$row['locker_id']);
            } catch (HttpError $e) {
                if (!self::isLockerTrouble($e)) {
                    throw $e;
                }
            }
        }
        Db::tx(static function () use ($row): void {
            Db::exec('DELETE FROM grid_project_revisions WHERE project_id = ?', [$row['id']]);
            Db::exec('DELETE FROM grid_shares WHERE project_id = ?', [$row['id']]);
            Db::exec('DELETE FROM grid_projects WHERE id = ?', [$row['id']]);
        });
        Json::send(null, 204);
    }

    public static function revisions(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::ownRow($s, $p['id']);
        $revs = Db::all(
            'SELECT revision, bytes, saved_at FROM grid_project_revisions WHERE project_id = ? ORDER BY revision DESC LIMIT ' . self::KEEP_REVISIONS,
            [$row['id']]
        );
        Json::send([
            'current' => ['revision' => (int)$row['revision'], 'savedAt' => Clock::iso($row['updated_at']), 'bytes' => (int)$row['bytes']],
            'revisions' => array_map(static fn(array $r): array => [
                'revision' => (int)$r['revision'],
                'savedAt' => Clock::iso($r['saved_at']),
                'bytes' => (int)$r['bytes'],
            ], $revs),
        ]);
    }

    public static function restore(array $p, array $body): void
    {
        $s = Auth::requireUser();
        $row = self::ownRow($s, $p['id']);
        if (!preg_match('/^\d{1,9}$/', $p['rev'])) {
            throw new HttpError(404, 'not_found', 'Version not found');
        }
        $rev = (int)$p['rev'];
        $id = $row['id'];
        Db::tx(static function () use ($s, $id, $rev): void {
            $cur = Db::one('SELECT * FROM grid_projects WHERE id = ? FOR UPDATE', [$id]);
            $old = Db::one('SELECT * FROM grid_project_revisions WHERE project_id = ? AND revision = ? ORDER BY id DESC LIMIT 1', [$id, $rev]);
            if ($cur === null || $old === null) {
                throw new HttpError(404, 'not_found', 'Version not found');
            }
            $b64 = self::revisionWorkbook($s, $cur, $old);
            $locker = self::pushToLocker($s, $cur, base64_decode($b64), $cur['name']);
            if (isset($locker['conflict'])) {
                self::adoptFromLocker($s, $cur);
                throw new HttpError(409, 'conflict', 'The workbook was changed somewhere else. Try again.');
            }
            self::keepRevision($cur, ($locker['locker_pending'] ?? 1) === 0);
            Db::exec(
                'UPDATE grid_projects SET document = ?, format = \'xlsx\', revision = ?, bytes = ?, updated_at = ? WHERE id = ?',
                [$b64, (int)$cur['revision'] + 1, self::decodedSize($b64), Clock::sql(), $id]
            );
            self::setLocker($id, $locker);
        });
        $row = self::ownRow($s, $id);
        Json::send(['project' => self::meta($row), 'workbook' => self::workbookOf($row)]);
    }

    // ---------------------------------------------------------------- helpers

    /**
     * Files the current state of a row as a revision and trims the history to KEEP_REVISIONS.
     * When Locker now keeps that content as a version, only the Locker revision is stored.
     */
    private static function keepRevision(array $row, bool $keptByLocker = false): void
    {
        $inLocker = $keptByLocker && !empty($row['locker_id']) && (int)($row['locker_pending'] ?? 0) === 0 && $row['locker_rev'] !== null;
        Db::exec(
            'INSERT INTO grid_project_revisions (project_id, revision, document, format, bytes, saved_at, locker_rev) VALUES (?, ?, ?, ?, ?, ?, ?)',
            [$row['id'], (int)$row['revision'], $inLocker ? '' : $row['document'], $row['format'] ?? 'json', (int)$row['bytes'], $row['updated_at'], $inLocker ? (int)$row['locker_rev'] : null]
        );
        $stale = Db::all(
            'SELECT id FROM grid_project_revisions WHERE project_id = ? ORDER BY revision DESC, id DESC LIMIT 100 OFFSET ' . self::KEEP_REVISIONS,
            [$row['id']]
        );
        foreach ($stale as $o) {
            Db::exec('DELETE FROM grid_project_revisions WHERE id = ?', [(int)$o['id']]);
        }
    }

    // ---------------------------------------------------------------- Locker

    /**
     * Puts the workbook in the user's Locker. Returns the Locker columns to store, or ['conflict' => true]
     * when the file was changed in Locker first. Locker trouble never loses a save: the row is marked
     * pending and sent with a later save.
     */
    private static function pushToLocker(array $s, array $row, string $bytes, string $name): array
    {
        if (!Locker::enabled()) {
            return [];
        }
        try {
            if (!empty($row['locker_id'])) {
                $r = Locker::replace($s, (string)$row['locker_id'], (int)$row['locker_rev'], $bytes);
                if ($r === Locker::CONFLICT) {
                    return ['conflict' => true];
                }
                if ($r !== Locker::GONE) {
                    return ['locker_rev' => $r['rev'], 'locker_pending' => 0];
                }
            }
            $f = Locker::create($s, $name, $bytes);
            return ['locker_id' => $f['id'], 'locker_rev' => $f['rev'], 'locker_pending' => 0];
        } catch (HttpError $e) {
            if (!self::isLockerTrouble($e)) {
                throw $e;
            }
            error_log('[grid] locker: save kept here and sent later (' . $e->getMessage() . ')');
            return ['locker_pending' => 1];
        }
    }

    /** Takes the file's current content from Locker after it was changed there; Grid's copy becomes a revision. */
    private static function adoptFromLocker(array $s, array $row): array
    {
        $f = Locker::fetch($s, (string)$row['locker_id']);
        $b64 = base64_encode($f['bytes']);
        self::keepRevision($row);
        Db::exec(
            'UPDATE grid_projects SET document = ?, format = \'xlsx\', revision = ?, bytes = ?, locker_rev = ?, locker_pending = 0, updated_at = ? WHERE id = ?',
            [$b64, (int)$row['revision'] + 1, strlen($f['bytes']), $f['rev'], Clock::sql(), $row['id']]
        );
        return Db::one('SELECT * FROM grid_projects WHERE id = ?', [$row['id']]) ?? $row;
    }

    /**
     * On open: moves a workbook from before Locker into it, sends a pending save, or picks up a change
     * made in Locker. If Locker can't be reached, Grid's own copy opens.
     */
    private static function syncFromLocker(array $s, array $row): array
    {
        if (!Locker::enabled()) {
            return $row;
        }
        try {
            if (empty($row['locker_id']) || (int)$row['locker_pending'] === 1) {
                $locker = self::pushToLocker($s, $row, base64_decode(self::workbookOf($row)), $row['name']);
                if (!isset($locker['conflict'])) {
                    self::setLocker($row['id'], $locker);
                }
                return self::ownRow($s, $row['id']);
            }
            $rev = Locker::rev($s, (string)$row['locker_id']);
            if ($rev === null) {
                Db::exec('UPDATE grid_projects SET locker_id = NULL, locker_rev = NULL, locker_pending = 1 WHERE id = ?', [$row['id']]);
                return self::ownRow($s, $row['id']);
            }
            if ($rev !== (int)$row['locker_rev']) {
                return Db::tx(static function () use ($s, $row): array {
                    $cur = Db::one('SELECT * FROM grid_projects WHERE id = ? FOR UPDATE', [$row['id']]) ?? $row;
                    return self::adoptFromLocker($s, $cur);
                });
            }
        } catch (HttpError $e) {
            if (!self::isLockerTrouble($e)) {
                throw $e;
            }
        }
        return $row;
    }

    /** A saved revision's workbook: Grid's own copy, or the version Locker keeps. */
    private static function revisionWorkbook(array $s, array $cur, array $old): string
    {
        if ((string)$old['document'] !== '') {
            return self::workbookOf($old);
        }
        $bytes = !empty($cur['locker_id']) && $old['locker_rev'] !== null
            ? Locker::version($s, (string)$cur['locker_id'], (int)$old['locker_rev'])
            : null;
        if ($bytes === null) {
            throw new HttpError(404, 'not_found', 'That version is no longer kept');
        }
        return base64_encode($bytes);
    }

    private static function renameInLocker(array $s, string $lockerId, string $name): void
    {
        try {
            Locker::rename($s, $lockerId, $name);
        } catch (HttpError $e) {
            if (!self::isLockerTrouble($e)) {
                throw $e;
            }
        }
    }

    private static function setLocker(string $id, array $cols): void
    {
        $cols = array_intersect_key($cols, ['locker_id' => 1, 'locker_rev' => 1, 'locker_pending' => 1]);
        if ($cols === []) {
            return;
        }
        Db::exec(
            'UPDATE grid_projects SET ' . implode(', ', array_map(static fn(string $c): string => $c . ' = ?', array_keys($cols))) . ' WHERE id = ?',
            array_merge(array_values($cols), [$id])
        );
    }

    private static function isLockerTrouble(HttpError $e): bool
    {
        return in_array($e->errorCode, ['storage', 'quota'], true);
    }

    /** The row's workbook as base64 .xlsx, converting a JSON document from the first Grid web app. */
    public static function workbookOf(array $row): string
    {
        if (($row['format'] ?? 'json') === 'xlsx') {
            return (string)$row['document'];
        }
        try {
            $doc = Json::decode((string)$row['document']);
        } catch (\JsonException) {
            throw new HttpError(500, 'server', 'This workbook could not be read');
        }
        if (!is_array($doc) || !is_array($doc['sheets'] ?? null)) {
            throw new HttpError(500, 'server', 'This workbook could not be read');
        }
        $tmp = Xlsx::build($doc);
        try {
            $bytes = file_get_contents($tmp);
        } finally {
            @unlink($tmp);
        }
        if ($bytes === false) {
            throw new HttpError(500, 'server', 'This workbook could not be converted');
        }
        return base64_encode($bytes);
    }

    /**
     * Checks a base64 .xlsx: strict base64, size limit, and a zip package with the parts every workbook has.
     * Returns the canonical base64 to store.
     */
    private static function validateWorkbook(mixed $b64, int $byteLimit): string
    {
        if (!is_string($b64) || $b64 === '') {
            throw new HttpError(422, 'validation', 'workbook (base64 .xlsx) required');
        }
        if (strlen($b64) > intdiv($byteLimit * 4, 3) + 4) {
            throw new HttpError(413, 'too_large', 'Workbook exceeds ' . intdiv($byteLimit, 1024 * 1024) . ' MB');
        }
        $bytes = base64_decode($b64, true);
        if ($bytes === false || strlen($bytes) < 22) {
            throw new HttpError(422, 'validation', 'workbook is not valid base64');
        }
        if (strlen($bytes) > $byteLimit) {
            throw new HttpError(413, 'too_large', 'Workbook exceeds ' . intdiv($byteLimit, 1024 * 1024) . ' MB');
        }
        if (strncmp($bytes, "PK\x03\x04", 4) !== 0) {
            throw new HttpError(422, 'validation', 'workbook is not an .xlsx package');
        }
        if (class_exists(\ZipArchive::class)) {
            $tmp = tempnam(sys_get_temp_dir(), 'gridv');
            try {
                file_put_contents($tmp, $bytes);
                $zip = new \ZipArchive();
                if ($zip->open($tmp, \ZipArchive::RDONLY) !== true) {
                    throw new HttpError(422, 'validation', 'workbook is not an .xlsx package');
                }
                $ok = $zip->locateName('[Content_Types].xml') !== false && $zip->locateName('xl/workbook.xml') !== false;
                $zip->close();
                if (!$ok) {
                    throw new HttpError(422, 'validation', 'workbook is not an .xlsx package');
                }
            } finally {
                @unlink($tmp);
            }
        }
        return base64_encode($bytes);
    }

    private static function decodedSize(string $b64): int
    {
        return intdiv(strlen($b64) * 3, 4) - substr_count(substr($b64, -2), '=');
    }

    private static function cleanSheets(mixed $n): int
    {
        return is_int($n) && $n >= 1 && $n <= self::MAX_SHEETS ? $n : 1;
    }

    private static function metaCols(): string
    {
        return 'id, user_id, name, revision, bytes, sheets, thumbnail, starred, archived, folder, color, tags, notes, created_at, updated_at, opened_at, locker_id, locker_pending';
    }

    public static function meta(array $r): array
    {
        return [
            'id' => $r['id'],
            'name' => $r['name'],
            'revision' => (int)$r['revision'],
            'createdAt' => Clock::iso($r['created_at']),
            'updatedAt' => Clock::iso($r['updated_at']),
            'openedAt' => Clock::iso($r['opened_at'] ?? null),
            'starred' => (int)$r['starred'] === 1,
            'archived' => (int)$r['archived'] === 1,
            'folder' => (string)($r['folder'] ?? ''),
            'bytes' => (int)$r['bytes'],
            'sheets' => (int)$r['sheets'],
            'thumbnail' => ($r['thumbnail'] ?? null) === null || $r['thumbnail'] === '' ? null : $r['thumbnail'],
            'lockerId' => empty($r['locker_id']) || (int)($r['locker_pending'] ?? 0) === 1 ? null : $r['locker_id'],
            'storage' => !Locker::enabled() ? 'grid' : (empty($r['locker_id']) || (int)($r['locker_pending'] ?? 0) === 1 ? 'pending' : 'locker'),
        ];
    }

    public static function cleanId(string $id): string
    {
        if (!preg_match('/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/', $id)) {
            throw new HttpError(404, 'not_found', 'Workbook not found');
        }
        return $id;
    }

    /** Full row owned by the session user, or 404 (never 403). */
    public static function ownRow(array $s, string $id): array
    {
        $id = self::cleanId($id);
        $row = Db::one('SELECT * FROM grid_projects WHERE id = ? AND user_id = ?', [$id, $s['user_id']]);
        if ($row === null) {
            throw new HttpError(404, 'not_found', 'Workbook not found');
        }
        return $row;
    }

    private static function cleanName(mixed $name, string $fallback): string
    {
        if (!is_string($name)) {
            return $fallback;
        }
        $n = trim(self::stripControl($name));
        return $n === '' ? $fallback : mb_substr($n, 0, 200);
    }

    private static function stripControl(string $s): string
    {
        return preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/u', '', $s) ?? '';
    }

    private static function validateThumbnail(mixed $thumb): ?string
    {
        if ($thumb === null || $thumb === '') {
            return null;
        }
        if (!is_string($thumb) || strlen($thumb) > self::MAX_THUMB || !preg_match('#^data:image/(jpeg|png);base64,[A-Za-z0-9+/]+={0,2}$#', $thumb)) {
            throw new HttpError(422, 'validation', 'thumbnail must be a JPEG or PNG data URL of at most 200 KB');
        }
        return $thumb;
    }
}
