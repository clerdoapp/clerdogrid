# Cloud backend for grid.clerdo.app

The desktop app and the offline web app (GitHub Pages) need no server. This folder is only for the
build served at **grid.clerdo.app**, which adds Clerdo Account sign-in, cloud workbooks, autosave,
version history and share links (`src/features/cloud`, built with `npm run build:cloud`).

PHP 8.3, no Composer, for DirectAdmin shared hosting (LiteSpeed + MariaDB). Tables live in Clerdo Core's
database with the `grid_` prefix.

```
server/
  public_html/          web root: .htaccess (routing, CSP, deny *.php except api.php), api.php
  grid_app/             OUTSIDE the web root
    bootstrap.php       strict errors → JSON, autoloader, config
    config.sample.php   copy to config.php (secrets, 0600, never committed)
    router.php          routes
    lib/Core.php        Db (MariaDB, sqlite in dev), Config, Json, Crypto (AES-256-GCM), Http, Id, Clock, RateLimit
    lib/Auth.php        OAuth 2.0 + PKCE against auth.clerdo.app; the browser only gets an HttpOnly cookie
    lib/Projects.php    cloud workbooks: schema, CRUD, 409 conflicts, 5 earlier versions
    lib/Share.php       read-only share links (rate-limited)
    lib/Xlsx.php        converts workbooks saved by the first Grid web app (JSON) to .xlsx
  dev/                  mock IdP, dev router, start.sh, api-test.sh, deployer template, package.sh
```

## API

| | |
|---|---|
| `GET /auth/login?next=` · `GET /auth/callback` · `POST /auth/logout` | sign-in flow |
| `GET /api/me` | user, CSRF token, limits |
| `GET/POST /api/projects` | list · create `{name, workbook, sheets}` |
| `GET/PUT/PATCH/DELETE /api/projects/:id` | open · save `{baseRevision, workbook, sheets, name}` (409 with the other version) · rename/star · delete |
| `POST /api/projects/:id/duplicate` | copy |
| `GET /api/projects/:id/revisions` · `POST …/revisions/:rev/restore` | version history |
| `POST/DELETE /api/projects/:id/share` · `GET /api/share/:token` | share link |

`workbook` is the .xlsx file as standard base64. The server checks it is a zip package with
`[Content_Types].xml` and `xl/workbook.xml`, within the plan's size limit. Rows written by the first Grid
web app (`format = 'json'`) are converted to .xlsx when opened and become .xlsx on their next save;
`ensureSchema()` adds the `format` column to an existing table on the first request.

Security: CSRF token plus same-origin `Origin` on every write, tokens encrypted at rest, refresh under a
row lock, cross-user access is always 404, prepared statements only, per-IP rate limits on sign-in and
share links. CSP `frame-ancestors` allows kelvo.clerdo.app; `X-Frame-Options` is not sent.

## Deploying

```
npm run build:cloud
GRID_DEPLOY_KEY=<key> bash server/dev/package.sh     # → server/deploy/deployment.php
```

The deploy key is passed in at packaging time and is never stored in this repository. Upload
`deployment.php` to `domains/grid.clerdo.app/public_html/` and open
`https://grid.clerdo.app/deployment.php?key=<key>`:

| action | does |
|---|---|
| `&go=1` | installs `public_html/*` and `../grid_app/*`, keeps the live `config.php`, checks every file's sha256 |
| `&config=1` | writes `config.php` (first install only) |
| `&client=1` | registers the `grid-web` OAuth client at auth.clerdo.app (first install only) |
| `&migrate=1` | creates and upgrades tables |
| `&probe=1` | checks PHP, extensions, database, encryption, issuer |
| `&delete=1` | removes the deployer: always do this last |

An upgrade is `go=1`, `migrate=1`, `delete=1`.

## Local development

```
npm run build:cloud
bash server/dev/start.sh          # mock IdP :8090, app :8080, sqlite in server/dev/.run
bash server/dev/api-test.sh       # end-to-end API checks on a fresh database
bash server/dev/start.sh stop
```

The mock IdP has two users: `u-hamzeh` and `u-eyas`.
