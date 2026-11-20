<?php
/**
 * Local dev config. GRID_DB=sqlite (default in the sandbox) or GRID_DB=mysql.
 */
$driver = getenv('GRID_DB') ?: 'sqlite';
$idp = getenv('GRID_IDP') ?: 'http://127.0.0.1:' . (getenv('GRID_IDP_PORT') ?: '8090');
$base = getenv('GRID_BASE') ?: 'http://127.0.0.1:' . (getenv('GRID_APP_PORT') ?: '8080');

return [
    'db' => $driver === 'sqlite'
        ? ['driver' => 'sqlite', 'path' => getenv('GRID_SQLITE') ?: sys_get_temp_dir() . '/grid-dev.sqlite']
        : ['driver' => 'mysql', 'host' => '127.0.0.1', 'port' => 3306, 'name' => 'grid', 'user' => 'grid', 'pass' => 'grid'],

    // Dev-only key. Generate your own for anything real.
    'encryption_key' => 'MDEyMzQ1Njc4OWFiY2RlZjAxMjM0NTY3ODlhYmNkZWY=',
    'install_key' => 'dev-install-key',

    'auth' => [
        'issuer' => $idp,
        'client_id' => 'grid-web',
        'redirect_uri' => $base . '/auth/callback',
        'authorize_url' => $idp . '/authorize',
        'token_url' => $idp . '/token',
        'userinfo_url' => $idp . '/userinfo',
        'revoke_url' => $idp . '/revoke',
        'scope' => 'openid profile email',
    ],

    'pro' => ['everyone_has_pro' => getenv('GRID_EVERYONE_PRO') === false || getenv('GRID_EVERYONE_PRO') === '1'],

    'limits' => [
        'free_projects' => 5,
        'free_doc_bytes' => 16 * 1024 * 1024,
        'pro_doc_bytes' => 25 * 1024 * 1024,
        'share_per_minute' => (int)(getenv('GRID_SHARE_RPM') ?: 60),
        'auth_per_minute' => 300,
    ],

    'app' => ['base_url' => $base],

    // Where workbooks are stored; empty keeps them in Grid's own database only.
    'locker' => ['url' => getenv('GRID_LOCKER_URL') ?: ''],
];
