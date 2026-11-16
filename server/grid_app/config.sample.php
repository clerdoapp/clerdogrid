<?php
/**
 * Grid configuration — copy to config.php (chmod 0600). Never commit config.php.
 */
return [
    'db' => [
        'driver' => 'mysql',          // 'mysql' (MariaDB, production) or 'sqlite' (dev only)
        'host' => '127.0.0.1',
        'port' => 3306,
        'name' => 'clerdo_core',
        'user' => 'clerdo_user',
        'pass' => 'change-me',
        // 'path' => '/path/to/grid.sqlite',   // sqlite only
    ],

    // base64 of 32 random bytes:  php -r 'echo base64_encode(random_bytes(32)), PHP_EOL;'
    // For rotation use an array: ['k2' => '<new>', 'k1' => '<old>'] — first entry encrypts, all decrypt.
    'encryption_key' => 'REPLACE_WITH_BASE64_32_BYTES',

    // Key for deployment.php maintenance actions.
    'install_key' => 'set-a-long-random-value',

    'auth' => [
        'issuer' => 'https://auth.clerdo.app',
        'client_id' => 'grid-web',
        'redirect_uri' => 'https://grid.clerdo.app/auth/callback',
        'authorize_url' => 'https://auth.clerdo.app/authorize',
        'token_url' => 'https://auth.clerdo.app/token',
        'userinfo_url' => 'https://auth.clerdo.app/userinfo',
        'revoke_url' => 'https://auth.clerdo.app/revoke',
        'scope' => 'openid profile email',
    ],

    'pro' => [
        'everyone_has_pro' => true,
    ],

    'limits' => [
        'free_projects' => 5,
        'free_doc_bytes' => 16 * 1024 * 1024,
        'pro_doc_bytes' => 25 * 1024 * 1024,
        'share_per_minute' => 60,     // GET /api/share/:token per IP per minute
        'auth_per_minute' => 30,      // /auth/* per IP per minute
    ],

    'app' => [
        'base_url' => 'https://grid.clerdo.app',
    ],
];
