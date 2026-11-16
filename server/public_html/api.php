<?php
declare(strict_types=1);

/** Grid front controller. All /api/*, /auth/*, /share/* requests land here via .htaccess. */
require __DIR__ . '/../grid_app/bootstrap.php';
require __DIR__ . '/../grid_app/router.php';

Grid\Router::dispatch();
