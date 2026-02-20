<?php

declare(strict_types=1);

return [
    'routes' => [
        ['name' => 'OAuth#authorize', 'url' => '/oauth/authorize', 'verb' => 'GET'],
        ['name' => 'OAuth#callback', 'url' => '/oauth/callback', 'verb' => 'GET'],
        ['name' => 'OAuth#disconnect', 'url' => '/oauth/disconnect', 'verb' => 'POST'],
        ['name' => 'OAuth#status', 'url' => '/oauth/status', 'verb' => 'GET'],
        ['name' => 'Settings#saveAdmin', 'url' => '/settings/admin', 'verb' => 'PUT'],
        ['name' => 'Settings#getAdmin', 'url' => '/settings/admin', 'verb' => 'GET'],
    ],
];
