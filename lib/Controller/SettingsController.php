<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Controller;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\AppFramework\Controller;
use OCP\AppFramework\Http\JSONResponse;
use OCP\IAppConfig;
use OCP\IRequest;

class SettingsController extends Controller {
    private const ADMIN_KEYS = [
        'tenant_id',
        'client_id',
        'client_secret',
        'auth_method',
        'cache_ttl',
        'reply_notifications_enabled',
    ];

    private const SENSITIVE_KEYS = [
        'client_secret',
    ];

    public function __construct(
        string $appName,
        IRequest $request,
        private IAppConfig $appConfig,
    ) {
        parent::__construct($appName, $request);
    }

    /**
     * Get admin settings (masks sensitive values).
     */
    public function getAdmin(): JSONResponse {
        $settings = [];
        foreach (self::ADMIN_KEYS as $key) {
            if ($key === 'cache_ttl') {
                $settings[$key] = $this->appConfig->getValueInt(Application::APP_ID, $key, 300);
                continue;
            }
            if ($key === 'reply_notifications_enabled') {
                $settings[$key] = $this->appConfig->getValueBool(Application::APP_ID, $key, false);
                continue;
            }
            $lazy = in_array($key, self::SENSITIVE_KEYS, true);
            $value = $this->appConfig->getValueString(Application::APP_ID, $key, lazy: $lazy);
            if ($lazy && $value !== '') {
                $settings[$key] = str_repeat('*', 8);
            } else {
                $settings[$key] = $value;
            }
        }
        return new JSONResponse($settings);
    }

    /**
     * Save admin settings.
     */
    public function saveAdmin(): JSONResponse {
        foreach (self::ADMIN_KEYS as $key) {
            $value = $this->request->getParam($key);
            if ($value === null) {
                continue;
            }

            // Don't overwrite secret with masked or empty value
            if (in_array($key, self::SENSITIVE_KEYS, true) && ($value === '' || preg_match('/^\*+$/', $value))) {
                continue;
            }

            if (in_array($key, self::SENSITIVE_KEYS, true)) {
                $this->appConfig->setValueString(
                    Application::APP_ID, $key, $value, sensitive: true
                );
            } elseif ($key === 'cache_ttl') {
                $this->appConfig->setValueInt(
                    Application::APP_ID, $key, (int)$value
                );
            } elseif ($key === 'reply_notifications_enabled') {
                $this->appConfig->setValueBool(
                    Application::APP_ID, $key, (bool)$value
                );
            } else {
                $this->appConfig->setValueString(
                    Application::APP_ID, $key, $value
                );
            }
        }

        return new JSONResponse(['status' => 'ok']);
    }
}
