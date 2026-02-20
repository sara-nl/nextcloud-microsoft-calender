<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Settings;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\AppFramework\Http\TemplateResponse;
use OCP\IAppConfig;
use OCP\Settings\ISettings;

class AdminSettings implements ISettings {
    public function __construct(
        private IAppConfig $appConfig,
    ) {
    }

    public function getForm(): TemplateResponse {
        $params = [
            'tenant_id' => $this->appConfig->getValueString(Application::APP_ID, 'tenant_id'),
            'client_id' => $this->appConfig->getValueString(Application::APP_ID, 'client_id'),
            'client_secret_set' => $this->appConfig->getValueString(Application::APP_ID, 'client_secret') !== '',
            'auth_method' => $this->appConfig->getValueString(Application::APP_ID, 'auth_method', 'client_secret'),
            'cache_ttl' => $this->appConfig->getValueInt(Application::APP_ID, 'cache_ttl', 300),
        ];

        return new TemplateResponse(Application::APP_ID, 'admin', $params);
    }

    public function getSection(): string {
        return Application::APP_ID;
    }

    public function getPriority(): int {
        return 10;
    }
}
