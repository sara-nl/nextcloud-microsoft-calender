<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Settings;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\AppFramework\Http\TemplateResponse;
use OCP\IUserSession;
use OCP\Settings\ISettings;

class PersonalSettings implements ISettings {
    public function __construct(
        private TokenService $tokenService,
        private IUserSession $userSession,
    ) {
    }

    public function getForm(): TemplateResponse {
        $user = $this->userSession->getUser();
        $status = $user !== null
            ? $this->tokenService->getConnectionStatus($user->getUID())
            : ['connected' => false, 'email' => '', 'connectedAt' => 0, 'tokenExpired' => false];

        return new TemplateResponse(Application::APP_ID, 'personal', $status);
    }

    public function getSection(): string {
        return Application::APP_ID;
    }

    public function getPriority(): int {
        return 10;
    }
}
