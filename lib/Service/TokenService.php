<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Service;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\Http\Client\IClientService;
use OCP\IAppConfig;
use OCP\IConfig;
use OCP\Security\ICrypto;
use Psr\Log\LoggerInterface;

class TokenService {
    private const TOKEN_EXPIRY_MARGIN = 300; // 5 minutes

    public function __construct(
        private IConfig $config,
        private IAppConfig $appConfig,
        private ICrypto $crypto,
        private IClientService $clientService,
        private LoggerInterface $logger,
    ) {
    }

    /**
     * Store OAuth2 tokens for a user (encrypted).
     */
    public function store(string $userId, array $tokenResponse): void {
        $this->setEncryptedUserValue(
            $userId, 'access_token', $tokenResponse['access_token']
        );
        $this->setEncryptedUserValue(
            $userId, 'refresh_token', $tokenResponse['refresh_token']
        );

        $expiresAt = time() + (int)$tokenResponse['expires_in'];
        $this->config->setUserValue(
            $userId, Application::APP_ID, 'token_expires', (string)$expiresAt
        );
        $this->config->setUserValue(
            $userId, Application::APP_ID, 'connected_at', (string)time()
        );
    }

    /**
     * Store the user's MS365 email address.
     */
    public function storeUserEmail(string $userId, string $email): void {
        $this->config->setUserValue(
            $userId, Application::APP_ID, 'user_email', $email
        );
    }

    /**
     * Get the user's MS365 email address.
     */
    public function getUserEmail(string $userId): string {
        return $this->config->getUserValue(
            $userId, Application::APP_ID, 'user_email', ''
        );
    }

    /**
     * Get a valid access token, refreshing if necessary.
     * Returns empty string if no token is available.
     */
    public function getAccessToken(string $userId): string {
        $accessToken = $this->getEncryptedUserValue($userId, 'access_token');
        if ($accessToken === '') {
            return '';
        }

        if (!$this->isTokenExpired($userId)) {
            return $accessToken;
        }

        // Token is expired or about to expire, try to refresh
        return $this->refresh($userId);
    }

    /**
     * Check if user has a connected MS365 account.
     */
    public function isConnected(string $userId): bool {
        $token = $this->getEncryptedUserValue($userId, 'refresh_token');
        return $token !== '';
    }

    /**
     * Get connection status info for the settings UI.
     */
    public function getConnectionStatus(string $userId): array {
        $connected = $this->isConnected($userId);

        // Auto-refresh expired token so the UI doesn't show a stale warning
        $tokenExpired = false;
        if ($connected && $this->isTokenExpired($userId)) {
            $this->refresh($userId);
            $tokenExpired = $this->isTokenExpired($userId);
            $connected = $this->isConnected($userId);
        }

        return [
            'connected' => $connected,
            'email' => $connected ? $this->getUserEmail($userId) : '',
            'connectedAt' => $connected ? (int)$this->config->getUserValue(
                $userId, Application::APP_ID, 'connected_at', '0'
            ) : 0,
            'tokenExpired' => $tokenExpired,
        ];
    }

    /**
     * Force a token refresh regardless of expiry time.
     * Used when the API returns 401 (token revoked server-side).
     */
    public function forceRefresh(string $userId): string {
        // Invalidate the stored expiry so getAccessToken() will also refresh next time
        $this->config->setUserValue(
            $userId, Application::APP_ID, 'token_expires', '0'
        );
        return $this->refresh($userId);
    }

    /**
     * Remove all tokens and user data (disconnect).
     */
    public function disconnect(string $userId): void {
        $keys = [
            'access_token', 'refresh_token', 'token_expires',
            'user_email', 'connected_at',
        ];
        foreach ($keys as $key) {
            $this->config->deleteUserValue($userId, Application::APP_ID, $key);
        }
    }

    /**
     * Refresh the access token using the refresh token.
     */
    private function refresh(string $userId): string {
        $refreshToken = $this->getEncryptedUserValue($userId, 'refresh_token');
        if ($refreshToken === '') {
            return '';
        }

        $tenantId = $this->appConfig->getValueString(Application::APP_ID, 'tenant_id');
        $clientId = $this->appConfig->getValueString(Application::APP_ID, 'client_id');
        $clientSecret = $this->appConfig->getValueString(Application::APP_ID, 'client_secret', lazy: true);

        $tokenUrl = "https://login.microsoftonline.com/{$tenantId}/oauth2/v2.0/token";

        try {
            $client = $this->clientService->newClient();
            $response = $client->post($tokenUrl, [
                'body' => [
                    'grant_type' => 'refresh_token',
                    'client_id' => $clientId,
                    'client_secret' => $clientSecret,
                    'refresh_token' => $refreshToken,
                    'scope' => 'User.ReadBasic.All People.Read Calendars.Read Calendars.Read.Shared offline_access',
                ],
            ]);

            $tokenData = json_decode($response->getBody(), true);

            if (!isset($tokenData['access_token'])) {
                $this->logger->error('Token refresh failed: no access_token in response');
                $this->disconnect($userId);
                return '';
            }

            $this->store($userId, $tokenData);
            return $tokenData['access_token'];
        } catch (\Exception $e) {
            $statusCode = method_exists($e, 'getResponse') && $e->getResponse() !== null
                ? $e->getResponse()->getStatusCode() : 0;
            $this->logger->error('Token refresh failed', [
                'statusCode' => $statusCode,
                'app' => Application::APP_ID,
            ]);
            $this->disconnect($userId);
            return '';
        }
    }

    private function isTokenExpired(string $userId): bool {
        $expiresAt = (int)$this->config->getUserValue(
            $userId, Application::APP_ID, 'token_expires', '0'
        );
        return $expiresAt < (time() + self::TOKEN_EXPIRY_MARGIN);
    }

    private function setEncryptedUserValue(string $userId, string $key, string $value): void {
        $encrypted = $this->crypto->encrypt($value);
        $this->config->setUserValue($userId, Application::APP_ID, $key, $encrypted);
    }

    private function getEncryptedUserValue(string $userId, string $key): string {
        $encrypted = $this->config->getUserValue($userId, Application::APP_ID, $key, '');
        if ($encrypted === '') {
            return '';
        }

        try {
            return $this->crypto->decrypt($encrypted);
        } catch (\Exception $e) {
            $this->logger->warning('Failed to decrypt ' . $key . ' for user ' . $userId, [
                'app' => Application::APP_ID,
            ]);
            return '';
        }
    }
}
