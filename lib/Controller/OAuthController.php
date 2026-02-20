<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Controller;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\AppFramework\Controller;
use OCP\AppFramework\Http;
use OCP\AppFramework\Http\JSONResponse;
use OCP\AppFramework\Http\RedirectResponse;
use OCP\Http\Client\IClientService;
use OCP\IAppConfig;
use OCP\IRequest;
use OCP\ISession;
use OCP\IURLGenerator;
use OCP\IUserSession;
use Psr\Log\LoggerInterface;

class OAuthController extends Controller {
    private const SCOPES = 'User.ReadBasic.All People.Read Calendars.Read Calendars.Read.Shared offline_access';

    public function __construct(
        string $appName,
        IRequest $request,
        private IAppConfig $appConfig,
        private ISession $session,
        private IUserSession $userSession,
        private IURLGenerator $urlGenerator,
        private IClientService $clientService,
        private TokenService $tokenService,
        private GraphApiClient $graphClient,
        private LoggerInterface $logger,
    ) {
        parent::__construct($appName, $request);
    }

    /**
     * Start the OAuth2 authorization flow with PKCE.
     *
     * @NoAdminRequired
     * @NoCSRFRequired
     */
    public function authorize(): RedirectResponse {
        $tenantId = $this->appConfig->getValueString(Application::APP_ID, 'tenant_id');
        $clientId = $this->appConfig->getValueString(Application::APP_ID, 'client_id');

        if ($tenantId === '' || $clientId === '') {
            return new RedirectResponse(
                $this->urlGenerator->linkToRoute('settings.PersonalSettings.index', ['section' => 'nc_ms365_calendar'])
            );
        }

        // Generate PKCE code verifier and challenge
        $codeVerifier = $this->generateCodeVerifier();
        $codeChallenge = $this->generateCodeChallenge($codeVerifier);

        // Generate state parameter for CSRF protection
        $state = bin2hex(random_bytes(32));

        // Store in session
        $this->session->set('nc_ms365_calendar_pkce_verifier', $codeVerifier);
        $this->session->set('nc_ms365_calendar_oauth_state', $state);

        $redirectUri = $this->urlGenerator->linkToRouteAbsolute(
            Application::APP_ID . '.OAuth.callback'
        );

        $authUrl = "https://login.microsoftonline.com/{$tenantId}/oauth2/v2.0/authorize?" . http_build_query([
            'client_id' => $clientId,
            'response_type' => 'code',
            'redirect_uri' => $redirectUri,
            'scope' => self::SCOPES,
            'state' => $state,
            'code_challenge' => $codeChallenge,
            'code_challenge_method' => 'S256',
            'response_mode' => 'query',
        ]);

        return new RedirectResponse($authUrl);
    }

    /**
     * Handle the OAuth2 callback from Microsoft.
     *
     * @NoAdminRequired
     * @NoCSRFRequired
     */
    public function callback(): RedirectResponse {
        $settingsUrl = $this->urlGenerator->linkToRoute(
            'settings.PersonalSettings.index', ['section' => 'nc_ms365_calendar']
        );

        $code = $this->request->getParam('code', '');
        $state = $this->request->getParam('state', '');
        $error = $this->request->getParam('error', '');

        // Check for errors from Microsoft
        if ($error !== '') {
            $errorDescription = $this->request->getParam('error_description', 'Unknown error');
            $this->logger->error('OAuth callback error: ' . $errorDescription, [
                'app' => Application::APP_ID,
            ]);
            return new RedirectResponse($settingsUrl);
        }

        // Validate state parameter
        $savedState = $this->session->get('nc_ms365_calendar_oauth_state');
        if ($state === '' || $state !== $savedState) {
            $this->logger->error('OAuth state mismatch', [
                'app' => Application::APP_ID,
            ]);
            return new RedirectResponse($settingsUrl);
        }

        // Get PKCE verifier from session
        $codeVerifier = $this->session->get('nc_ms365_calendar_pkce_verifier');

        // Clean up session
        $this->session->remove('nc_ms365_calendar_oauth_state');
        $this->session->remove('nc_ms365_calendar_pkce_verifier');

        // Exchange code for tokens
        $tenantId = $this->appConfig->getValueString(Application::APP_ID, 'tenant_id');
        $clientId = $this->appConfig->getValueString(Application::APP_ID, 'client_id');
        $clientSecret = $this->appConfig->getValueString(Application::APP_ID, 'client_secret', lazy: true);
        $redirectUri = $this->urlGenerator->linkToRouteAbsolute(
            Application::APP_ID . '.OAuth.callback'
        );

        try {
            $client = $this->clientService->newClient();
            $response = $client->post(
                "https://login.microsoftonline.com/{$tenantId}/oauth2/v2.0/token",
                [
                    'body' => [
                        'grant_type' => 'authorization_code',
                        'code' => $code,
                        'redirect_uri' => $redirectUri,
                        'client_id' => $clientId,
                        'client_secret' => $clientSecret,
                        'code_verifier' => $codeVerifier,
                    ],
                ]
            );

            $tokenData = json_decode($response->getBody(), true);

            if (!isset($tokenData['access_token'])) {
                $this->logger->error('Token exchange failed: no access_token in response', [
                    'app' => Application::APP_ID,
                ]);
                return new RedirectResponse($settingsUrl);
            }

            $user = $this->userSession->getUser();
            if ($user === null) {
                $this->logger->error('OAuth callback: no active user session', [
                    'app' => Application::APP_ID,
                ]);
                return new RedirectResponse($settingsUrl);
            }
            $userId = $user->getUID();

            // Store tokens
            $this->tokenService->store($userId, $tokenData);

            // Fetch and store user email from Graph API
            $this->fetchAndStoreUserEmail($userId);

            $this->logger->info('Successfully connected MS365 account', [
                'userId' => $userId,
                'app' => Application::APP_ID,
            ]);
        } catch (\Exception $e) {
            $this->logger->error('Token exchange failed: ' . $e->getMessage(), [
                'app' => Application::APP_ID,
            ]);
        }

        return new RedirectResponse($settingsUrl);
    }

    /**
     * Disconnect the MS365 account.
     *
     * @NoAdminRequired
     */
    public function disconnect(): JSONResponse {
        $user = $this->userSession->getUser();
        if ($user === null) {
            return new JSONResponse(['error' => 'Not logged in'], Http::STATUS_UNAUTHORIZED);
        }
        $userId = $user->getUID();
        $this->tokenService->disconnect($userId);

        $this->logger->info('Disconnected MS365 account', [
            'userId' => $userId,
            'app' => Application::APP_ID,
        ]);

        return new JSONResponse(['status' => 'ok']);
    }

    /**
     * Get the current connection status.
     *
     * @NoAdminRequired
     */
    public function status(): JSONResponse {
        $user = $this->userSession->getUser();
        if ($user === null) {
            return new JSONResponse(['error' => 'Not logged in'], Http::STATUS_UNAUTHORIZED);
        }
        $userId = $user->getUID();
        return new JSONResponse($this->tokenService->getConnectionStatus($userId));
    }

    private function fetchAndStoreUserEmail(string $userId): void {
        $result = $this->graphClient->get($userId, '/me', [
            '$select' => 'mail,userPrincipalName',
        ]);

        if ($result !== null) {
            $email = $result['mail'] ?? $result['userPrincipalName'] ?? '';
            if ($email !== '') {
                $this->tokenService->storeUserEmail($userId, $email);
            }
        }
    }

    private function generateCodeVerifier(): string {
        return rtrim(strtr(base64_encode(random_bytes(64)), '+/', '-_'), '=');
    }

    private function generateCodeChallenge(string $verifier): string {
        return rtrim(strtr(base64_encode(hash('sha256', $verifier, true)), '+/', '-_'), '=');
    }
}
