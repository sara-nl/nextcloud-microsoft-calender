<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Service;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\Http\Client\IClientService;
use Psr\Log\LoggerInterface;

class GraphApiClient {
    private const BASE_URL = 'https://graph.microsoft.com/v1.0';
    private const MAX_RETRIES = 3;
    private const INITIAL_BACKOFF_MS = 1000;

    public function __construct(
        private TokenService $tokenService,
        private IClientService $clientService,
        private LoggerInterface $logger,
    ) {
    }

    /**
     * Perform a GET request to the Graph API.
     *
     * @return array|null Decoded JSON response, or null on failure
     */
    public function get(string $userId, string $endpoint, array $queryParams = []): ?array {
        return $this->request($userId, 'GET', $endpoint, ['query' => $queryParams]);
    }

    /**
     * Perform a POST request to the Graph API.
     *
     * @return array|null Decoded JSON response, or null on failure
     */
    public function post(string $userId, string $endpoint, array $body = []): ?array {
        return $this->request($userId, 'POST', $endpoint, [
            'json' => $body,
            'headers' => ['Content-Type' => 'application/json'],
        ]);
    }

    /**
     * Get a user's profile photo as a base64 data URI.
     */
    public function getProfilePhoto(string $userId, string $msUserId): ?string {
        $token = $this->tokenService->getAccessToken($userId);
        if ($token === '') {
            return null;
        }

        $url = self::BASE_URL . "/users/{$msUserId}/photo/\$value";

        try {
            $client = $this->clientService->newClient();
            $response = $client->get($url, [
                'headers' => [
                    'Authorization' => "Bearer {$token}",
                ],
            ]);

            $contentType = $response->getHeader('Content-Type');
            $imageData = $response->getBody();
            return "data:{$contentType};base64," . base64_encode($imageData);
        } catch (\Exception $e) {
            $this->logger->debug('Could not fetch profile photo', [
                'app' => Application::APP_ID,
                'statusCode' => $this->getStatusCode($e),
            ]);
            return null;
        }
    }

    /**
     * Execute a Graph API request with retry logic.
     */
    private function request(string $userId, string $method, string $endpoint, array $options): ?array {
        $retries = 0;
        $tokenRefreshed = false;

        while ($retries <= self::MAX_RETRIES) {
            $token = $this->tokenService->getAccessToken($userId);
            if ($token === '') {
                $this->logger->warning('No valid access token available', [
                    'app' => Application::APP_ID,
                ]);
                return null;
            }

            $options['headers'] = array_merge($options['headers'] ?? [], [
                'Authorization' => "Bearer {$token}",
                'ConsistencyLevel' => 'eventual', // Required for $search on /users
            ]);

            $url = self::BASE_URL . $endpoint;

            try {
                $client = $this->clientService->newClient();
                $response = match ($method) {
                    'POST' => $client->post($url, $options),
                    default => $client->get($url, $options),
                };

                return json_decode($response->getBody(), true);
            } catch (\Exception $e) {
                $statusCode = $this->getStatusCode($e);

                // 401: force token refresh once, then retry
                if ($statusCode === 401 && !$tokenRefreshed) {
                    $this->logger->info('Got 401, forcing token refresh', [
                        'app' => Application::APP_ID,
                    ]);
                    $this->tokenService->forceRefresh($userId);
                    $tokenRefreshed = true;
                    continue;
                }

                // 429: rate limited, respect Retry-After
                if ($statusCode === 429) {
                    $retryAfter = $this->getRetryAfter($e);
                    $this->logger->info("Rate limited, waiting {$retryAfter}s", [
                        'app' => Application::APP_ID,
                    ]);
                    usleep($retryAfter * 1_000_000);
                    $retries++;
                    continue;
                }

                // 503/504: service unavailable, exponential backoff
                if (in_array($statusCode, [503, 504], true)) {
                    $backoffMs = self::INITIAL_BACKOFF_MS * (2 ** $retries);
                    $this->logger->info("Service unavailable ({$statusCode}), backoff {$backoffMs}ms", [
                        'app' => Application::APP_ID,
                    ]);
                    usleep($backoffMs * 1000);
                    $retries++;
                    continue;
                }

                // 403: insufficient permissions
                if ($statusCode === 403) {
                    $this->logger->warning('Insufficient Graph API permissions', [
                        'endpoint' => $endpoint,
                        'statusCode' => $statusCode,
                        'app' => Application::APP_ID,
                    ]);
                    return null;
                }

                // Other errors: don't retry
                $this->logger->error('Graph API request failed', [
                    'endpoint' => $endpoint,
                    'statusCode' => $statusCode,
                    'app' => Application::APP_ID,
                ]);
                return null;
            }
        }

        $this->logger->error('Graph API request failed after max retries', [
            'endpoint' => $endpoint,
            'app' => Application::APP_ID,
        ]);
        return null;
    }

    /**
     * Extract HTTP status code from exception.
     * NC's HTTP client wraps Guzzle — the status is on the response object.
     */
    private function getStatusCode(\Exception $e): int {
        if (method_exists($e, 'getResponse') && ($response = $e->getResponse()) !== null) {
            if (method_exists($response, 'getStatusCode')) {
                return (int)$response->getStatusCode();
            }
        }
        return 0;
    }

    private function getRetryAfter(\Exception $e): int {
        if (method_exists($e, 'getResponse') && ($response = $e->getResponse()) !== null) {
            if (method_exists($response, 'getHeader')) {
                $header = $response->getHeader('Retry-After');
                if ($header !== '' && is_numeric($header)) {
                    return (int)$header;
                }
            }
        }
        return 10;
    }
}
