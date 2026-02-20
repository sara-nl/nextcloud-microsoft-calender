<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Tests\Unit\Service;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\Http\Client\IClient;
use OCP\Http\Client\IClientService;
use OCP\Http\Client\IResponse;
use OCP\IAppConfig;
use OCP\IConfig;
use OCP\Security\ICrypto;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;
use Psr\Log\LoggerInterface;

class TokenServiceTest extends TestCase {
    private TokenService $tokenService;
    private IConfig&MockObject $config;
    private IAppConfig&MockObject $appConfig;
    private ICrypto&MockObject $crypto;
    private IClientService&MockObject $clientService;
    private LoggerInterface&MockObject $logger;

    protected function setUp(): void {
        $this->config = $this->createMock(IConfig::class);
        $this->appConfig = $this->createMock(IAppConfig::class);
        $this->crypto = $this->createMock(ICrypto::class);
        $this->clientService = $this->createMock(IClientService::class);
        $this->logger = $this->createMock(LoggerInterface::class);

        $this->tokenService = new TokenService(
            $this->config,
            $this->appConfig,
            $this->crypto,
            $this->clientService,
            $this->logger,
        );
    }

    public function testStoreEncryptsTokens(): void {
        $this->crypto->expects($this->exactly(2))
            ->method('encrypt')
            ->willReturnCallback(fn(string $val) => 'encrypted_' . $val);

        $this->config->expects($this->exactly(4))
            ->method('setUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key, string $value) {
                $this->assertEquals('testuser', $userId);
                $this->assertEquals(Application::APP_ID, $app);

                match ($key) {
                    'access_token' => $this->assertEquals('encrypted_my_access_token', $value),
                    'refresh_token' => $this->assertEquals('encrypted_my_refresh_token', $value),
                    'token_expires' => $this->assertIsNumeric($value),
                    'connected_at' => $this->assertIsNumeric($value),
                    default => $this->fail("Unexpected key: {$key}"),
                };
            });

        $this->tokenService->store('testuser', [
            'access_token' => 'my_access_token',
            'refresh_token' => 'my_refresh_token',
            'expires_in' => 3600,
        ]);
    }

    public function testGetAccessTokenReturnsDecryptedToken(): void {
        $this->config->method('getUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key) {
                return match ($key) {
                    'access_token' => 'encrypted_token',
                    'token_expires' => (string)(time() + 3600),
                    default => '',
                };
            });

        $this->crypto->method('decrypt')
            ->with('encrypted_token')
            ->willReturn('decrypted_access_token');

        $result = $this->tokenService->getAccessToken('testuser');
        $this->assertEquals('decrypted_access_token', $result);
    }

    public function testGetAccessTokenReturnsEmptyWhenNoToken(): void {
        $this->config->method('getUserValue')
            ->willReturn('');

        $result = $this->tokenService->getAccessToken('testuser');
        $this->assertEquals('', $result);
    }

    public function testIsConnectedReturnsTrueWhenRefreshTokenExists(): void {
        $this->config->method('getUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key) {
                if ($key === 'refresh_token') {
                    return 'encrypted_refresh';
                }
                return '';
            });

        $this->crypto->method('decrypt')
            ->willReturn('refresh_token_value');

        $this->assertTrue($this->tokenService->isConnected('testuser'));
    }

    public function testIsConnectedReturnsFalseWhenNoRefreshToken(): void {
        $this->config->method('getUserValue')
            ->willReturn('');

        $this->assertFalse($this->tokenService->isConnected('testuser'));
    }

    public function testDisconnectRemovesAllKeys(): void {
        $this->config->expects($this->exactly(5))
            ->method('deleteUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key) {
                $this->assertEquals('testuser', $userId);
                $this->assertEquals(Application::APP_ID, $app);
                $this->assertContains($key, [
                    'access_token', 'refresh_token', 'token_expires',
                    'user_email', 'connected_at',
                ]);
            });

        $this->tokenService->disconnect('testuser');
    }

    public function testGetConnectionStatusWhenConnected(): void {
        $connectedAt = time() - 3600;

        $this->config->method('getUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key) use ($connectedAt) {
                return match ($key) {
                    'refresh_token' => 'encrypted_refresh',
                    'user_email' => 'user@example.com',
                    'connected_at' => (string)$connectedAt,
                    'access_token' => 'encrypted_access',
                    'token_expires' => (string)(time() + 3600),
                    default => '',
                };
            });

        $this->crypto->method('decrypt')
            ->willReturn('decrypted_value');

        $status = $this->tokenService->getConnectionStatus('testuser');

        $this->assertTrue($status['connected']);
        $this->assertEquals('user@example.com', $status['email']);
        $this->assertEquals($connectedAt, $status['connectedAt']);
        $this->assertFalse($status['tokenExpired']);
    }

    public function testGetConnectionStatusWhenDisconnected(): void {
        $this->config->method('getUserValue')
            ->willReturn('');

        $status = $this->tokenService->getConnectionStatus('testuser');

        $this->assertFalse($status['connected']);
        $this->assertEquals('', $status['email']);
        $this->assertEquals(0, $status['connectedAt']);
    }

    public function testDecryptFailureReturnsEmpty(): void {
        $this->config->method('getUserValue')
            ->willReturnCallback(function (string $userId, string $app, string $key) {
                return match ($key) {
                    'access_token' => 'corrupted_encrypted_data',
                    default => '',
                };
            });

        $this->crypto->method('decrypt')
            ->willThrowException(new \Exception('Decryption failed'));

        $result = $this->tokenService->getAccessToken('testuser');
        $this->assertEquals('', $result);
    }
}
