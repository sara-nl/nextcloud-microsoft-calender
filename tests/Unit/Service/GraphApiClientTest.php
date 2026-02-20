<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Tests\Unit\Service;

use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\Http\Client\IClient;
use OCP\Http\Client\IClientService;
use OCP\Http\Client\IResponse;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;
use Psr\Log\LoggerInterface;

class GraphApiClientTest extends TestCase {
    private GraphApiClient $graphClient;
    private TokenService&MockObject $tokenService;
    private IClientService&MockObject $clientService;
    private IClient&MockObject $httpClient;
    private LoggerInterface&MockObject $logger;

    protected function setUp(): void {
        $this->tokenService = $this->createMock(TokenService::class);
        $this->clientService = $this->createMock(IClientService::class);
        $this->httpClient = $this->createMock(IClient::class);
        $this->logger = $this->createMock(LoggerInterface::class);

        $this->clientService->method('newClient')
            ->willReturn($this->httpClient);

        $this->graphClient = new GraphApiClient(
            $this->tokenService,
            $this->clientService,
            $this->logger,
        );
    }

    public function testGetReturnsDecodedJson(): void {
        $this->tokenService->method('getAccessToken')
            ->willReturn('valid_token');

        $response = $this->createMock(IResponse::class);
        $response->method('getBody')
            ->willReturn(json_encode(['value' => [['displayName' => 'Test User']]]));

        $this->httpClient->method('get')
            ->willReturn($response);

        $result = $this->graphClient->get('testuser', '/me/people', ['$search' => 'test']);

        $this->assertIsArray($result);
        $this->assertArrayHasKey('value', $result);
        $this->assertEquals('Test User', $result['value'][0]['displayName']);
    }

    public function testPostReturnsDecodedJson(): void {
        $this->tokenService->method('getAccessToken')
            ->willReturn('valid_token');

        $response = $this->createMock(IResponse::class);
        $response->method('getBody')
            ->willReturn(json_encode(['value' => [['scheduleId' => 'user@example.com']]]));

        $this->httpClient->method('post')
            ->willReturn($response);

        $result = $this->graphClient->post('testuser', '/me/calendar/getSchedule', [
            'schedules' => ['user@example.com'],
        ]);

        $this->assertIsArray($result);
        $this->assertEquals('user@example.com', $result['value'][0]['scheduleId']);
    }

    public function testReturnsNullWhenNoToken(): void {
        $this->tokenService->method('getAccessToken')
            ->willReturn('');

        $result = $this->graphClient->get('testuser', '/me/people');

        $this->assertNull($result);
    }

    public function testGetProfilePhotoReturnsDataUri(): void {
        $this->tokenService->method('getAccessToken')
            ->willReturn('valid_token');

        $response = $this->createMock(IResponse::class);
        $response->method('getHeader')
            ->willReturn('image/jpeg');
        $response->method('getBody')
            ->willReturn('fake-image-data');

        $this->httpClient->method('get')
            ->willReturn($response);

        $result = $this->graphClient->getProfilePhoto('testuser', 'ms-user-id');

        $this->assertStringStartsWith('data:image/jpeg;base64,', $result);
    }

    public function testGetProfilePhotoReturnsNullOnError(): void {
        $this->tokenService->method('getAccessToken')
            ->willReturn('valid_token');

        $this->httpClient->method('get')
            ->willThrowException(new \Exception('404 Not Found', 404));

        $result = $this->graphClient->getProfilePhoto('testuser', 'ms-user-id');

        $this->assertNull($result);
    }
}
