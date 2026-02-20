<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Tests\Unit\AddressBook;

use OCA\NcMs365Calendar\AddressBook\MsGraphAddressBook;
use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\ICache;
use OCP\ICacheFactory;
use OCP\IUser;
use OCP\IUserSession;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;
use Psr\Log\LoggerInterface;

class MsGraphAddressBookTest extends TestCase {
    private MsGraphAddressBook $addressBook;
    private GraphApiClient&MockObject $graphClient;
    private TokenService&MockObject $tokenService;
    private IUserSession&MockObject $userSession;
    private ICache&MockObject $cache;
    private LoggerInterface&MockObject $logger;

    protected function setUp(): void {
        $this->graphClient = $this->createMock(GraphApiClient::class);
        $this->tokenService = $this->createMock(TokenService::class);
        $this->userSession = $this->createMock(IUserSession::class);
        $this->cache = $this->createMock(ICache::class);
        $this->logger = $this->createMock(LoggerInterface::class);

        $cacheFactory = $this->createMock(ICacheFactory::class);
        $cacheFactory->method('createDistributed')
            ->willReturn($this->cache);

        $user = $this->createMock(IUser::class);
        $user->method('getUID')->willReturn('testuser');
        $this->userSession->method('getUser')->willReturn($user);

        $this->tokenService->method('isConnected')->willReturn(true);

        $this->addressBook = new MsGraphAddressBook(
            $this->graphClient,
            $this->tokenService,
            $this->userSession,
            $cacheFactory,
            $this->logger,
        );
    }

    public function testGetKeyReturnsExpectedValue(): void {
        $this->assertEquals('ms365-addressbook', $this->addressBook->getKey());
    }

    public function testGetDisplayNameReturnsMicrosoft365(): void {
        $this->assertEquals('Microsoft 365', $this->addressBook->getDisplayName());
    }

    public function testGetPermissionsReturnsReadOnly(): void {
        $this->assertEquals(\OCP\Constants::PERMISSION_READ, $this->addressBook->getPermissions());
    }

    public function testSearchReturnsCachedResults(): void {
        $cachedResults = [
            ['FN' => 'Jan de Vries', 'EMAIL' => ['j.devries@org.nl']],
        ];

        $this->cache->method('get')
            ->willReturn($cachedResults);

        $results = $this->addressBook->search('jan', [], []);

        $this->assertCount(1, $results);
        $this->assertEquals('Jan de Vries', $results[0]['FN']);
    }

    public function testSearchMergesPeopleAndUsers(): void {
        $this->cache->method('get')
            ->willReturn(null);

        // People API response
        $this->graphClient->method('get')
            ->willReturnCallback(function (string $userId, string $endpoint) {
                if (str_contains($endpoint, 'people')) {
                    return [
                        'value' => [
                            [
                                'displayName' => 'Jan de Vries',
                                'emailAddresses' => [['address' => 'j.devries@org.nl']],
                                'department' => 'ICT',
                                'id' => 'people-1',
                            ],
                        ],
                    ];
                }
                if (str_contains($endpoint, 'users')) {
                    return [
                        'value' => [
                            [
                                'displayName' => 'Pieter Jansen',
                                'mail' => 'p.jansen@org.nl',
                                'department' => 'HR',
                                'id' => 'user-1',
                            ],
                            // Duplicate of People API result (should be deduped)
                            [
                                'displayName' => 'Jan de Vries',
                                'mail' => 'j.devries@org.nl',
                                'department' => 'ICT',
                                'id' => 'user-2',
                            ],
                        ],
                    ];
                }
                return null;
            });

        // No photos
        $this->graphClient->method('getProfilePhoto')
            ->willReturn(null);

        $results = $this->addressBook->search('jan', [], []);

        // Should have 2 results (deduped by email)
        $this->assertCount(2, $results);

        $emails = array_map(fn($c) => $c['EMAIL'][0], $results);
        $this->assertContains('j.devries@org.nl', $emails);
        $this->assertContains('p.jansen@org.nl', $emails);
    }

    public function testSearchReturnsEmptyForShortQuery(): void {
        $results = $this->addressBook->search('j', [], []);
        $this->assertEmpty($results);
    }

    public function testSearchReturnsEmptyWhenNotConnected(): void {
        $this->tokenService = $this->createMock(TokenService::class);
        $this->tokenService->method('isConnected')->willReturn(false);

        $cacheFactory = $this->createMock(ICacheFactory::class);
        $cacheFactory->method('createDistributed')->willReturn($this->cache);

        $addressBook = new MsGraphAddressBook(
            $this->graphClient,
            $this->tokenService,
            $this->userSession,
            $cacheFactory,
            $this->logger,
        );

        $results = $addressBook->search('jan', [], []);
        $this->assertEmpty($results);
    }

    public function testSearchRespectsLimit(): void {
        $this->cache->method('get')
            ->willReturn(null);

        // Return many results from People API
        $people = [];
        for ($i = 0; $i < 20; $i++) {
            $people[] = [
                'displayName' => "User {$i}",
                'emailAddresses' => [['address' => "user{$i}@org.nl"]],
                'department' => 'Test',
                'id' => "id-{$i}",
            ];
        }

        $this->graphClient->method('get')
            ->willReturnCallback(function (string $userId, string $endpoint) use ($people) {
                if (str_contains($endpoint, 'people')) {
                    return ['value' => $people];
                }
                return ['value' => []];
            });

        $this->graphClient->method('getProfilePhoto')
            ->willReturn(null);

        $results = $this->addressBook->search('user', [], ['limit' => 5]);
        $this->assertCount(5, $results);
    }

    public function testCreateOrUpdateReturnsEmpty(): void {
        $this->assertEmpty($this->addressBook->createOrUpdate([]));
    }

    public function testDeleteReturnsFalse(): void {
        $this->assertFalse($this->addressBook->delete('any-id'));
    }
}
