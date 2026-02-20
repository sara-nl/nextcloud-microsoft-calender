<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Tests\Unit\Calendar;

use DateTimeImmutable;
use OCA\NcMs365Calendar\Calendar\MsGraphFreeBusyProvider;
use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\ICache;
use OCP\ICacheFactory;
use OCP\IUser;
use OCP\IUserSession;
use PHPUnit\Framework\MockObject\MockObject;
use PHPUnit\Framework\TestCase;
use Psr\Log\LoggerInterface;

class MsGraphFreeBusyProviderTest extends TestCase {
    private MsGraphFreeBusyProvider $provider;
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

        $this->provider = new MsGraphFreeBusyProvider(
            $this->graphClient,
            $this->tokenService,
            $this->userSession,
            $cacheFactory,
            $this->logger,
        );
    }

    public function testGetIdReturnsExpectedValue(): void {
        $this->assertEquals('nc_ms365_calendar-freebusy', $this->provider->getId());
    }

    public function testHandlesMissingReturnsTrue(): void {
        $this->assertTrue($this->provider->handlesMissing('any-uri'));
    }

    public function testGetFreeBusyReturnsBusyPeriods(): void {
        $this->cache->method('get')->willReturn(null);

        $graphResponse = [
            'value' => [
                [
                    'scheduleId' => 'user@example.com',
                    'scheduleItems' => [
                        [
                            'status' => 'busy',
                            'start' => ['dateTime' => '2026-02-17T09:00:00'],
                            'end' => ['dateTime' => '2026-02-17T10:00:00'],
                        ],
                        [
                            'status' => 'tentative',
                            'start' => ['dateTime' => '2026-02-17T14:00:00'],
                            'end' => ['dateTime' => '2026-02-17T15:00:00'],
                        ],
                        [
                            'status' => 'free',
                            'start' => ['dateTime' => '2026-02-17T11:00:00'],
                            'end' => ['dateTime' => '2026-02-17T12:00:00'],
                        ],
                    ],
                ],
            ],
        ];

        $this->graphClient->method('post')
            ->willReturn($graphResponse);

        $start = new DateTimeImmutable('2026-02-17T08:00:00');
        $end = new DateTimeImmutable('2026-02-17T18:00:00');

        $results = $this->provider->getFreeBusy(
            'event-uri',
            'calendar-uri',
            ['mailto:user@example.com'],
            $start,
            $end,
        );

        $this->assertCount(1, $results);

        $vCalendar = $results[0];
        $vfreebusy = $vCalendar->VFREEBUSY;
        $this->assertNotNull($vfreebusy);

        // Should have 2 FREEBUSY entries (busy + tentative, free is skipped)
        $freebusyProps = $vfreebusy->select('FREEBUSY');
        $this->assertCount(2, $freebusyProps);
    }

    public function testGetFreeBusyStatusMapping(): void {
        $this->cache->method('get')->willReturn(null);

        $graphResponse = [
            'value' => [
                [
                    'scheduleId' => 'user@example.com',
                    'scheduleItems' => [
                        [
                            'status' => 'busy',
                            'start' => ['dateTime' => '2026-02-17T09:00:00'],
                            'end' => ['dateTime' => '2026-02-17T10:00:00'],
                        ],
                        [
                            'status' => 'tentative',
                            'start' => ['dateTime' => '2026-02-17T10:00:00'],
                            'end' => ['dateTime' => '2026-02-17T11:00:00'],
                        ],
                        [
                            'status' => 'oof',
                            'start' => ['dateTime' => '2026-02-17T11:00:00'],
                            'end' => ['dateTime' => '2026-02-17T12:00:00'],
                        ],
                        [
                            'status' => 'workingElsewhere',
                            'start' => ['dateTime' => '2026-02-17T13:00:00'],
                            'end' => ['dateTime' => '2026-02-17T14:00:00'],
                        ],
                        [
                            'status' => 'free',
                            'start' => ['dateTime' => '2026-02-17T14:00:00'],
                            'end' => ['dateTime' => '2026-02-17T15:00:00'],
                        ],
                        [
                            'status' => 'unknown',
                            'start' => ['dateTime' => '2026-02-17T15:00:00'],
                            'end' => ['dateTime' => '2026-02-17T16:00:00'],
                        ],
                    ],
                ],
            ],
        ];

        $this->graphClient->method('post')
            ->willReturn($graphResponse);

        $start = new DateTimeImmutable('2026-02-17T08:00:00');
        $end = new DateTimeImmutable('2026-02-17T18:00:00');

        $results = $this->provider->getFreeBusy(
            'event-uri', 'calendar-uri',
            ['mailto:user@example.com'], $start, $end,
        );

        $vfreebusy = $results[0]->VFREEBUSY;
        $freebusyProps = $vfreebusy->select('FREEBUSY');

        // 4 entries: busy, tentative, oof, workingElsewhere (free and unknown skipped)
        $this->assertCount(4, $freebusyProps);

        $fbTypes = array_map(fn($fb) => (string)$fb['FBTYPE'], $freebusyProps);
        $this->assertContains('BUSY', $fbTypes);
        $this->assertContains('BUSY-TENTATIVE', $fbTypes);
        $this->assertContains('BUSY-UNAVAILABLE', $fbTypes);
    }

    public function testGetFreeBusyReturnsEmptyWhenNotConnected(): void {
        $tokenService = $this->createMock(TokenService::class);
        $tokenService->method('isConnected')->willReturn(false);

        $cacheFactory = $this->createMock(ICacheFactory::class);
        $cacheFactory->method('createDistributed')->willReturn($this->cache);

        $provider = new MsGraphFreeBusyProvider(
            $this->graphClient,
            $tokenService,
            $this->userSession,
            $cacheFactory,
            $this->logger,
        );

        $start = new DateTimeImmutable('2026-02-17T08:00:00');
        $end = new DateTimeImmutable('2026-02-17T18:00:00');

        $results = $provider->getFreeBusy(
            'event-uri', 'calendar-uri',
            ['mailto:user@example.com'], $start, $end,
        );

        $this->assertEmpty($results);
    }

    public function testGetFreeBusyReturnsEmptyForInvalidEmails(): void {
        $this->cache->method('get')->willReturn(null);

        $start = new DateTimeImmutable('2026-02-17T08:00:00');
        $end = new DateTimeImmutable('2026-02-17T18:00:00');

        $results = $this->provider->getFreeBusy(
            'event-uri', 'calendar-uri',
            ['not-an-email'], $start, $end,
        );

        $this->assertEmpty($results);
    }

    public function testGetFreeBusyStripsMailtoPrefix(): void {
        $this->cache->method('get')->willReturn(null);

        $this->graphClient->expects($this->once())
            ->method('post')
            ->willReturnCallback(function (string $userId, string $endpoint, array $body) {
                $this->assertContains('user@example.com', $body['schedules']);
                return ['value' => []];
            });

        $start = new DateTimeImmutable('2026-02-17T08:00:00');
        $end = new DateTimeImmutable('2026-02-17T18:00:00');

        $this->provider->getFreeBusy(
            'event-uri', 'calendar-uri',
            ['mailto:user@example.com'], $start, $end,
        );
    }
}
