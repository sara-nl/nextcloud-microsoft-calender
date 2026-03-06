<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Dav;

use DateTimeImmutable;
use DateTimeZone;
use OCA\NcMs365Calendar\AppInfo\Application;
use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\ICache;
use OCP\ICacheFactory;
use OCP\IUserManager;
use OCP\IUserSession;
use Psr\Log\LoggerInterface;
use Sabre\DAV\Server;
use Sabre\DAV\ServerPlugin;
use Sabre\HTTP\RequestInterface;
use Sabre\HTTP\ResponseInterface;
use Sabre\VObject\Component\VCalendar;
use Sabre\VObject\Reader;

/**
 * Sabre DAV plugin that intercepts CalDAV outbox free/busy requests
 * and provides schedule data for external MS365 attendees.
 *
 * Strategy:
 * 1. Intercept method:POST at priority 99 (before built-in Schedule Plugin at 100)
 * 2. If the request contains external attendees and user is MS365-connected:
 *    a. Temporarily remove ourselves and let the built-in plugin handle the full request
 *    b. Parse the built-in response (which returns 3.7 for external attendees)
 *    c. Query MS365 Graph API for the external attendees
 *    d. Merge the MS365 results into the response, replacing 3.7 entries
 */
class MS365FreeBusyPlugin extends ServerPlugin {
	private const STATUS_MAP = [
		'busy' => 'BUSY',
		'tentative' => 'BUSY-TENTATIVE',
		'oof' => 'BUSY-UNAVAILABLE',
		'workingElsewhere' => 'BUSY-TENTATIVE',
	];

	/** Max emails per Graph API getSchedule call */
	private const GRAPH_BATCH_SIZE = 20;

	/** Cache TTL for free/busy results (2 minutes) */
	private const CACHE_TTL = 120;

	private const CACHE_PREFIX = 'ms365:freebusy:';

	private ?Server $server = null;

	/** Lazy-loaded dependencies — only resolved on actual free/busy requests */
	private ?GraphApiClient $graphClient = null;
	private ?TokenService $tokenService = null;
	private ?IUserSession $userSession = null;
	private ?IUserManager $userManager = null;
	private ?ICache $cache = null;
	private ?LoggerInterface $logger = null;

	/**
	 * No constructor dependencies — keeps plugin instantiation zero-cost
	 * for the 99.9% of DAV requests that are not free/busy.
	 */
	public function __construct() {
	}

	/**
	 * Resolve heavy dependencies lazily on first actual use.
	 */
	private function resolveServices(): void {
		if ($this->graphClient !== null) {
			return;
		}
		$this->graphClient = \OCP\Server::get(GraphApiClient::class);
		$this->tokenService = \OCP\Server::get(TokenService::class);
		$this->userSession = \OCP\Server::get(IUserSession::class);
		$this->userManager = \OCP\Server::get(IUserManager::class);
		$this->logger = \OCP\Server::get(LoggerInterface::class);
		$cacheFactory = \OCP\Server::get(ICacheFactory::class);
		$this->cache = $cacheFactory->isAvailable()
			? $cacheFactory->createDistributed(Application::APP_ID)
			: $cacheFactory->createLocal(Application::APP_ID);
	}

	public function getPluginName(): string {
		return 'ms365-freebusy';
	}

	public function initialize(Server $server): void {
		$this->server = $server;
		// Priority 99 runs before the built-in Schedule Plugin (default 100)
		$server->on('method:POST', [$this, 'httpPost'], 99);
	}

	/**
	 * Intercept POST requests to CalDAV outbox for free/busy.
	 *
	 * @return bool|null false to stop event chain, null to let others run
	 */
	public function httpPost(RequestInterface $request, ResponseInterface $response): ?bool {
		$path = $request->getPath();
		if (!str_ends_with($path, '/outbox') && !str_ends_with($path, '/outbox/')) {
			return null;
		}

		// Read and preserve the body
		$body = $request->getBodyAsString();
		$request->setBody($body);

		if (empty($body)) {
			return null;
		}

		// Check Content-Type before expensive iCal parsing
		$contentType = $request->getHeader('Content-Type') ?? '';
		if (!str_contains($contentType, 'text/calendar') && !str_contains($contentType, 'application/calendar')) {
			return null;
		}

		// Parse iCalendar
		try {
			$vcalendar = Reader::read($body);
		} catch (\Exception $e) {
			return null;
		}

		if (!($vcalendar instanceof VCalendar)) {
			return null;
		}

		// Only handle VFREEBUSY REQUEST
		$vfreebusy = $vcalendar->VFREEBUSY;
		if ($vfreebusy === null) {
			return null;
		}

		if ((string)($vcalendar->METHOD ?? '') !== 'REQUEST') {
			return null;
		}

		// At this point we know it's a free/busy request — resolve dependencies
		$this->resolveServices();

		// Get current user
		$user = $this->userSession->getUser();
		if ($user === null) {
			return null;
		}
		$userId = $user->getUID();

		if (!$this->tokenService->isConnected($userId)) {
			return null;
		}

		// Extract attendees
		$attendees = [];
		if (isset($vfreebusy->ATTENDEE)) {
			foreach ($vfreebusy->ATTENDEE as $attendee) {
				$email = $this->extractEmail((string)$attendee);
				if ($email !== '') {
					$attendees[] = $email;
				}
			}
		}

		if (empty($attendees)) {
			return null;
		}

		// Identify external attendees (batch-friendly: single pass, results cached in NC UserManager)
		$externalAttendees = array_values(array_filter($attendees, function (string $email): bool {
			return empty($this->userManager->getByEmail($email));
		}));

		if (empty($externalAttendees)) {
			return null; // All local, let built-in handler deal with it
		}

		// Extract time range
		try {
			$start = $vfreebusy->DTSTART ? new DateTimeImmutable((string)$vfreebusy->DTSTART) : null;
			$end = $vfreebusy->DTEND ? new DateTimeImmutable((string)$vfreebusy->DTEND) : null;
		} catch (\Exception $e) {
			return null;
		}
		if ($start === null || $end === null) {
			return null;
		}

		$this->logger->debug('MS365FreeBusyPlugin: intercepting free/busy request', [
			'app' => Application::APP_ID,
			'external' => $externalAttendees,
			'totalAttendees' => count($attendees),
		]);

		// Step 1: Let the built-in Schedule Plugin handle the full request first
		// Remove ourselves to avoid recursion, then re-emit the event
		$this->server->removeListener('method:POST', [$this, 'httpPost']);

		// Reset body for the built-in handler
		$request->setBody($body);

		// Let the built-in Schedule Plugin process the request
		// It will write its response (with 3.7 for external attendees) to $response
		$this->server->emit('method:POST', [$request, $response]);

		// Re-register ourselves
		$this->server->on('method:POST', [$this, 'httpPost'], 99);

		// Step 2: Parse the built-in response
		$existingBody = $response->getBodyAsString();
		$existingResults = $this->parseScheduleResponse($existingBody);

		// Step 3: Query MS365 for external attendees
		$ms365Results = $this->getMs365FreeBusy($userId, $externalAttendees, $start, $end);

		// Step 4: Merge — replace 3.7 (not found) entries with MS365 data
		foreach ($ms365Results as $email => $result) {
			$existingResults[$email] = $result;
		}

		// Step 5: Write the merged response
		$this->buildResponse($response, $existingResults);

		// Return false — we've already handled the full response
		return false;
	}

	/**
	 * Query MS365 Graph API for free/busy data.
	 * Handles batching (max 20 per request) and caching (2 min TTL).
	 *
	 * @return array<string, array{status: string, calendar-data?: string}>
	 */
	private function getMs365FreeBusy(
		string $userId,
		array $emails,
		DateTimeImmutable $start,
		DateTimeImmutable $end,
	): array {
		$utcStart = $start->setTimezone(new DateTimeZone('UTC'));
		$utcEnd = $end->setTimezone(new DateTimeZone('UTC'));
		$timeKey = $utcStart->getTimestamp() . '-' . $utcEnd->getTimestamp();

		$results = [];
		$uncachedEmails = [];

		// Check cache per email
		foreach ($emails as $email) {
			$cacheKey = self::CACHE_PREFIX . $userId . ':' . md5($email . $timeKey);
			$cached = $this->cache?->get($cacheKey);
			if ($cached !== null) {
				$results[$email] = $cached;
			} else {
				$uncachedEmails[] = $email;
			}
		}

		if (empty($uncachedEmails)) {
			return $results;
		}

		// Batch uncached emails in groups of 20 (Graph API limit)
		$batches = array_chunk($uncachedEmails, self::GRAPH_BATCH_SIZE);

		foreach ($batches as $batch) {
			$batchResults = $this->fetchScheduleBatch($userId, $batch, $utcStart, $utcEnd);

			// Cache each result individually
			foreach ($batchResults as $email => $result) {
				$cacheKey = self::CACHE_PREFIX . $userId . ':' . md5($email . $timeKey);
				$this->cache?->set($cacheKey, $result, self::CACHE_TTL);
				$results[$email] = $result;
			}
		}

		return $results;
	}

	/**
	 * Fetch a single batch of schedules from Graph API (max 20 emails).
	 *
	 * @return array<string, array{status: string, calendar-data?: string}>
	 */
	private function fetchScheduleBatch(
		string $userId,
		array $emails,
		DateTimeImmutable $utcStart,
		DateTimeImmutable $utcEnd,
	): array {
		$results = [];

		$body = [
			'schedules' => array_values($emails),
			'startTime' => [
				'dateTime' => $utcStart->format('Y-m-d\\TH:i:s'),
				'timeZone' => 'UTC',
			],
			'endTime' => [
				'dateTime' => $utcEnd->format('Y-m-d\\TH:i:s'),
				'timeZone' => 'UTC',
			],
			'availabilityViewInterval' => 15,
		];

		$scheduleData = $this->graphClient->post($userId, '/me/calendar/getSchedule', $body);

		if ($scheduleData === null || !isset($scheduleData['value'])) {
			$this->logger->warning('MS365FreeBusyPlugin: Graph API returned no data', [
				'app' => Application::APP_ID,
			]);
			foreach ($emails as $email) {
				$results[$email] = [
					'status' => '2.0;Success',
					'calendar-data' => $this->buildEmptyFreeBusy($email, $utcStart, $utcEnd),
				];
			}
			return $results;
		}

		// Map results by email (case-insensitive)
		$scheduleByEmail = [];
		foreach ($scheduleData['value'] as $schedule) {
			$scheduleEmail = strtolower($schedule['scheduleId'] ?? '');
			if ($scheduleEmail !== '') {
				$scheduleByEmail[$scheduleEmail] = $schedule;
			}
		}

		foreach ($emails as $email) {
			$schedule = $scheduleByEmail[strtolower($email)] ?? null;

			if ($schedule === null) {
				$results[$email] = [
					'status' => '2.0;Success',
					'calendar-data' => $this->buildEmptyFreeBusy($email, $utcStart, $utcEnd),
				];
				continue;
			}

			$vcalendar = $this->buildFreeBusyFromSchedule($schedule, $email, $utcStart, $utcEnd);
			$results[$email] = [
				'status' => '2.0;Success',
				'calendar-data' => $vcalendar->serialize(),
			];
		}

		return $results;
	}

	/**
	 * Build a VFREEBUSY VCalendar from MS Graph schedule data.
	 */
	private function buildFreeBusyFromSchedule(
		array $schedule,
		string $email,
		DateTimeImmutable $start,
		DateTimeImmutable $end,
	): VCalendar {
		$vcalendar = new VCalendar();
		$vfreebusy = $vcalendar->createComponent('VFREEBUSY');

		$vfreebusy->DTSTART = $start->format('Ymd\\THis\\Z');
		$vfreebusy->DTEND = $end->format('Ymd\\THis\\Z');
		$vfreebusy->ORGANIZER = "mailto:{$email}";
		$vfreebusy->ATTENDEE = "mailto:{$email}";

		$hasItems = false;

		// Use scheduleItems for detailed busy periods
		if (!empty($schedule['scheduleItems']) && is_array($schedule['scheduleItems'])) {
			foreach ($schedule['scheduleItems'] as $item) {
				$status = $item['status'] ?? 'unknown';
				$fbType = self::STATUS_MAP[$status] ?? null;

				if ($fbType === null) {
					continue;
				}

				if (empty($item['start']['dateTime']) || empty($item['end']['dateTime'])) {
					continue;
				}
				try {
					$itemStart = new DateTimeImmutable($item['start']['dateTime'], new DateTimeZone('UTC'));
					$itemEnd = new DateTimeImmutable($item['end']['dateTime'], new DateTimeZone('UTC'));
				} catch (\Exception $e) {
					continue;
				}

				$freebusy = $vcalendar->createProperty(
					'FREEBUSY',
					$itemStart->format('Ymd\\THis\\Z') . '/' . $itemEnd->format('Ymd\\THis\\Z')
				);
				$freebusy['FBTYPE'] = $fbType;
				$vfreebusy->add($freebusy);
				$hasItems = true;
			}
		}

		// Fallback: parse availabilityView string
		if (!$hasItems && !empty($schedule['availabilityView'])) {
			$this->parseAvailabilityView($vcalendar, $vfreebusy, $schedule['availabilityView'], $start);
		}

		$vcalendar->add($vfreebusy);
		return $vcalendar;
	}

	/**
	 * Parse the availabilityView string into FREEBUSY periods.
	 * 0 = free, 1 = tentative, 2 = busy, 3 = oof, 4 = working elsewhere
	 */
	private function parseAvailabilityView(
		VCalendar $vcalendar,
		mixed $vfreebusy,
		string $view,
		DateTimeImmutable $rangeStart,
	): void {
		$intervalMinutes = 15;
		$viewMap = [
			'1' => 'BUSY-TENTATIVE',
			'2' => 'BUSY',
			'3' => 'BUSY-UNAVAILABLE',
			'4' => 'BUSY-TENTATIVE',
		];

		$len = strlen($view);
		$i = 0;

		while ($i < $len) {
			$char = $view[$i];
			$fbType = $viewMap[$char] ?? null;

			if ($fbType === null) {
				$i++;
				continue;
			}

			$blockStart = $i;
			while ($i < $len && $view[$i] === $char) {
				$i++;
			}

			$slotStart = $rangeStart->modify('+' . ($blockStart * $intervalMinutes) . ' minutes');
			$slotEnd = $rangeStart->modify('+' . ($i * $intervalMinutes) . ' minutes');

			$freebusy = $vcalendar->createProperty(
				'FREEBUSY',
				$slotStart->format('Ymd\\THis\\Z') . '/' . $slotEnd->format('Ymd\\THis\\Z')
			);
			$freebusy['FBTYPE'] = $fbType;
			$vfreebusy->add($freebusy);
		}
	}

	/**
	 * Build an empty VFREEBUSY (all free) for an email.
	 */
	private function buildEmptyFreeBusy(string $email, DateTimeImmutable $start, DateTimeImmutable $end): string {
		$vcalendar = new VCalendar();
		$vfreebusy = $vcalendar->createComponent('VFREEBUSY');
		$vfreebusy->DTSTART = $start->format('Ymd\\THis\\Z');
		$vfreebusy->DTEND = $end->format('Ymd\\THis\\Z');
		$vfreebusy->ORGANIZER = "mailto:{$email}";
		$vfreebusy->ATTENDEE = "mailto:{$email}";
		$vcalendar->add($vfreebusy);
		return $vcalendar->serialize();
	}

	/**
	 * Build the cal:schedule-response XML for the HTTP response.
	 *
	 * @param array<string, array{status: string, calendar-data?: string}> $results
	 */
	private function buildResponse(ResponseInterface $response, array $results): void {
		$dom = new \DOMDocument('1.0', 'utf-8');
		$dom->formatOutput = false;

		$scheduleResponse = $dom->createElement('cal:schedule-response');
		$scheduleResponse->setAttribute('xmlns:D', 'DAV:');
		$scheduleResponse->setAttribute('xmlns:cal', 'urn:ietf:params:xml:ns:caldav');
		$dom->appendChild($scheduleResponse);

		foreach ($results as $email => $result) {
			$xresponse = $dom->createElement('cal:response');

			$recipient = $dom->createElement('cal:recipient');
			$href = $dom->createElement('D:href');
			$href->appendChild($dom->createTextNode('mailto:' . $email));
			$recipient->appendChild($href);
			$xresponse->appendChild($recipient);

			$status = $dom->createElement('cal:request-status');
			$status->appendChild($dom->createTextNode($result['status']));
			$xresponse->appendChild($status);

			if (isset($result['calendar-data'])) {
				$calendarData = $dom->createElement('cal:calendar-data');
				// Normalize line endings to \n, matching Sabre's built-in behavior
				$calendarData->appendChild($dom->createTextNode(
					str_replace("\r\n", "\n", $result['calendar-data'])
				));
				$xresponse->appendChild($calendarData);
			}

			$scheduleResponse->appendChild($xresponse);
		}

		$xml = $dom->saveXML();

		$response->setStatus(200);
		$response->setHeader('Content-Type', 'application/xml');
		$response->setBody($xml);
	}

	/**
	 * Parse a cal:schedule-response XML body into our result format.
	 *
	 * @return array<string, array{status: string, calendar-data?: string}>
	 */
	private function parseScheduleResponse(string $xmlBody): array {
		$results = [];

		if (empty($xmlBody)) {
			return $results;
		}

		try {
			$dom = new \DOMDocument();
			$dom->loadXML($xmlBody, LIBXML_NONET | LIBXML_NOENT);
			$xpath = new \DOMXPath($dom);
			$xpath->registerNamespace('C', 'urn:ietf:params:xml:ns:caldav');
			$xpath->registerNamespace('D', 'DAV:');

			$responses = $xpath->query('//C:response');
			foreach ($responses as $resp) {
				$recipientNode = $xpath->query('.//D:href', $resp)->item(0);
				$statusNode = $xpath->query('.//C:request-status', $resp)->item(0);
				$calDataNode = $xpath->query('.//C:calendar-data', $resp)->item(0);

				if ($recipientNode === null) {
					continue;
				}

				$email = $this->extractEmail($recipientNode->textContent);
				$status = $statusNode ? $statusNode->textContent : '3.7;Could not find principal';

				$entry = ['status' => $status];
				if ($calDataNode !== null && $calDataNode->textContent !== '') {
					$entry['calendar-data'] = $calDataNode->textContent;
				}

				$results[$email] = $entry;
			}
		} catch (\Exception $e) {
			$this->logger->warning('Failed to parse schedule response: ' . $e->getMessage(), [
				'app' => Application::APP_ID,
			]);
		}

		return $results;
	}

	/**
	 * Extract email from a mailto: URI or plain email.
	 */
	private function extractEmail(string $address): string {
		$address = trim($address);
		if (str_starts_with($address, 'mailto:')) {
			$address = substr($address, 7);
		}
		return $address;
	}
}
