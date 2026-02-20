<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\AddressBook;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCA\NcMs365Calendar\Service\GraphApiClient;
use OCA\NcMs365Calendar\Service\TokenService;
use OCP\IAddressBook;
use OCP\ICache;
use OCP\ICacheFactory;
use OCP\IUserSession;
use Psr\Log\LoggerInterface;

class MsGraphAddressBook implements IAddressBook {
    private const CACHE_PREFIX = 'ms365:search:';
    private const CACHE_TTL = 300; // 5 minutes default
    private const PHOTO_CACHE_PREFIX = 'ms365:photo:';
    private const PHOTO_CACHE_TTL = 86400; // 24 hours

    private ?ICache $cache;

    public function __construct(
        private GraphApiClient $graphClient,
        private TokenService $tokenService,
        private IUserSession $userSession,
        private ICacheFactory $cacheFactory,
        private LoggerInterface $logger,
    ) {
        $this->cache = $cacheFactory->createDistributed(Application::APP_ID);
    }

    public function getKey(): string {
        return 'ms365-addressbook';
    }

    public function getUri(): string {
        return 'ms365-addressbook';
    }

    public function getDisplayName(): string {
        return 'Microsoft 365';
    }

    /**
     * Search for MS365 users matching the pattern.
     *
     * @param string $pattern Search pattern
     * @param array $searchProperties Properties to search in (ignored, we always search name/email)
     * @param array $options Additional options (limit, offset)
     * @return array Array of contacts in NC format
     */
    public function search($pattern, $searchProperties, $options): array {
        if (strlen($pattern) < 2) {
            return [];
        }

        $user = $this->userSession->getUser();
        if ($user === null) {
            return [];
        }
        $userId = $user->getUID();

        if (!$this->tokenService->isConnected($userId)) {
            return [];
        }

        // Check cache
        $cacheKey = self::CACHE_PREFIX . $userId . ':' . md5($pattern);
        $cached = $this->cache?->get($cacheKey);
        if ($cached !== null) {
            return $cached;
        }

        // Fetch from both People API and Users API, then merge
        $contacts = $this->searchAndMerge($userId, $pattern);

        // Apply limit if specified
        $limit = $options['limit'] ?? 25;
        if (count($contacts) > $limit) {
            $contacts = array_slice($contacts, 0, $limit);
        }

        // Cache results
        $this->cache?->set($cacheKey, $contacts, self::CACHE_TTL);

        return $contacts;
    }

    public function createOrUpdate($properties): array {
        // Read-only address book
        return [];
    }

    public function getPermissions(): int {
        return \OCP\Constants::PERMISSION_READ;
    }

    public function delete($id): bool {
        // Read-only address book
        return false;
    }

    public function isShared(): bool {
        return false;
    }

    public function isSystemAddressBook(): bool {
        return true;
    }

    /**
     * Search People API, Users API ($search) and Users API ($filter), merge and deduplicate.
     */
    private function searchAndMerge(string $userId, string $pattern): array {
        $contacts = [];
        $seenEmails = [];

        // 1. People API - returns relevant contacts based on communication frequency
        // Note: People API uses 'scoredEmailAddresses' internally; omit $select to get full response
        $peopleResult = $this->graphClient->get($userId, '/me/people', [
            '$search' => "\"$pattern\"",
            '$top' => '10',
        ]);

        if ($peopleResult !== null && isset($peopleResult['value'])) {
            foreach ($peopleResult['value'] as $person) {
                $email = $this->extractEmail($person);
                if ($email === '' || isset($seenEmails[$email])) {
                    continue;
                }
                $seenEmails[$email] = true;
                $contacts[] = $this->mapToContact($userId, $person, $email);
            }
        }

        // 2. Users API - $search by displayName (separate call, OR is unreliable)
        $usersResult = $this->graphClient->get($userId, '/users', [
            '$search' => "\"displayName:{$pattern}\"",
            '$select' => 'displayName,mail,department,id',
            '$top' => '10',
            '$count' => 'true',
        ]);

        if ($usersResult !== null && isset($usersResult['value'])) {
            foreach ($usersResult['value'] as $user) {
                $email = $user['mail'] ?? '';
                if ($email === '' || isset($seenEmails[$email])) {
                    continue;
                }
                $seenEmails[$email] = true;
                $contacts[] = $this->mapUserToContact($userId, $user);
            }
        }

        // 3. Users API - $search by mail
        $mailResult = $this->graphClient->get($userId, '/users', [
            '$search' => "\"mail:{$pattern}\"",
            '$select' => 'displayName,mail,department,id',
            '$top' => '10',
            '$count' => 'true',
        ]);

        if ($mailResult !== null && isset($mailResult['value'])) {
            foreach ($mailResult['value'] as $user) {
                $email = $user['mail'] ?? '';
                if ($email === '' || isset($seenEmails[$email])) {
                    continue;
                }
                $seenEmails[$email] = true;
                $contacts[] = $this->mapUserToContact($userId, $user);
            }
        }

        // 4. Fallback: $filter with startsWith when $search returns nothing
        if (empty($contacts)) {
            $escaped = str_replace("'", "''", $pattern);
            $filterResult = $this->graphClient->get($userId, '/users', [
                '$filter' => "startsWith(displayName,'{$escaped}') or startsWith(mail,'{$escaped}')",
                '$select' => 'displayName,mail,department,id',
                '$top' => '10',
            ]);

            if ($filterResult !== null && isset($filterResult['value'])) {
                foreach ($filterResult['value'] as $user) {
                    $email = $user['mail'] ?? '';
                    if ($email === '' || isset($seenEmails[$email])) {
                        continue;
                    }
                    $seenEmails[$email] = true;
                    $contacts[] = $this->mapUserToContact($userId, $user);
                }
            }
        }

        return $contacts;
    }

    private function extractEmail(array $person): string {
        if (!empty($person['emailAddresses'])) {
            foreach ($person['emailAddresses'] as $emailEntry) {
                if (!empty($emailEntry['address'])) {
                    return $emailEntry['address'];
                }
            }
        }
        return '';
    }

    /**
     * Map a People API result to NC contact format.
     */
    private function mapToContact(string $userId, array $person, string $email): array {
        $contact = [
            'FN' => $person['displayName'] ?? $email,
            'EMAIL' => [$email],
            'ORG' => $person['department'] ?? '',
            'X-NC-SCOPE' => 'ms365',
            'isLocalSystemBook' => true,
        ];

        $photo = $this->getPhoto($userId, $person['id'] ?? '');
        if ($photo !== null) {
            $contact['PHOTO'] = $photo;
        }

        return $contact;
    }

    /**
     * Map a Users API result to NC contact format.
     */
    private function mapUserToContact(string $userId, array $user): array {
        $contact = [
            'FN' => $user['displayName'] ?? $user['mail'] ?? '',
            'EMAIL' => [$user['mail']],
            'ORG' => $user['department'] ?? '',
            'X-NC-SCOPE' => 'ms365',
            'isLocalSystemBook' => true,
        ];

        $photo = $this->getPhoto($userId, $user['id'] ?? '');
        if ($photo !== null) {
            $contact['PHOTO'] = $photo;
        }

        return $contact;
    }

    /**
     * Get profile photo with caching.
     */
    private function getPhoto(string $userId, string $msUserId): ?string {
        if ($msUserId === '') {
            return null;
        }

        $cacheKey = self::PHOTO_CACHE_PREFIX . $msUserId;
        $cached = $this->cache?->get($cacheKey);
        if ($cached !== null) {
            return $cached === '' ? null : $cached;
        }

        $photo = $this->graphClient->getProfilePhoto($userId, $msUserId);

        // Cache even empty result to avoid repeated failed requests
        $this->cache?->set($cacheKey, $photo ?? '', self::PHOTO_CACHE_TTL);

        return $photo;
    }
}
