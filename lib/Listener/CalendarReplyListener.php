<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Listener;

use OCP\Calendar\Events\CalendarObjectUpdatedEvent;
use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\EventDispatcher\Event;
use OCP\EventDispatcher\IEventListener;
use OCP\IAppConfig;
use OCP\Notification\IManager as INotificationManager;
use Psr\Log\LoggerInterface;
use Sabre\VObject\Reader;

/**
 * Detects PARTSTAT changes on calendar events and sends NC notifications
 * to the organizer so they know when attendees accept/decline.
 *
 * @template-implements IEventListener<CalendarObjectUpdatedEvent>
 */
class CalendarReplyListener implements IEventListener {
    public function __construct(
        private INotificationManager $notificationManager,
        private IAppConfig $appConfig,
        private LoggerInterface $logger,
    ) {
    }

    public function handle(Event $event): void {
        if (!($event instanceof CalendarObjectUpdatedEvent)) {
            return;
        }

        if (!$this->appConfig->getValueBool(Application::APP_ID, 'reply_notifications_enabled', false)) {
            return;
        }

        $objectRow = $event->getObjectData();
        $calendarData = $objectRow['calendardata'] ?? null;
        if ($calendarData === null) {
            return;
        }

        // Determine the calendar owner (NC user ID)
        $calendarRow = $event->getCalendarData();
        $principalUri = $calendarRow['principaluri'] ?? '';
        // principaluri is like "principals/users/admin"
        $ownerUid = basename($principalUri);
        if ($ownerUid === '' || str_contains($principalUri, 'calendar-rooms') || str_contains($principalUri, 'calendar-resources')) {
            return;
        }

        try {
            $vcalendar = Reader::read($calendarData);
        } catch (\Exception $e) {
            return;
        }

        $vevent = $vcalendar->VEVENT ?? null;
        if ($vevent === null) {
            return;
        }

        $summary = (string)($vevent->SUMMARY ?? 'Untitled event');

        // Extract event date/time
        $dtStart = '';
        $dtEnd = '';
        if (isset($vevent->DTSTART)) {
            $start = $vevent->DTSTART->getDateTime();
            $dtStart = $start->format('c');
            if (isset($vevent->DTEND)) {
                $dtEnd = $vevent->DTEND->getDateTime()->format('c');
            }
        }

        // Build link to calendar app
        $calendarUri = $calendarRow['uri'] ?? '';
        $objectUri = $objectRow['uri'] ?? '';

        // Check all attendees for PARTSTAT values that indicate a reply
        if (!isset($vevent->ATTENDEE)) {
            return;
        }

        foreach ($vevent->ATTENDEE as $attendee) {
            $partstat = isset($attendee['PARTSTAT'])
                ? strtoupper((string)$attendee['PARTSTAT'])
                : 'NEEDS-ACTION';

            // Only notify for actual responses
            if (!in_array($partstat, ['ACCEPTED', 'DECLINED', 'TENTATIVE'], true)) {
                continue;
            }

            $attendeeName = isset($attendee['CN'])
                ? (string)$attendee['CN']
                : str_replace('mailto:', '', (string)$attendee->getValue());

            $subject = match ($partstat) {
                'ACCEPTED' => 'invitation_accepted',
                'DECLINED' => 'invitation_declined',
                'TENTATIVE' => 'invitation_tentative',
            };

            // Use event UID + attendee email as stable dedup key (independent of partstat)
            $eventUid = (string)($vevent->UID ?? $objectRow['uri'] ?? 'unknown');
            $objectId = md5($eventUid . $attendee->getValue());

            try {
                // Remove any previous notification for this attendee+event combo
                $existing = $this->notificationManager->createNotification();
                $existing
                    ->setApp(Application::APP_ID)
                    ->setUser($ownerUid)
                    ->setObject('reply', $objectId);
                $this->notificationManager->markProcessed($existing);

                // Send fresh notification with current status
                $notification = $this->notificationManager->createNotification();
                $notification
                    ->setApp(Application::APP_ID)
                    ->setUser($ownerUid)
                    ->setDateTime(new \DateTime())
                    ->setObject('reply', $objectId)
                    ->setSubject($subject, [
                        'attendee' => $attendeeName,
                        'event' => $summary,
                        'dtstart' => $dtStart,
                        'dtend' => $dtEnd,
                        'calendarUri' => $calendarUri,
                        'objectUri' => $objectUri,
                    ]);

                $this->notificationManager->notify($notification);
            } catch (\Exception $e) {
                $this->logger->warning('Failed to send reply notification: ' . $e->getMessage(), [
                    'app' => Application::APP_ID,
                ]);
            }
        }
    }
}
