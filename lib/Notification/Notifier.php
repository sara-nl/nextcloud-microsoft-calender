<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Notification;

use OCA\NcMs365Calendar\AppInfo\Application;
use OCP\IURLGenerator;
use OCP\L10N\IFactory;
use OCP\Notification\INotification;
use OCP\Notification\INotifier;
use OCP\Notification\UnknownNotificationException;

class Notifier implements INotifier {
    public function __construct(
        private IFactory $l10nFactory,
        private IURLGenerator $urlGenerator,
    ) {
    }

    public function getID(): string {
        return Application::APP_ID;
    }

    public function getName(): string {
        return $this->l10nFactory->get(Application::APP_ID)->t('MS365 Calendar');
    }

    public function prepare(INotification $notification, string $languageCode): INotification {
        if ($notification->getApp() !== Application::APP_ID) {
            throw new UnknownNotificationException('Unknown app');
        }

        $l = $this->l10nFactory->get(Application::APP_ID, $languageCode);
        $params = $notification->getSubjectParameters();

        $attendee = $params['attendee'] ?? '?';
        $event = $params['event'] ?? '?';

        // Determine subject text based on type
        $subjectText = match ($notification->getSubject()) {
            'invitation_accepted' => $l->t('{attendee} accepted your invitation'),
            'invitation_declined' => $l->t('{attendee} declined your invitation'),
            'invitation_tentative' => $l->t('{attendee} tentatively accepted your invitation'),
            default => throw new UnknownNotificationException('Unknown subject'),
        };

        // Rich subject with highlighted attendee name
        $notification->setRichSubject($subjectText, [
            'attendee' => [
                'type' => 'highlight',
                'id' => $attendee,
                'name' => $attendee,
            ],
        ]);

        $notification->setParsedSubject(
            str_replace('{attendee}', $attendee, $subjectText)
        );

        // Message body with event name and date/time
        $messageLines = [$l->t('Event: %s', [$event])];

        if (!empty($params['dtstart'])) {
            try {
                $start = new \DateTime($params['dtstart']);
                $dateStr = $start->format('D j M Y, H:i');
                if (!empty($params['dtend'])) {
                    $end = new \DateTime($params['dtend']);
                    if ($start->format('Y-m-d') === $end->format('Y-m-d')) {
                        $dateStr .= ' – ' . $end->format('H:i');
                    } else {
                        $dateStr .= ' – ' . $end->format('D j M Y, H:i');
                    }
                }
                $messageLines[] = $l->t('When: %s', [$dateStr]);
            } catch (\Exception) {
                // skip date if unparseable
            }
        }

        $notification->setParsedMessage(implode("\n", $messageLines));
        $notification->setRichMessage(implode("\n", $messageLines));

        // Link to calendar event — /apps/calendar/edit/{base64(davUrl)}
        if (!empty($params['calendarUri']) && !empty($params['objectUri'])) {
            $davUrl = '/remote.php/dav/calendars/' . $notification->getUser()
                . '/' . $params['calendarUri']
                . '/' . $params['objectUri'];
            $objectId = base64_encode($davUrl);
            $link = $this->urlGenerator->getAbsoluteURL(
                '/index.php/apps/calendar/edit/' . $objectId
            );
            $notification->setLink($link);
        }

        $notification->setIcon(
            $this->urlGenerator->getAbsoluteURL(
                $this->urlGenerator->imagePath('core', 'places/calendar.svg')
            )
        );

        return $notification;
    }
}
