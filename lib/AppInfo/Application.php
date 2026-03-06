<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\AppInfo;

use OCP\Calendar\Events\CalendarObjectUpdatedEvent;
use OCA\DAV\Events\SabrePluginAddEvent;
use OCA\NcMs365Calendar\AddressBook\MsGraphAddressBook;
use OCA\NcMs365Calendar\Listener\CalendarReplyListener;
use OCA\NcMs365Calendar\Listener\FreeBusyPluginListener;
use OCA\NcMs365Calendar\Notification\Notifier;
use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IRegistrationContext;
use OCP\Contacts\IManager as IContactsManager;

class Application extends App implements IBootstrap {
    public const APP_ID = 'nc_ms365_calendar';

    public function __construct() {
        parent::__construct(self::APP_ID);
    }

    public function register(IRegistrationContext $ctx): void {
        // Register Sabre DAV plugin for MS365 free/busy
        $ctx->registerEventListener(
            SabrePluginAddEvent::class,
            FreeBusyPluginListener::class,
        );

        // Calendar reply notifications (PARTSTAT changes)
        $ctx->registerEventListener(
            CalendarObjectUpdatedEvent::class,
            CalendarReplyListener::class,
        );

        // Notification formatter
        $ctx->registerNotifierService(Notifier::class);
    }

    public function boot(IBootContext $ctx): void {
        $server = $ctx->getServerContainer();

        // Register address book lazily so it only loads when contacts are searched
        $server->get(IContactsManager::class)->register(function () use ($server) {
            $server->get(IContactsManager::class)
                ->registerAddressBook(
                    $server->get(MsGraphAddressBook::class)
                );
        });
    }
}
