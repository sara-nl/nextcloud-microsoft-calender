<?php

declare(strict_types=1);

namespace OCA\NcMs365Calendar\Listener;

use OCA\NcMs365Calendar\Dav\MS365FreeBusyPlugin;
use OCA\DAV\Events\SabrePluginAddEvent;
use OCP\EventDispatcher\Event;
use OCP\EventDispatcher\IEventListener;

/**
 * Listens for SabrePluginAddEvent to register our MS365 free/busy Sabre plugin.
 * The plugin itself is zero-cost: no constructor dependencies, lazy service resolution.
 *
 * @template-implements IEventListener<SabrePluginAddEvent>
 */
class FreeBusyPluginListener implements IEventListener {
	public function handle(Event $event): void {
		if (!($event instanceof SabrePluginAddEvent)) {
			return;
		}

		$event->getServer()->addPlugin(new MS365FreeBusyPlugin());
	}
}
