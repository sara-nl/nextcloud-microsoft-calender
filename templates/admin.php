<?php

declare(strict_types=1);

use OCA\NcMs365Calendar\AppInfo\Application;

\OCP\Util::addScript(Application::APP_ID, 'nc_ms365_calendar-admin');
\OCP\Util::addStyle(Application::APP_ID, 'admin');

?>
<div id="ms365-calendar-admin-settings"></div>
