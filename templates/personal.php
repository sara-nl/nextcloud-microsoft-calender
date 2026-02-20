<?php

declare(strict_types=1);

use OCA\NcMs365Calendar\AppInfo\Application;

\OCP\Util::addScript(Application::APP_ID, 'nc_ms365_calendar-personal');
\OCP\Util::addStyle(Application::APP_ID, 'personal');

?>
<div id="ms365-calendar-personal-settings"></div>
