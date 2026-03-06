# Microsoft 365 Calendar Integration for Nextcloud

Integrate Microsoft 365 users into the Nextcloud Calendar with attendee search and free/busy availability.

## Features

- **Attendee Search** — Search Microsoft 365 users when adding attendees to calendar events
- **Free/Busy** — View availability of MS365 users directly in Nextcloud Calendar
- **Reply Notifications** — Get notified when attendees accept, decline or tentatively accept invitations
- **OAuth2 with PKCE** — Secure per-user authentication with Microsoft Entra ID

## Requirements

- Nextcloud 28–33
- PHP 8.2+
- An App Registration in Microsoft Entra ID (Azure AD)

## Installation

1. Clone or download this repository into your Nextcloud `apps/` directory:
   ```bash
   cd /var/www/nextcloud/apps
   git clone https://github.com/sara-nl/nextcloud-microsoft-calender.git nc_ms365_calendar
   ```

2. Install dependencies and build the frontend:
   ```bash
   cd nc_ms365_calendar
   composer install --no-dev
   npm install
   npm run build
   ```

3. Enable the app:
   ```bash
   occ app:enable nc_ms365_calendar
   ```

## Microsoft Entra ID Setup

### 1. Create an App Registration

- Go to [Azure Portal](https://portal.azure.com) → **App registrations** → **New registration**
- Name: e.g. "Nextcloud Calendar Integration"
- Supported account types: Single tenant (your organization)
- Redirect URI: **Web** → `https://your-nextcloud.com/apps/nc_ms365_calendar/oauth/callback`

### 2. Configure API Permissions

Add the following **Delegated** permissions under **Microsoft Graph**:

| Permission | Purpose | Admin Consent |
|---|---|---|
| `User.ReadBasic.All` | Search users | Required |
| `People.Read` | Relevant contacts | No |
| `Calendars.Read` | Free/busy information | No |
| `Calendars.Read.Shared` | Shared calendar free/busy | No |
| `offline_access` | Refresh tokens | No |

Click **Grant admin consent** after adding permissions.

### 3. Create a Client Secret

- Go to **Certificates & secrets** → **New client secret**
- Copy the secret value (it's only shown once)

### 4. Note Your IDs

You'll need:
- **Tenant ID** — found on the Overview page of your App Registration
- **Client ID** (Application ID) — found on the Overview page
- **Client Secret** — from step 3

## Nextcloud Configuration

### Admin Settings

Go to **Administration settings** → **Microsoft 365 Calendar Integration** and enter:

- **Tenant ID**
- **Client ID**
- **Client Secret**
- **Cache TTL** — How long search results are cached (default: 300s)
- **Reply Notifications** — Send notifications when attendees respond to invitations (toggle on/off)

### User Connection

Each user connects their own Microsoft 365 account:

1. Go to **Personal settings** → **Microsoft 365 Calendar**
2. Click **Connect to Microsoft 365**
3. Sign in with your Microsoft account and grant permissions

## How It Works

### Attendee Search

When a user searches for attendees in the Calendar app, the app queries the Microsoft Graph API using multiple strategies:

1. **People API** — Finds frequently contacted people
2. **Users API** — Searches by display name and email
3. Results are merged, deduplicated, and cached (5-minute TTL)

### Free/Busy

The app registers a Sabre DAV plugin that intercepts CalDAV free/busy requests:

1. Identifies external attendees (not in Nextcloud)
2. Queries Microsoft Graph `/me/calendar/getSchedule`
3. Merges MS365 availability with Nextcloud's response
4. Works transparently with any CalDAV client (Apple Calendar, Thunderbird, etc.)

Status mapping:

| MS365 Status | CalDAV Status |
|---|---|
| `busy` | `BUSY` |
| `tentative` | `BUSY-TENTATIVE` |
| `oof` (Out of Office) | `BUSY-UNAVAILABLE` |
| `workingElsewhere` | `BUSY-TENTATIVE` |

### Reply Notifications

When enabled in admin settings, the app monitors calendar event updates for PARTSTAT changes (accept/decline/tentative) and sends Nextcloud notifications to the event organizer.

## Development

```bash
# Install dependencies
composer install
npm install

# Build frontend
npm run build

# Watch mode
npm run dev

# Run tests
./vendor/bin/phpunit
```

## License

AGPL-3.0-or-later
