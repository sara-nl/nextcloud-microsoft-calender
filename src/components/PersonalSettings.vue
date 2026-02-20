<template>
	<div id="ms365-personal-settings" class="section">
		<h2>Microsoft 365 Calendar</h2>
		<p class="settings-hint">
			Connect your Microsoft 365 account to search colleagues as attendees
			and view their availability in the Calendar.
		</p>

		<div v-if="loading" class="ms365-loading">
			Loading...
		</div>

		<div v-else-if="connected" class="ms365-connected">
			<div class="ms365-status">
				<span class="ms365-status-icon ms365-status-connected" />
				<strong>Connected</strong>
			</div>

			<div class="ms365-details">
				<p><strong>Email:</strong> {{ email }}</p>
				<p v-if="connectedAt">
					<strong>Connected since:</strong> {{ connectedAtFormatted }}
				</p>
				<p v-if="tokenExpired" class="ms365-warning">
					Your token has expired. Please reconnect your account.
				</p>
			</div>

			<div class="ms365-actions">
				<button
					v-if="tokenExpired"
					class="primary"
					@click="connect">
					Reconnect Microsoft Account
				</button>
				<button
					class="error"
					:disabled="disconnecting"
					@click="disconnect">
					{{ disconnecting ? 'Disconnecting...' : 'Disconnect' }}
				</button>
			</div>
		</div>

		<div v-else class="ms365-disconnected">
			<div class="ms365-status">
				<span class="ms365-status-icon ms365-status-disconnected" />
				<strong>Not connected</strong>
			</div>
			<p>Connect your Microsoft 365 account to enable attendee search and free/busy information.</p>

			<button
				class="primary"
				@click="connect">
				Connect Microsoft Account
			</button>
		</div>

		<p v-if="error" class="ms365-error">{{ error }}</p>
	</div>
</template>

<script>
import axios from '@nextcloud/axios'
import { generateUrl } from '@nextcloud/router'

export default {
	name: 'PersonalSettings',
	data() {
		return {
			loading: true,
			connected: false,
			email: '',
			connectedAt: 0,
			tokenExpired: false,
			disconnecting: false,
			error: '',
		}
	},
	computed: {
		connectedAtFormatted() {
			if (!this.connectedAt) {
				return ''
			}
			return new Date(this.connectedAt * 1000).toLocaleString()
		},
	},
	mounted() {
		this.loadStatus()
	},
	methods: {
		async loadStatus() {
			this.loading = true
			try {
				const url = generateUrl('/apps/nc_ms365_calendar/oauth/status')
				const response = await axios.get(url)
				this.connected = response.data.connected
				this.email = response.data.email || ''
				this.connectedAt = response.data.connectedAt || 0
				this.tokenExpired = response.data.tokenExpired || false
			} catch (e) {
				this.error = 'Failed to load connection status'
			} finally {
				this.loading = false
			}
		},
		connect() {
			const url = generateUrl('/apps/nc_ms365_calendar/oauth/authorize')
			window.location.href = url
		},
		async disconnect() {
			this.disconnecting = true
			this.error = ''
			try {
				const url = generateUrl('/apps/nc_ms365_calendar/oauth/disconnect')
				await axios.post(url)
				this.connected = false
				this.email = ''
				this.connectedAt = 0
				this.tokenExpired = false
			} catch (e) {
				this.error = 'Failed to disconnect'
			} finally {
				this.disconnecting = false
			}
		},
	},
}
</script>

<style scoped>
#ms365-personal-settings {
	max-width: 700px;
}

.ms365-status {
	display: flex;
	align-items: center;
	gap: 8px;
	margin-bottom: 12px;
}

.ms365-status-icon {
	display: inline-block;
	width: 12px;
	height: 12px;
	border-radius: 50%;
}

.ms365-status-connected {
	background: var(--color-success);
}

.ms365-status-disconnected {
	background: var(--color-text-lighter);
}

.ms365-details {
	margin-bottom: 16px;
}

.ms365-details p {
	margin: 4px 0;
}

.ms365-actions {
	display: flex;
	gap: 12px;
	margin-top: 16px;
}

.ms365-warning {
	color: var(--color-warning);
	font-weight: bold;
}

.ms365-error {
	color: var(--color-error);
	margin-top: 12px;
}

.ms365-loading {
	color: var(--color-text-lighter);
}
</style>
