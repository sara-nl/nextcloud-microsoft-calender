<template>
	<div id="ms365-admin-settings" class="section">
		<h2>Microsoft 365 Calendar Integration</h2>
		<p class="settings-hint">
			Configure the connection to Microsoft Entra ID (Azure AD).
			You need an App Registration in your Microsoft tenant.
		</p>

		<div class="ms365-form">
			<div class="ms365-form-row">
				<label for="ms365-tenant-id">Tenant ID</label>
				<input
					id="ms365-tenant-id"
					v-model="tenantId"
					type="text"
					placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
					@input="markDirty">
			</div>

			<div class="ms365-form-row">
				<label for="ms365-client-id">Client ID (Application ID)</label>
				<input
					id="ms365-client-id"
					v-model="clientId"
					type="text"
					placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
					@input="markDirty">
			</div>

			<div class="ms365-form-row">
				<label for="ms365-client-secret">Client Secret</label>
				<input
					id="ms365-client-secret"
					v-model="clientSecret"
					type="password"
					placeholder="Enter client secret"
					@input="markDirty">
				<p v-if="clientSecretSet && !dirty" class="hint">
					A client secret is currently configured.
				</p>
			</div>

			<div class="ms365-form-row">
				<label for="ms365-cache-ttl">Cache TTL (seconds)</label>
				<input
					id="ms365-cache-ttl"
					v-model.number="cacheTtl"
					type="number"
					min="60"
					max="3600"
					@input="markDirty">
				<p class="hint">How long search results are cached (default: 300s).</p>
			</div>

			<div class="ms365-form-actions">
				<button
					class="primary"
					:disabled="saving || !dirty"
					@click="save">
					{{ saving ? 'Saving...' : 'Save' }}
				</button>
				<span v-if="saved" class="ms365-saved-indicator">Settings saved</span>
				<span v-if="error" class="ms365-error-indicator">{{ error }}</span>
			</div>
		</div>

		<div class="ms365-info">
			<h3>Required API Permissions (Delegated)</h3>
			<ul>
				<li><strong>User.ReadBasic.All</strong> — Search users (requires admin consent)</li>
				<li><strong>People.Read</strong> — Relevant contacts</li>
				<li><strong>Calendars.Read</strong> — Free/busy information</li>
				<li><strong>Calendars.Read.Shared</strong> — Shared calendar free/busy</li>
				<li><strong>offline_access</strong> — Refresh tokens</li>
			</ul>

			<h3>Redirect URI</h3>
			<p>Set the following redirect URI in your App Registration:</p>
			<code>{{ redirectUri }}</code>
		</div>
	</div>
</template>

<script>
import axios from '@nextcloud/axios'
import { generateUrl } from '@nextcloud/router'

export default {
	name: 'AdminSettings',
	data() {
		return {
			tenantId: '',
			clientId: '',
			clientSecret: '',
			clientSecretSet: false,
			cacheTtl: 300,
			saving: false,
			saved: false,
			dirty: false,
			error: '',
			redirectUri: window.location.origin + generateUrl('/apps/nc_ms365_calendar/oauth/callback'),
		}
	},
	mounted() {
		this.load()
	},
	methods: {
		async load() {
			try {
				const url = generateUrl('/apps/nc_ms365_calendar/settings/admin')
				const response = await axios.get(url)
				this.tenantId = response.data.tenant_id || ''
				this.clientId = response.data.client_id || ''
				this.clientSecretSet = response.data.client_secret === '********'
				this.cacheTtl = parseInt(response.data.cache_ttl, 10) || 300
			} catch (e) {
				this.error = 'Failed to load settings'
			}
		},
		markDirty() {
			this.dirty = true
			this.saved = false
			this.error = ''
		},
		async save() {
			this.saving = true
			this.error = ''
			try {
				const url = generateUrl('/apps/nc_ms365_calendar/settings/admin')
				await axios.put(url, {
					tenant_id: this.tenantId,
					client_id: this.clientId,
					client_secret: this.clientSecret,
					cache_ttl: this.cacheTtl,
				})
				this.saved = true
				this.dirty = false
				this.clientSecretSet = this.clientSecret !== ''
				this.clientSecret = ''
			} catch (e) {
				this.error = 'Failed to save settings'
			} finally {
				this.saving = false
			}
		},
	},
}
</script>

<style scoped>
#ms365-admin-settings {
	max-width: 700px;
}

.ms365-form-row {
	margin-bottom: 16px;
}

.ms365-form-row label {
	display: block;
	font-weight: bold;
	margin-bottom: 4px;
}

.ms365-form-row input[type="text"],
.ms365-form-row input[type="password"],
.ms365-form-row input[type="number"] {
	width: 100%;
	max-width: 400px;
}

.ms365-form-row .hint {
	color: var(--color-text-lighter);
	font-size: 0.9em;
	margin-top: 4px;
}

.ms365-form-actions {
	margin-top: 20px;
	display: flex;
	align-items: center;
	gap: 12px;
}

.ms365-saved-indicator {
	color: var(--color-success);
}

.ms365-error-indicator {
	color: var(--color-error);
}

.ms365-info {
	margin-top: 32px;
	padding-top: 16px;
	border-top: 1px solid var(--color-border);
}

.ms365-info code {
	display: block;
	padding: 8px 12px;
	background: var(--color-background-dark);
	border-radius: var(--border-radius);
	word-break: break-all;
	margin-top: 4px;
}

.ms365-info ul {
	padding-left: 20px;
}

.ms365-info li {
	margin-bottom: 4px;
}
</style>
