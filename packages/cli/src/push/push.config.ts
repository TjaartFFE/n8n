import { Config, Env } from '@n8n/config';

@Config
export class PushConfig {
	/** Backend to use for push notifications */
	@Env('N8N_PUSH_BACKEND')
	backend: 'sse' | 'websocket' = 'websocket';

	/**
	 * Comma-separated list of hostnames explicitly allowed to connect to the push endpoint,
	 * bypassing the proxy-header origin check. Use this when your reverse proxy does not
	 * forward `X-Forwarded-Host` / `Forwarded` headers correctly.
	 *
	 * @example N8N_ALLOWED_HOSTS=ffe.app.n8n.cloud,other.example.com
	 */
	@Env('N8N_ALLOWED_HOSTS')
	allowedHosts: string = '';
}
