# n8n

n8n runs as a standalone Docker container. It does not use the legacy
`compose.yaml`, Traefik, Watchtower, or Renovate.

Traffic follows the same pattern as `cargo.mleczki.pl`:

```text
Cloudflare (proxied A record) -> Cytrus -> [Mikrus IPv6]:8082 -> n8n
```

The provision workflow expects these GitHub environment secrets:

- `MIKRUS_IPV6`
- `CYTRUS_IPV4`
- `CYTRUS_API_TOKEN`
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ZONE_ID`
- `N8N_ENCRYPTION_KEY`

Before the first deployment, add `n8n.mleczki.pl` to Cytrus. The playbook
verifies that it is present and manages the proxied Cloudflare A record.

## Storage and backups

The complete n8n state is bind-mounted at `/var/lib/n8n`. Weekly maintenance
uses SQLite's online-backup command and `PRAGMA integrity_check` before it
checks for a newer `docker.n8n.io/n8nio/n8n:stable` image. Verified pre-update
copies are retained in `/var/lib/n8n-backups/pre-update` for 14 days.

For off-host backups, configure the private instance created from
`mleczakm/backup` with a `files` target like the following. Store this JSON in
the instance repository secret `BAKTIME_TARGET_N8N`, and store the SSH key in
`BAKTIME_TARGET_N8N_SSH_KEY`:

```json
{
  "type": "files",
  "host": "wanda192.mikrus.xyz",
  "sshUser": "root",
  "sshPort": 10192,
  "sshKeySecretName": "BAKTIME_TARGET_N8N_SSH_KEY",
  "paths": ["/var/lib/n8n", "/var/lib/n8n-backups/pre-update"],
  "excludes": ["database.sqlite", "database.sqlite-*"],
  "sqliteBackups": [
    {
      "source": "/var/lib/n8n/database.sqlite",
      "destination": "/var/lib/n8n-backups/baktime/database.sqlite"
    }
  ],
  "schedule": "0 3 * * 1"
}
```

This uses the backup repository's SQLite flow: online backup, integrity check,
restic snapshot, and removal of the temporary staging copy.
