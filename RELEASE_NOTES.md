# Release Notes

## v1.0.0-beta.2

### Features

* Classic POX Autodiscover (`POST /autodiscover/autodiscover.xml`) per `[MS-ASCMD]`'s MobileSync request/
  response schema, resolving a client-supplied `EMailAddress` to a `Mailbox` and returning this deployment's
  Exchange ActiveSync server URL
* A real Outlook desktop response shape per `[MS-OXDSCLI]`'s Outlook/EXCH `2006a` schema, served automatically
  instead of the MobileSync response whenever a request's `AcceptableResponseSchema` asks for it — points a
  real Outlook client at this deployment's MAPI/HTTP endpoint instead of EAS
* Autodiscover v2 (JSON) (`GET /autodiscover/autodiscover.json/v1.0/:email?Protocol=ActiveSync`) for modern
  clients that skip the POX request/response XML entirely
* Mailbox resolution against both `Mailbox.primarySmtpAddress` and `Mailbox.aliasAddresses`, so an alias
  address discovers the same server URLs as the mailbox's primary address
* Deliberately unauthenticated by design, matching Autodiscover v2's own spec intent — reveals only
  deployment-wide server URLs (never a per-mailbox secret) once the requested address is confirmed to belong
  to a real mailbox; real mailbox access is still fully gated by the JWT-protected EAS/MAPI/REST layers
  downstream
* MongoDB and SQL persistence backends (`@rapidmx/autodiscover/mongo`, `@rapidmx/autodiscover/sql`), matching
  whichever backend the rest of a `@rapidmx/restapi`-based deployment already uses

### Fixes

* The "Public server URL" setting's help text now says what DNS an administrator also needs to add - a
  `CNAME` (`autodiscover.<domain>` to this host) or, to avoid a second TLS certificate, a `_autodiscover._tcp.<domain>`
  SRV record - and points at the deployment's Domain DNS setup checklist, which now recommends both
* Rewrote this package's own `src/index.ts` doc comment, which still described an old subclass-override
  mounting pattern (`protected readonly easUrl = "..."`) that `BaseAutodiscoverRoute` no longer uses - a
  deployment now just loads `AutodiscoverRouteMongo`/`AutodiscoverRouteSQL` directly and sets
  `mail:autodiscover:public_url`
