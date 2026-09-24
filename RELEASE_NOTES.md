# Release Notes

## Unreleased

## v1.1.0

### Fixes

* Rate limit the anonymous POX (`autodiscover.xml`) and Autodiscover v2 (`autodiscover.json/v1.0/:email`)
  endpoints, both of which reveal via their HTTP status whether an email address belongs to a real mailbox -
  without this, either endpoint could be hammered to enumerate real mailboxes at a domain
* Fixed the POX request XML tag scanner mis-extracting an element's text when a sibling attribute value on its
  opening tag contains a literal `>` (e.g. `xmlns:x="a>b"`); it now tracks quote state so that inner `>` no
  longer ends the tag early
* Closed a gap in the Autodiscover v2 endpoint's rate limiting where the limit was keyed per queried email
  address (taken from the URL), so an anonymous caller could enumerate any number of candidate addresses from
  one source with no endpoint-specific throttling; it's now a fixed, shared per-endpoint budget regardless of
  which or how many addresses are queried
* Changed the POX success response's `Content-Type` from `application/xml` to `text/xml`, matching real
  Exchange Autodiscover responses - some older/strict mobile mail clients hard-check the exact MIME type

## v1.0.0

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
