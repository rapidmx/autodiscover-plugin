# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-09-22

### Changed
- Expand the public server URL setting's help text to say what DNS is also needed, a CNAME or a SRV record that avoids a second certificate, pointing at the Domain DNS setup checklist that now recommends both
- Document the fix in the release notes and NOTES
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed a stale doc comment describing a subclass-override mounting pattern this plugin no longer uses

## [1.0.0-beta.2] - 2026-09-15

### Changed
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Require @rapidmx/activesync-plugin ^1.0.0-beta.2 and @rapidmx/mapi-plugin ^1.0.0-beta.3 in the plugin manifest, so they're installed and loaded before Autodiscover
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Validate mail:autodiscover:public_url, requiring https (http only for localhost) and no credentials, query or fragment; advertise nothing when it's invalid
- Stop double-decoding the Autodiscover v2 email path parameter, which returned 500 for addresses containing %
- Update the README to the @rapidmx/autodiscover-plugin package name
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Validate Autodiscover email addresses as a single plain address and look mailboxes up literally with eq(), so anonymous callers can't enumerate mailboxes with like(), regex() or in()
- Return the requested address as the POX DisplayName instead of the mailbox's display name
- Accept whitespace before a closing tag's >, skip comments and unwrap CDATA when reading the POX request
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Updated @rapidrest/service-core to ^2.1.0 as both the dev dependency and the peer range
- Use ModelUtils.literal() for mailbox address lookups instead of eq() query strings
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Upgarding restapi dep

### Fixed
- Fixed Autodiscover never advertising ActiveSync or MAPI by checking for the renamed @rapidmx/activesync-plugin and @rapidmx/mapi-plugin packages
- Fixed a ReDoS in POX Autodiscover: replace the EMailAddress and AcceptableResponseSchema regexes with a linear tag scan, and reject request bodies over 16KB with 413
- Fixed peer dep range for service-core

## [1.0.0-beta.1] - 2026-09-14

### Added
- Added tests for the plugin entry points, the manifest and protocol gating

### Changed
- Convert this library into a RapidMX server plugin: package.json carries a rapidmx.plugin manifest, and the ./mongo and ./sql entry points export AutodiscoverRouteMongo/AutodiscoverRouteSQL mounted at /autodiscover
- Change the advertised ActiveSync and MAPI URLs to come from the mail:autodiscover:public_url setting instead of abstract subclass fields, and only advertise a protocol while its plugin is loaded, answering 404 (POX) or ProtocolNotSupported (v2) otherwise
- Updated the peer and dev dependencies to @rapidmx/restapi 0.8 and @rapidrest/service-core 2
- Change the test harnesses to re-export only MailboxMongo/MailboxSQL rather than every restapi route and job
- Patch @rapidmx/restapi 0.8.0 with its unreleased plugin contract until the next restapi release
- Co-Authored-By: Claude Opus 5 (1M context) <noreply@anthropic.com>
- Upgraded deps
- Changing package name to @rapid/autodiscover-plugin

### Removed
- Removed @rapidrest/cli as a dep

## [1.0.0-beta.0] - 2026-09-09

### Added
- Added changelog, contributing guide, contributors

### Changed
- Initial commit
- Upgraded all dependencies
- Updated CI workflows
- Updated claude commit instructions
- Register a DnsResolver test double so DomainVerificationJob starts cleanly
- restapi's unreleased next version adds a pluggable DnsResolver interface
- (Domain TXT-record verification), injected by DomainVerificationJob, which
- the shared test/server-mongo and test/server-sql fixture apps boot
- unconditionally alongside every other background service. Without a
- registered double, that job failed to start on every integration test run.
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

[Unreleased]: https://github.com/RapidMX/autodiscover/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/RapidMX/autodiscover/compare/v1.0.0-beta.2...v1.0.0
[1.0.0-beta.2]: https://github.com/RapidMX/autodiscover/compare/v1.0.0-beta.1...v1.0.0-beta.2
[1.0.0-beta.1]: https://github.com/RapidMX/autodiscover/compare/v1.0.0-beta.0...v1.0.0-beta.1
[1.0.0-beta.0]: https://github.com/RapidMX/autodiscover/releases/tag/v1.0.0-beta.0
