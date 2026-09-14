# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/RapidMX/autodiscover/compare/v1.0.0-beta.1...HEAD
[1.0.0-beta.1]: https://github.com/RapidMX/autodiscover/compare/v1.0.0-beta.0...v1.0.0-beta.1
[1.0.0-beta.0]: https://github.com/RapidMX/autodiscover/releases/tag/v1.0.0-beta.0
