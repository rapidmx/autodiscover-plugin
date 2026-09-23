# autodiscover — Design Decisions & Session Notes

This file exists so that Claude sessions working in this repo don't re-litigate settled
decisions or re-discover the same issues from scratch. It is local to this repo (not tied to
any one machine's global Claude memory), so it travels with the code.

**Maintenance rule:** when a standing decision changes, update the section below in place
(don't just append a contradiction lower down). When a new investigation/session produces a
decision, finding, or reverted approach worth remembering, add a dated entry under Session Log.
Keep entries terse — this is a reference, not a transcript.

## Standing design decisions & constraints

- **Vulnerability/review threat model: externally-exploitable only.** Only count issues reachable
  from a downstream, untrusted HTTP client hitting a service built on this package (anonymous or
  low-privilege caller). Do NOT flag developer-only footguns or purely theoretical races with no
  concrete external trigger path.
- **Commit discipline.** Don't `git commit` unless explicitly asked for *that specific piece of
  work*. An autonomous-execution/"commit as you go" approval given for one approved plan (e.g. via
  plan mode) is scoped to that plan only — it does not carry forward to later, separate requests in
  the same session, even ones that look similar in kind (a follow-up review-and-fix pass, a
  refactor, a new feature), and even after a full review-and-fix cycle with passing tests. Default
  to leaving changes staged/unstaged and saying so; only commit automatically within the exact
  scope of a plan that was explicitly approved as autonomous. If unsure whether new work falls
  inside that scope, treat it as outside and ask.
- **Commit message style: a flat list of one-line, verb-led items — no summary/title line, no
  `-`/`*` bullet markers.** This isn't just a style preference — it's dictated by how `release`
  (`@rapidrest/cli`) actually builds `CHANGELOG.md`. `collectChangelogBullets`/
  `classifyChangelogLine` (that repo's `src/lib/release.ts`) parse `git log --pretty=format:%B` and
  treat **every non-blank line of a commit's full message as its own changelog bullet** — there is
  no subject/body distinction. A conventional "short imperative subject + blank line + prose body"
  commit therefore leaks one changelog bullet per body sentence, and a `-`/`*`-prefixed line breaks
  `classifyChangelogLine`'s verb detection (it reads the line's first whitespace-delimited word as
  the verb; a leading `-` defeats that lookup and the dash leaks into the changelog text as
  `"- - Added foo"`). Correct format:
  - No separate summary/title line — if a commit needs an overview, that overview is itself just
    one more flat line, not a heading distinct from the rest.
  - No bullet-marker prefix of any kind — write bare lines.
  - Lead each line with an imperative verb where it fits: `Add`/`Fix`/`Remove` (and `-ing` forms)
    are recognized and become `Added`/`Fixed`/`Removed` entries; `Configuring`/`Converting`/
    `Refactoring`/`Updating`/etc. become `Changed`. Anything else still works, defaulting to
    `Changed` verbatim — see `CHANGELOG_VERB_REWRITES` in that repo's `src/lib/release.ts` for the
    full map.
  - A blank line before a trailing git trailer (`Co-Authored-By:`, `Signed-off-by:`, etc.) is fine
    — trailers matching `CHANGELOG_NOISE_PATTERNS` are dropped from the changelog — but nothing
    else should follow the item list.
  This mirrors JP's standing convention across his other repos; copy this exact rule verbatim into
  each sibling repo's own NOTES.md rather than paraphrasing it, since the paraphrase is what caused
  this to be gotten wrong in the first place (see `@rapidrest/cli`'s own NOTES.md, 2026-09-07 entry,
  for the full incident writeup and the `CHANGELOG_NOISE_PATTERNS` fix that accompanied it).

## Session Log

### 2026-09-06 — Repo split: `@rapidrest/mail` → four RapidMX packages

- **This repo is `@rapidmx/autodiscover`**, carved out of the former monolith `@rapidrest/mail`
  (`d:\github\rapidrest\mail`, still present there for reference/history). It's a **new 4th
  package** not originally in the split plan — added because Autodiscover has zero code coupling
  to Exchange ActiveSync or MAPI over HTTP (it only ever touches the `Mailbox` model and builds
  response strings), so keeping it as its own package avoids duplicating it into both
  `@rapidmx/activesync` and `@rapidmx/mapi`.
- Depends on [`@rapidmx/restapi`](https://github.com/RapidMX/restapi) for the `Mailbox`
  model/repo and the `BlobStore`/`SearchProvider`/`SpamScanProvider`/`AvScanProvider`/
  `MailTransport` interfaces its test harness registers test-doubles for (even though this
  package's own routes never touch most of those — the shared integration-test `Server` fixture
  eagerly instantiates every route it discovers, which is why `testDoubles.ts` still registers all
  of them; see that file's own doc comment). Originally linked locally via Yarn Berry's
  `portal:../restapi` during development; switched to the real published `^0.1.0` once JP published
  it to npm (see the `activesync` repo's own notes for exactly why the portal approach turned out to
  be a real problem, not just a temporary convenience - it silently duplicated the
  `@rapidrest/service-core` module instance at test-runtime).
- **`vitest.config.ts`'s `ssr.noExternal` must list `@rapidmx/restapi`** alongside
  `@rapidrest/service-core`/`@rapidrest/core` - without it, Vite's SSR pipeline can load a second,
  natively-required copy of the framework packages for anything reached *through* `@rapidmx/restapi`,
  breaking static/instanceof-based state shared with the test's own directly-imported copy (see
  `activesync`'s NOTES.md for the full diagnosis - this repo's own narrower test surface didn't
  happen to trigger the symptom, but the same risk applies here too).
- **Mechanical migration gotcha, worth remembering for the sibling `activesync`/`mapi` splits too**:
  several distinct old import targets (`blob/BlobStore.js`, `transport/MailTransport.js`,
  `models/types.js`) all collapse onto the same new specifier (`@rapidmx/restapi`) once rewritten,
  which turns a mechanical per-line import rewrite into multiple separate `import ... from
  "@rapidmx/restapi"` statements — `no-duplicate-imports` correctly flags that. Fixed in
  `test/testDoubles.ts` by hand-merging into one statement (using inline `type` specifier modifiers
  to mix value and type-only imports in the same line, since `AvVerdict`/`SpamVerdict` are real
  values but the rest are types-only).
- For the original design rationale behind the Outlook/EXCH response shape (confirmed against the
  real `[MS-OXDSCLI]` spec, not assumed) and the MobileSync/EAS response shape, see the monolith's
  own `.claude/NOTES.md` (`d:\github\rapidrest\mail`) — that history wasn't duplicated here since
  it predates this repo's existence and mostly concerns code that never lived under this package's
  own directory.

### 2026-09-14 — Review-finding fix pass (POX ReDoS, public_url validation, v2 double-decode)

Each finding was confirmed in code first. Not committed; no version/peerDependency/manifest changes (the
manifest `requires` on ActiveSync and MAPI stays, per JP).

- **POX ReDoS (unauthenticated).** `EMAIL_ELEMENT_PATTERN`/`ACCEPTABLE_RESPONSE_SCHEMA_PATTERN`
  (`\s*([^<]*?)\s*<\/...`) backtracked cubically: `<EMailAddress>` plus 4000 spaces and no close tag took
  about 13.5s in Node. Both replaced by one linear forward tag scan (`extractElementText` in
  `AutodiscoverXml.ts`). It matches the local name exactly with an optional `ns:` prefix, case-insensitive.
  The old regex also accepted any `[\w:]*` prefix such as `FooEMailAddress`. Self-closing elements are skipped.
  `pox()` also returns `413` for a body over `MAX_POX_BODY_BYTES` (16KB) before scanning.
- **`public_url` validation.** A private `baseUrl` getter trims the value, parses it with `URL`, and requires
  `https:` (`http:` only for localhost/127.0.0.1/[::1]). It rejects credentials, `?` and `#` (checked on the raw
  string, since `URL.search` is empty for a bare trailing `?`) and joins origin+path without a trailing slash.
  An invalid value logs a warning at `@Init` and advertises nothing.
- **v2 double decode.** service-core's uWS/Bun routers already `decodeURIComponent` path params, so
  `v2()`'s own second decode threw on a literal `%` and returned 500. Removed.
- README package names updated to the `-plugin` names.
- Tests: linear-time and edge-case tests in `AutodiscoverXml.test.ts`; URL-validation and 413 unit tests in
  `test/routes/BaseAutodiscoverRoute.test.ts`; pathological-body timing, 413, and `%` address HTTP tests in
  `test/routes/mongo/AutodiscoverRoute.test.ts`.

### 2026-09-14 (2) — Round-2 review fixes (directory enumeration, display-name disclosure, XML scan gaps)

Each finding was confirmed in code first. Not committed; no version, peerDependency or manifest changes
(`requires` stays, per JP).

- **HIGH: unauthenticated directory enumeration.** `resolveMailbox()` passed the caller's email straight into
  `find({ primarySmtpAddress })` and the alias query. service-core's query parser reads `like(*)`, `regex(^a)`,
  `in(a,b)` and `ne(x)` in a value, so POX or v2 could confirm "some mailbox matches". A new private
  `normalizeAddress()` trims and lower-cases the value. It then requires a single plain address
  (`/^[^\s@(),"\\]+@[^\s@(),"\\]+$/`) of at most 254 characters (`MAX_ADDRESS_LENGTH`). Anything else gets `400`
  from POX or `404 UserNotFound` from v2, and no query runs. Lookups now use a literal `eq(<address>)` with
  `limit: 1` in both the query and the options, mirroring restapi's `BaseMailboxAccessRoute`. The base
  `aliasQueryValue()` returns `eq(...)`. The SQL override (escaped `LIKE` on the `simple-json` column) is
  unchanged; `"` and `\` are excluded by the pattern, so they can't break out of its `%"..."%` quoting.
- **MEDIUM: POX disclosed `Mailbox.displayName` to anonymous callers**, contrary to the class doc. Both POX
  builders now get the caller's own (normalized) address as `DisplayName`. That keeps the Outlook schema's
  `DisplayName` populated, and the MobileSync element is still valid. The response `EMailAddress` is now the
  lower-cased address.
- **LOW: tag-scan gaps in `AutodiscoverXml.extractElementText`.** It now allows whitespace before a closing
  tag's `>`, skips `<!-- -->` comments (inside the element, and around it so a commented-out element never
  matches) and unwraps `<![CDATA[...]]>` verbatim. Entities are decoded per text segment, never inside CDATA.
  An unterminated comment or CDATA returns `undefined`. It is still one forward scan; every `indexOf` starts
  past the previous one.
- Tests: the literal `eq()` query and no-query-for-operators unit test is in `test/routes/BaseAutodiscoverRoute.test.ts`.
  Operator-value 400/404 HTTP tests and the display-name non-disclosure assertions are in both
  `test/routes/{mongo,sql}/AutodiscoverRoute.test.ts`. Closing-tag whitespace, comment and CDATA tests are in
  `AutodiscoverXml.test.ts`.

### 2026-09-14 (3) — `@rapidrest/service-core` 2.1.0 migration

Not committed; package version unchanged. `@rapidmx/restapi` left at `^0.9.0`.

- **Deps.** service-core devDependency `^2.0.0` -> `^2.1.0`, peerDependency `2.x` -> `^2.1.0`; `yarn install`
  installed 2.1.0.
- **Suite before any code change: 70/70 passing** (6 files), 100% coverage. No 2.1.0 breaking change affected this
  plugin (it only reads `Mailbox`; no creates, updates, truncates, date columns, `$or`, Redis or `@RateLimit`), and no
  restapi 0.9.0 code path it reaches failed.
- **Simplification.** `resolveMailbox()`'s `primarySmtpAddress` query and the base `aliasQueryValue()` now use
  `ModelUtils.literal(address)` instead of a hand-built `eq(${address})` string. Behaviour-identical: the address has
  already passed `normalizeAddress()` (one `@`, no whitespace/parens/commas/quotes/backslashes), so it could never be
  `me`, `null` or a number that `eq()` would have substituted or coerced. The SQL `aliasQueryValue()` override (escaped
  `LIKE` via `Raw`) is unchanged. `test/routes/BaseAutodiscoverRoute.test.ts` now expects `ModelUtils.literal(...)`.
- Final: `yarn lint` and `npx tsc --noEmit -p .` clean; `yarn vitest run --coverage` 70/70, 100% statements/branches/
  functions/lines.

### 2026-09-22 - Autodiscover never actually worked: the setting's help text didn't say what DNS was still needed, and a stale doc comment

Two small fixes alongside restapi/server/web-client's half of this (see restapi's own NOTES for the full picture - the DNS checklist gaining
`autodiscover_cname`/`autodiscover_srv`, and server finally getting a default `autodiscover.public_url` config block). Here: expanded this
plugin's own `mail:autodiscover:public_url` admin-console setting's `help` text to say what DNS is also needed (a CNAME, or a SRV record to
avoid a second certificate) and point at the Domain DNS setup checklist; added its `"default": ""`. Rewrote `src/index.ts`'s stale doc-comment
example, which still described an old subclass-override mounting pattern (`protected readonly easUrl = "..."`) `BaseAutodiscoverRoute` hasn't
used in some time - it builds both endpoint URLs itself from the config setting, so a deployment just loads `AutodiscoverRouteMongo`/
`AutodiscoverRouteSQL` directly. `yarn test:prod` clean: 70/70 tests, 100/100/100/100.

### 2026-09-22 (2) - Round-3 review fixes (rate limiting, XML tag-scan quote handling)

Each finding was confirmed in code first. Not committed as a version bump - see the commit itself for what
landed; `RELEASE_NOTES.md` gained an `Unreleased` section per the usual convention.

- **HIGH: unauthenticated mailbox-existence oracle, no rate limiting.** Both `pox()` and `v2()` answer
  anonymously with an observably different outcome (404/`UserNotFound` vs. success) depending on whether the
  requested address matches a real `Mailbox`, and nothing throttled repeated calls. Both are now
  `@RateLimit()`-decorated (`@rapidrest/service-core`'s `RouteDecorators`), the same bare usage
  `BaseKeyDiscoveryRoute`'s own public key lookup uses in restapi. `@RateLimit()` with no options defaults
  `perUser: true`; for an anonymous caller that scopes the identifier-keyed counter to the client's source IP
  (`ip:<address>|<method>|<path>`), and the framework's own always-on secondary per-IP layer applies on top of
  that whenever `req` is available - so this is IP-scoped in two independent ways, not merely the class-level
  `<method>|<path>` default some other undecorated-option routes rely on.
  - **Test gotcha worth remembering:** since neither route has a path parameter (unlike `BaseKeyDiscoveryRoute`'s
    `/:hash`, where each test's distinct hash value naturally lands in its own counter bucket), every request
    to `pox()` across an entire test file shares one identifier bucket (supertest always calls from the same
    loopback address). A rate-limit test that just lowers `RateLimiter.config.maxAttempts` without resetting
    state will immediately 429 on its very first request, because earlier tests in the same file already
    consumed most of the default budget. Fixed by calling `rateLimiter.memoryStore.clear()` (this `RateLimiter`
    instance's own dedicated in-memory counter store - confirmed by inspecting its keys mid-test, only
    `ratelimit:*` entries for this plugin's two routes were ever present, nothing else) immediately before
    installing the temporary low-`maxAttempts` config, in each of the 4 new rate-limit tests
    (`test/routes/{mongo,sql}/AutodiscoverRoute.test.ts`, one per protocol per backend).
- **LOW: XML tag scanner mis-extracted element text when a sibling attribute value contained a literal `>`.**
  `AutodiscoverXml.ts`'s `extractElementText()` found the opening tag's closing `>` via a plain
  `xml.indexOf(">", pos)`, which stops early on an unescaped `>` inside a quoted attribute value (e.g.
  `<EMailAddress xmlns:x="a>b">real@example.com</EMailAddress>`) - previously failed closed (the corrupted
  string contains a `"`, which `BaseAutodiscoverRoute`'s `PLAIN_ADDRESS_PATTERN` rejects, producing a 400), not
  exploitable, but a genuine parsing bug. Fixed by a new `findTagClose()` helper that tracks single-/
  double-quote state and only reports a `>` seen outside any quote; still one forward, linear-time scan (same
  ReDoS-safety property as the rest of this module - verified against the existing pathological-input timing
  test, which still passes unchanged).
- Tests: quoted-`>`-in-attribute regression cases (double-quoted, single-quoted, and a self-closing variant) in
  `test/AutodiscoverXml.test.ts`; 429-after-N-anonymous-calls tests for both `pox()` and `v2()` in both
  `test/routes/mongo/AutodiscoverRoute.test.ts` and `test/routes/sql/AutodiscoverRoute.test.ts`.
- Final: `yarn lint`, `npx tsc --noEmit -p .`, and `yarn vitest run --coverage` all clean - 75/75 tests,
  100/100/100/100.
