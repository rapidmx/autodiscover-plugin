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
