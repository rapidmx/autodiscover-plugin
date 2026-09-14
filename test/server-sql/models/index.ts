// Re-exports just the model Autodiscover needs so the test Server's ClassLoader (rooted at `test/server-sql`) can
// discover its `@DataStore` metadata. A named, not wildcard, re-export: `@rapidmx/restapi/sql` also exports every
// route and job, which this harness doesn't configure dependencies for.
export { MailboxSQL } from "@rapidmx/restapi/sql";
