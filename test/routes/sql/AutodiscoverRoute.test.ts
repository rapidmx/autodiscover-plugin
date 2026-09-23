///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// See the identical file header in test/routes/mongo/AutodiscoverRoute.test.ts for the full rationale - this
// verifies the same BaseAutodiscoverRoute behavior on the SQL-backed variant, including the
// AutodiscoverRouteSQL.aliasQueryValue() escaping override (test/routes/mongo's alias-match test doesn't
// exercise that SQL-specific `simple-json`-column code path at all).
import config from "../../config.sql.js";
import { request } from "@rapidrest/service-core/test";
import { Server, ObjectFactory, ConnectionManager, isSqlDataSource, RateLimiter } from "@rapidrest/service-core";
import { Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { Repository } from "typeorm";
import { MailboxSQL } from "@rapidmx/restapi/sql";
import { registerTestDoubles } from "../../testDoubles.js";
import { PluginRegistry } from "@rapidmx/restapi";

describe("Route:AutodiscoverRouteSQL Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-sql", logger, objectFactory });
    const baseUrl = "/sql/autodiscover";
    let mailboxRepo: Repository<MailboxSQL>;

    const createMailbox = async function (data?: Partial<MailboxSQL>): Promise<MailboxSQL> {
        const obj: MailboxSQL = new MailboxSQL({
            ownerUserUid: uuid.v4(),
            primarySmtpAddress: `${uuid.v4()}@example.com`,
            aliasAddresses: [],
            displayName: "Test Mailbox",
            timezone: "UTC",
            quotaBytes: 1_000_000_000,
            usedBytes: 0,
            ...data,
        });
        return await mailboxRepo.save(obj);
    };

    const poxRequestBody = function (email: string): string {
        return `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="https://schemas.microsoft.com/exchange/autodiscover/mobilesync/requestschema/2006">
    <Request>
        <EMailAddress>${email}</EMailAddress>
        <AcceptableResponseSchema>https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006</AcceptableResponseSchema>
    </Request>
</Autodiscover>`;
    };

    const outlookPoxRequestBody = function (email: string): string {
        return `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
    <Request>
        <EMailAddress>${email}</EMailAddress>
        <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
    </Request>
</Autodiscover>`;
    };

    beforeAll(async () => {
        registerTestDoubles(objectFactory);
        await server.start();
        PluginRegistry.setLoaded([
            { name: "@rapidmx/activesync-plugin", version: "1.0.0" },
            { name: "@rapidmx/mapi-plugin", version: "1.0.0" },
        ]);

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("sql");
        if (isSqlDataSource(conn)) {
            mailboxRepo = conn.getRepository(MailboxSQL);
        } else {
            throw new Error("Could not find sql connection");
        }
    });

    afterAll(async () => {
        await server.stop();
        await objectFactory.destroy();
    });

    beforeEach(async () => {
        await mailboxRepo.clear();
    });

    describe("POST .../autodiscover.xml (POX)", () => {
        it("Returns a MobileSync settings response for a known primarySmtpAddress.", async () => {
            const mailbox = await createMailbox({ displayName: "Ada Lovelace" });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress));

            expect(result.status).toBe(200);
            expect(result.text).toContain(`<autodiscover:EMailAddress>${mailbox.primarySmtpAddress}</autodiscover:EMailAddress>`);
            expect(result.text).toContain(`<autodiscover:DisplayName>${mailbox.primarySmtpAddress}</autodiscover:DisplayName>`);
            expect(result.text).not.toContain("Ada Lovelace");
        });

        it("Matches against the serialized aliasAddresses column via the escaped LIKE override.", async () => {
            const alias = `${uuid.v4()}@example.com`;
            await createMailbox({ aliasAddresses: [alias] });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(alias));

            expect(result.status).toBe(200);
            expect(result.text).toContain(`<autodiscover:EMailAddress>${alias}</autodiscover:EMailAddress>`);
        });

        it("Does not false-positive-match a substring of a stored alias.", async () => {
            await createMailbox({ aliasAddresses: [`bob@example.com`] });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody("ob@example.co"));

            expect(result.status).toBe(404);
        });

        it("Treats '%'/'_' in the request address as literal characters, not SQL LIKE wildcards.", async () => {
            // See the identical rationale in test/routes/sql/MailIngestRoute.test.ts - without escaping, a `_`
            // (SQL "match any one character" wildcard) in the request address would let "b_b@example.com"
            // falsely match a stored alias "bob@example.com".
            await createMailbox({ aliasAddresses: ["bob@example.com"] });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody("b_b@example.com"));

            expect(result.status).toBe(404);
        });

        it("Returns 400, never a directory match, for query-operator address values.", async () => {
            const mailbox = await createMailbox({ aliasAddresses: ["alias@example.com"] });
            for (const value of ["like(*)", "regex(^a)", `in(${mailbox.primarySmtpAddress},alias@example.com)`]) {
                const result = await request(server.getApplication())
                    .post(`${baseUrl}/autodiscover.xml`)
                    .set("Content-Type", "text/xml")
                    .send(poxRequestBody(value));
                expect(result.status).toBe(400);
            }
        });

        it("Returns 404 for an address with no matching Mailbox.", async () => {
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody("nobody@example.com"));
            expect(result.status).toBe(404);
        });

        it("Returns an Outlook/EXCH mapiHttp response when AcceptableResponseSchema requests the Outlook schema.", async () => {
            const mailbox = await createMailbox({ displayName: "Ada Lovelace" });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(outlookPoxRequestBody(mailbox.primarySmtpAddress));

            expect(result.status).toBe(200);
            const xml = result.text;
            expect(xml).toContain("http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a");
            expect(xml).toContain(`<AutoDiscoverSMTPAddress>${mailbox.primarySmtpAddress}</AutoDiscoverSMTPAddress>`);
            expect(xml).toContain('<Protocol Type="mapiHttp" Version="1">');
            expect(xml).toContain("<InternalUrl>https://mail.example.com/mapi/emsmdb</InternalUrl>");
        });

        it("Rate limits repeated anonymous POX lookups (429) - this endpoint is an unauthenticated mailbox-existence oracle.", async () => {
            const rateLimiter: any = objectFactory.getInstance(RateLimiter);
            const original = rateLimiter.config;
            // `@RateLimit()` scopes anonymous callers by client IP, so every prior request in this file (all
            // from supertest's own loopback address) already shares this counter - reset it so the assertions
            // below start from a known, empty count rather than whatever the earlier tests left behind.
            rateLimiter.memoryStore.clear();
            rateLimiter.config = { enabled: true, maxAttempts: 2, windowSeconds: 300, ip: { enabled: false } };
            try {
                const mailbox = await createMailbox();
                const send = () =>
                    request(server.getApplication())
                        .post(`${baseUrl}/autodiscover.xml`)
                        .set("Content-Type", "text/xml")
                        .send(poxRequestBody(mailbox.primarySmtpAddress));

                expect((await send()).status).toBe(200);
                expect((await send()).status).toBe(200);
                expect((await send()).status).toBe(429);
            } finally {
                rateLimiter.config = original;
            }
        });

        it("Shares one rate-limit bucket across different queried addresses (429) - the address is in the POST body, never the URL, so it can never fragment the bucket.", async () => {
            const rateLimiter: any = objectFactory.getInstance(RateLimiter);
            const original = rateLimiter.config;
            rateLimiter.memoryStore.clear();
            rateLimiter.config = { enabled: true, maxAttempts: 2, windowSeconds: 300, ip: { enabled: false } };
            try {
                const mailboxA = await createMailbox();
                const mailboxB = await createMailbox();
                const lookup = (address: string) =>
                    request(server.getApplication())
                        .post(`${baseUrl}/autodiscover.xml`)
                        .set("Content-Type", "text/xml")
                        .send(poxRequestBody(address));

                expect((await lookup(mailboxA.primarySmtpAddress)).status).toBe(200);
                expect((await lookup(mailboxB.primarySmtpAddress)).status).toBe(200);
                // A third request, for yet another, still-unqueried address, is still throttled: the two
                // addresses above already exhausted the one shared budget.
                expect((await lookup("someone-else@example.com")).status).toBe(429);
            } finally {
                rateLimiter.config = original;
            }
        });
    });

    describe("protocols whose plugins aren't loaded", () => {
        afterEach(() => {
            PluginRegistry.setLoaded([
                { name: "@rapidmx/activesync-plugin", version: "1.0.0" },
                { name: "@rapidmx/mapi-plugin", version: "1.0.0" },
            ]);
        });

        it("Returns 404 from POX and 400 from v2 when ActiveSync isn't loaded.", async () => {
            PluginRegistry.setLoaded([{ name: "@rapidmx/mapi-plugin", version: "1.0.0" }]);
            const mailbox = await createMailbox();
            const pox = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress));
            expect(pox.status).toBe(404);
            const v2 = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}?Protocol=ActiveSync`,
            );
            expect(v2.status).toBe(400);
            expect(v2.body.ErrorCode).toBe("ProtocolNotSupported");
        });

        it("Returns 404 to an Outlook client when MAPI isn't loaded.", async () => {
            PluginRegistry.setLoaded([{ name: "@rapidmx/activesync-plugin", version: "1.0.0" }]);
            const mailbox = await createMailbox();
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(outlookPoxRequestBody(mailbox.primarySmtpAddress));
            expect(result.status).toBe(404);
        });
    });

    describe("GET .../autodiscover.json/v1.0/:email (v2)", () => {
        it("Returns the EAS URL for a known address.", async () => {
            const mailbox = await createMailbox();

            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}?Protocol=ActiveSync`,
            );

            expect(result.status).toBe(200);
            expect(result.body).toEqual({
                Protocol: "ActiveSync",
                Url: "https://mail.example.com/Microsoft-Server-ActiveSync",
            });
        });

        it("Returns 404 JSON for an address with no matching Mailbox.", async () => {
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/nobody@example.com?Protocol=ActiveSync`,
            );
            expect(result.status).toBe(404);
            expect(result.body.ErrorCode).toBe("UserNotFound");
        });

        it("Returns 404 JSON, never a directory match, for query-operator or non-plain address values.", async () => {
            const mailbox = await createMailbox({ aliasAddresses: ["alias@example.com"] });
            for (const value of ["like(*)", "regex(^a)", `in(${mailbox.primarySmtpAddress},alias@example.com)`, `"alias@example.com"`]) {
                const result = await request(server.getApplication()).get(
                    `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(value)}?Protocol=ActiveSync`,
                );
                expect(result.status).toBe(404);
                expect(result.body.ErrorCode).toBe("UserNotFound");
            }
        });

        it("Returns 400 JSON for an unsupported Protocol value.", async () => {
            const mailbox = await createMailbox();
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}?Protocol=EWS`,
            );
            expect(result.status).toBe(400);
            expect(result.body.ErrorCode).toBe("ProtocolNotSupported");
        });

        it("Rate limits repeated anonymous v2 lookups (429) - this endpoint is an unauthenticated mailbox-existence oracle.", async () => {
            const rateLimiter: any = objectFactory.getInstance(RateLimiter);
            const original = rateLimiter.config;
            // See the identical comment in the POX rate-limit test above - resets the counter every prior
            // request in this file (all from the same loopback address) already shares.
            rateLimiter.memoryStore.clear();
            rateLimiter.config = { enabled: true, maxAttempts: 2, windowSeconds: 300, ip: { enabled: false } };
            try {
                const mailbox = await createMailbox();
                const send = () =>
                    request(server.getApplication()).get(
                        `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}?Protocol=ActiveSync`,
                    );

                expect((await send()).status).toBe(200);
                expect((await send()).status).toBe(200);
                expect((await send()).status).toBe(429);
            } finally {
                rateLimiter.config = original;
            }
        });

        it("Shares one rate-limit bucket across different queried addresses (429) - v2()'s fixed @RateLimit id keeps the bucket keyed per endpoint, not per :email URL parameter.", async () => {
            const rateLimiter: any = objectFactory.getInstance(RateLimiter);
            const original = rateLimiter.config;
            rateLimiter.memoryStore.clear();
            rateLimiter.config = { enabled: true, maxAttempts: 2, windowSeconds: 300, ip: { enabled: false } };
            try {
                const mailboxA = await createMailbox();
                const mailboxB = await createMailbox();
                const lookup = (address: string) =>
                    request(server.getApplication()).get(
                        `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(address)}?Protocol=ActiveSync`,
                    );

                expect((await lookup(mailboxA.primarySmtpAddress)).status).toBe(200);
                expect((await lookup(mailboxB.primarySmtpAddress)).status).toBe(200);
                // A third request, for yet another, still-unqueried address, is still throttled: without the
                // fixed `id`, RouteUtils.getRateLimitPath() would substitute this new address into the
                // identifier and hand it a fresh, never-exceeded bucket of its own - exactly the gap this test
                // guards against.
                expect((await lookup("someone-else@example.com")).status).toBe(429);
            } finally {
                rateLimiter.config = original;
            }
        });
    });
});
