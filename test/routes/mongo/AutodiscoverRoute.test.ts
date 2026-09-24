///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// These tests prove BaseAutodiscoverRoute's real HTTP behavior for both endpoints it exposes - classic POX
// (`POST .../autodiscover.xml`) and Autodiscover v2 (`GET .../autodiscover.json/v1.0/:email`) - against a real
// server + real Mongo, with no authentication (see the architecture plan's Autodiscover section for why both
// endpoints are deliberately unauthenticated).
import config from "../../config.js";
import { request } from "@rapidrest/service-core/test";
import { MongoConnection, MongoRepository, Server, ObjectFactory, ConnectionManager, RateLimiter } from "@rapidrest/service-core";
import { Logger } from "@rapidrest/core";
import * as uuid from "uuid";
import { MailboxMongo } from "@rapidmx/restapi/mongo";
import { MongoMemoryServer } from "mongodb-memory-server";
import { registerTestDoubles } from "../../testDoubles.js";
import { PluginRegistry } from "@rapidmx/restapi";

const mongod: MongoMemoryServer = new MongoMemoryServer({
    instance: {
        port: 9999,
        dbName: "rrst-test",
    },
});

describe("Route:AutodiscoverRouteMongo Tests", () => {
    const logger = Logger();
    const objectFactory: ObjectFactory = new ObjectFactory(config, logger);
    const server: Server = new Server({ config, basePath: "./test/server-mongo", logger, objectFactory });
    const baseUrl = "/mongo/autodiscover";
    let mailboxRepo: MongoRepository<MailboxMongo>;

    const createMailbox = async function (data?: Partial<MailboxMongo>): Promise<MailboxMongo> {
        const obj: MailboxMongo = new MailboxMongo({
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
        await mongod.start();
        registerTestDoubles(objectFactory);
        await server.start();
        PluginRegistry.setLoaded([
            { name: "@rapidmx/activesync-plugin", version: "1.0.0" },
            { name: "@rapidmx/mapi-plugin", version: "1.0.0" },
        ]);

        const connMgr: ConnectionManager | undefined = objectFactory.getInstance(ConnectionManager);
        const conn: any = connMgr?.connections.get("mongo");
        if (conn instanceof MongoConnection) {
            mailboxRepo = conn.getMongoRepository("MailboxMongo");
        } else {
            throw new Error("Could not find mongo connection");
        }
    });

    afterAll(async () => {
        await server.stop();
        await mongod.stop();
        await objectFactory.destroy();
    });

    beforeEach(async () => {
        try {
            await mailboxRepo.clear();
        } catch (err: any) {
            if (err.message !== "ns not found") {
                throw err;
            }
        }
    });

    describe("POST .../autodiscover.xml (POX)", () => {
        it("Requires no authentication.", async () => {
            const mailbox = await createMailbox();
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress));
            expect(result.status).toBe(200);
        });

        it("Returns a MobileSync settings response for a known primarySmtpAddress.", async () => {
            const mailbox = await createMailbox({ displayName: "Ada Lovelace" });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress));

            expect(result.status).toBe(200);
            expect(result.headers["content-type"]).toContain("text/xml");
            const xml = result.text;
            expect(xml).toContain(`<autodiscover:EMailAddress>${mailbox.primarySmtpAddress}</autodiscover:EMailAddress>`);
            expect(xml).toContain(`<autodiscover:DisplayName>${mailbox.primarySmtpAddress}</autodiscover:DisplayName>`);
            expect(result.text).not.toContain("Ada Lovelace");
            expect(xml).toContain("<autodiscover:Type>MobileSync</autodiscover:Type>");
            expect(xml).toContain("<autodiscover:Url>https://mail.example.com/Microsoft-Server-ActiveSync</autodiscover:Url>");
        });

        it("Matches against aliasAddresses, not only primarySmtpAddress.", async () => {
            const alias = `${uuid.v4()}@example.com`;
            await createMailbox({ aliasAddresses: [alias] });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(alias));

            expect(result.status).toBe(200);
            expect(result.text).toContain(`<autodiscover:EMailAddress>${alias}</autodiscover:EMailAddress>`);
        });

        it("Matches case-insensitively.", async () => {
            const mailbox = await createMailbox();

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress.toUpperCase()));

            expect(result.status).toBe(200);
        });

        it("Returns 404 for an address with no matching Mailbox.", async () => {
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody("nobody@example.com"));
            expect(result.status).toBe(404);
        });

        it("Returns 400, never a directory match, for query-operator or non-plain address values.", async () => {
            const mailbox = await createMailbox();
            const alias = `a${uuid.v4()}@example.com`;
            await createMailbox({ aliasAddresses: [alias] });
            for (const value of [
                "like(*)",
                "regex(^a)",
                "ne(nobody@example.com)",
                `in(${mailbox.primarySmtpAddress},${alias})`,
                "like(*@example.com)",
                `eq(${mailbox.primarySmtpAddress})`,
                `${mailbox.primarySmtpAddress},${alias}`,
                "no-at-sign",
                `a b@example.com`,
                `${"a".repeat(250)}@example.com`,
            ]) {
                const result = await request(server.getApplication())
                    .post(`${baseUrl}/autodiscover.xml`)
                    .set("Content-Type", "text/xml")
                    .send(poxRequestBody(value));
                expect(result.status).toBe(400);
            }
        });

        it("Returns 400 when the request body has no EMailAddress element.", async () => {
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send("<Autodiscover><Request></Request></Autodiscover>");
            expect(result.status).toBe(400);
        });

        it("Answers a pathological unterminated, whitespace-padded EMailAddress body quickly with 400.", async () => {
            const started = Date.now();
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(`<EMailAddress>${" ".repeat(15000)}`);
            expect(result.status).toBe(400);
            expect(Date.now() - started).toBeLessThan(1000);
        });

        it("Returns 413 for a request body over 16KB without scanning it.", async () => {
            const mailbox = await createMailbox();
            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(poxRequestBody(mailbox.primarySmtpAddress) + " ".repeat(17 * 1024));
            expect(result.status).toBe(413);
        });

        it("Returns an Outlook/EXCH mapiHttp response when AcceptableResponseSchema requests the Outlook schema.", async () => {
            const mailbox = await createMailbox({ displayName: "Ada Lovelace" });

            const result = await request(server.getApplication())
                .post(`${baseUrl}/autodiscover.xml`)
                .set("Content-Type", "text/xml")
                .send(outlookPoxRequestBody(mailbox.primarySmtpAddress));

            expect(result.status).toBe(200);
            expect(result.headers["content-type"]).toContain("text/xml");
            const xml = result.text;
            expect(xml).toContain("http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a");
            expect(xml).toContain(`<AutoDiscoverSMTPAddress>${mailbox.primarySmtpAddress}</AutoDiscoverSMTPAddress>`);
            expect(xml).toContain(`<DisplayName>${mailbox.primarySmtpAddress}</DisplayName>`);
            expect(xml).not.toContain("Ada Lovelace");
            expect(xml).toContain('<Protocol Type="mapiHttp" Version="1">');
            expect(xml).toContain("<InternalUrl>https://mail.example.com/mapi/emsmdb</InternalUrl>");
            expect(xml).toContain("<AddressBook>");
            expect(xml).toContain("<InternalUrl>https://mail.example.com/mapi/nspi</InternalUrl>");
            expect(xml).toContain("<ExternalUrl>https://mail.example.com/mapi/nspi</ExternalUrl>");
            expect(xml).not.toContain("<autodiscover:Type>MobileSync</autodiscover:Type>");
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
        it("Requires no authentication and returns the EAS URL for a known address.", async () => {
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

        it("Matches against aliasAddresses, not only primarySmtpAddress.", async () => {
            const alias = `${uuid.v4()}@example.com`;
            await createMailbox({ aliasAddresses: [alias] });

            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(alias)}?Protocol=ActiveSync`,
            );

            expect(result.status).toBe(200);
            expect(result.body.Url).toBe("https://mail.example.com/Microsoft-Server-ActiveSync");
        });

        it("Returns 404 JSON for an address with no matching Mailbox.", async () => {
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/nobody@example.com?Protocol=ActiveSync`,
            );
            expect(result.status).toBe(404);
            expect(result.body.ErrorCode).toBe("UserNotFound");
        });

        it("Returns 404 JSON, never a directory match, for query-operator or non-plain address values.", async () => {
            const mailbox = await createMailbox();
            for (const value of ["like(*)", "regex(^a)", `in(${mailbox.primarySmtpAddress},x@example.com)`, "ne(x@example.com)"]) {
                const result = await request(server.getApplication()).get(
                    `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(value)}?Protocol=ActiveSync`,
                );
                expect(result.status).toBe(404);
                expect(result.body.ErrorCode).toBe("UserNotFound");
            }
        });

        it("Returns 404 JSON (not 500) for an address containing a literal percent sign.", async () => {
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent("100%off@example.com")}?Protocol=ActiveSync`,
            );
            expect(result.status).toBe(404);
            expect(result.body.ErrorCode).toBe("UserNotFound");
        });

        it("Returns 400 JSON for an unsupported Protocol value.", async () => {
            const mailbox = await createMailbox();
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}?Protocol=EWS`,
            );
            expect(result.status).toBe(400);
            expect(result.body.ErrorCode).toBe("ProtocolNotSupported");
        });

        it("Returns 400 JSON when Protocol is missing entirely.", async () => {
            const mailbox = await createMailbox();
            const result = await request(server.getApplication()).get(
                `${baseUrl}/autodiscover.json/v1.0/${encodeURIComponent(mailbox.primarySmtpAddress)}`,
            );
            expect(result.status).toBe(400);
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
