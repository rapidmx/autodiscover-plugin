///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// Isolated unit tests for BaseAutodiscoverRoute, reserved for defensive guard branches a real wired server can
// never exercise: the `!this.mailboxRepo` guards (DI always populates the repo via `@Init` before a request can
// reach a route - the same rationale test/routes/BaseEasRoute.test.ts already uses for its own identical guard)
// and `pox()`'s `req.rawBody ? ... : ""` fallback (a real HTTP transport always supplies at least an empty,
// still-truthy `Buffer` for a body-eligible request, so the falsy/`undefined` branch is unreachable over real
// HTTP - only a direct unit call can construct a request object without a `rawBody` at all). Every other
// behavior (email extraction, mailbox resolution, both success/error response shapes, schema branching) is
// exercised via real HTTP+DB requests in test/routes/mongo/AutodiscoverRoute.test.ts (and its sql/ counterpart).
import config from "../config.js";
import { ObjectFactory } from "@rapidrest/service-core";
import { Logger } from "@rapidrest/core";
import { PluginRegistry } from "@rapidmx/restapi";
import { BaseAutodiscoverRoute, MAX_POX_BODY_BYTES } from "../../src/BaseAutodiscoverRoute.js";

class TestAutodiscoverRoute extends BaseAutodiscoverRoute<any> {
    protected mailboxClass: any = { name: "TestMailbox" };
}

function makeRes(): any {
    return {
        status: vi.fn().mockReturnThis(),
        setHeader: vi.fn().mockReturnThis(),
        json: vi.fn().mockReturnThis(),
        send: vi.fn().mockReturnThis(),
    };
}

describe("BaseAutodiscoverRoute Tests (guard clauses only)", () => {
    const objectFactory: ObjectFactory = new ObjectFactory(config, Logger());

    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("pox() sends a bare 500 when mailboxRepo is not set.", async () => {
        // `initialize: false` skips `@Init`, leaving mailboxRepo genuinely `undefined` - exactly what this
        // guard clause exists to catch.
        const route = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, {
            initialize: false,
        });
        const res = makeRes();

        await route.pox({ rawBody: undefined } as any, res);

        expect(res.status).toHaveBeenCalledWith(500);
        expect(res.send).toHaveBeenCalledWith();
    });

    it("pox() sends a 400 when req.rawBody is undefined (falls back to an empty body string).", async () => {
        const route = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, {
            initialize: false,
        });
        (route as any).mailboxRepo = { find: vi.fn().mockResolvedValue([]) };
        const res = makeRes();

        await route.pox({ rawBody: undefined } as any, res);

        expect(res.status).toHaveBeenCalledWith(400);
        expect(res.send).toHaveBeenCalledWith();
    });

    it("advertises nothing and warns at startup without a public URL, and only loaded protocols with one.", async () => {
        const route: any = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, { initialize: false });
        route.logger = { warn: vi.fn() };
        vi.spyOn(objectFactory, "newInstance").mockResolvedValue({} as any);
        route.publicUrl = "";
        await route.init();
        expect(route.logger.warn).toHaveBeenCalledWith(expect.stringMatching(/mail:autodiscover:public_url/));
        PluginRegistry.setLoaded([{ name: "@rapidmx/activesync-plugin", version: "1.0.0" }]);
        try {
            expect(route.easUrl).toBeUndefined();
            route.publicUrl = "https://mail.example.com//";
            await route.init();
            expect(route.logger.warn).toHaveBeenCalledTimes(1);
            expect(route.easUrl).toBe("https://mail.example.com/Microsoft-Server-ActiveSync");
            expect(route.mapiUrl).toBeUndefined();
        } finally {
            PluginRegistry.setLoaded([]);
        }
    });

    it("rejects an insecure or malformed public URL: warns at startup and advertises nothing.", async () => {
        const route: any = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, { initialize: false });
        vi.spyOn(objectFactory, "newInstance").mockResolvedValue({} as any);
        PluginRegistry.setLoaded([{ name: "@rapidmx/activesync-plugin", version: "1.0.0" }]);
        try {
            for (const bad of [
                "http://mail.example.com",
                "https://mail.example.com/?x=1",
                "https://mail.example.com/?",
                "https://mail.example.com/#frag",
                "https://user:pw@mail.example.com",
                "mail.example.com",
                "ftp://mail.example.com",
                "https://mail example.com",
            ]) {
                route.logger = { warn: vi.fn() };
                route.publicUrl = bad;
                await route.init();
                expect(route.logger.warn).toHaveBeenCalledWith(expect.stringMatching(/not a valid https/));
                expect(route.easUrl).toBeUndefined();
            }
            for (const [good, expected] of [
                ["  https://mail.example.com/base/  ", "https://mail.example.com/base/Microsoft-Server-ActiveSync"],
                ["http://localhost:3000", "http://localhost:3000/Microsoft-Server-ActiveSync"],
                ["http://127.0.0.1:3000/", "http://127.0.0.1:3000/Microsoft-Server-ActiveSync"],
            ]) {
                route.logger = { warn: vi.fn() };
                route.publicUrl = good;
                await route.init();
                expect(route.logger.warn).not.toHaveBeenCalled();
                expect(route.easUrl).toBe(expected);
            }
        } finally {
            PluginRegistry.setLoaded([]);
        }
    });

    it("pox() sends a 413 for an oversized body before scanning it.", async () => {
        const route = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, {
            initialize: false,
        });
        const find = vi.fn().mockResolvedValue([]);
        (route as any).mailboxRepo = { find };
        const res = makeRes();

        await route.pox({ rawBody: Buffer.from(`<EMailAddress>a@example.com</EMailAddress>${" ".repeat(MAX_POX_BODY_BYTES)}`) } as any, res);

        expect(res.status).toHaveBeenCalledWith(413);
        expect(find).not.toHaveBeenCalled();
    });

    it("never queries for a non-plain address, and always queries a plain one as a literal eq() value.", async () => {
        const route = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, {
            initialize: false,
        });
        const find = vi.fn().mockResolvedValue([]);
        (route as any).mailboxRepo = { find };
        PluginRegistry.setLoaded([{ name: "@rapidmx/activesync-plugin", version: "1.0.0" }]);
        (route as any).publicUrl = "https://mail.example.com";
        try {
            for (const value of ["like(*)", "regex(^a)", "in(a@example.com,b@example.com)", "ne(a@example.com)", "a@b@c"]) {
                const res = makeRes();
                await route.pox({ rawBody: Buffer.from(`<EMailAddress>${value}</EMailAddress>`) } as any, res);
                expect(res.status).toHaveBeenCalledWith(400);
                const v2Res = makeRes();
                await route.v2(value, "ActiveSync", v2Res);
                expect(v2Res.status).toHaveBeenCalledWith(404);
            }
            expect(find).not.toHaveBeenCalled();

            await route.v2("  Ada@Example.com ", "ActiveSync", makeRes());
            expect(find).toHaveBeenCalledWith({ primarySmtpAddress: "eq(ada@example.com)", limit: 1 }, { ignoreACL: true, limit: 1 });
            expect(find).toHaveBeenCalledWith({ aliasAddresses: "eq(ada@example.com)", limit: 1 }, { ignoreACL: true, limit: 1 });
        } finally {
            PluginRegistry.setLoaded([]);
        }
    });

    it("v2() sends a bare 500 when mailboxRepo is not set.", async () => {
        const route = objectFactory.newInstance<TestAutodiscoverRoute>(TestAutodiscoverRoute, {
            initialize: false,
        });
        const res = makeRes();

        await route.v2("someone@example.com", "ActiveSync", res);

        expect(res.status).toHaveBeenCalledWith(500);
        expect(res.send).toHaveBeenCalledWith();
    });
});
