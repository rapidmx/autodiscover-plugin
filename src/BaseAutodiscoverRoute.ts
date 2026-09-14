///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ObjectDecorators } from "@rapidrest/core";
import { HttpRequest, HttpResponse, ObjectFactory, RepoUtils, RouteDecorators } from "@rapidrest/service-core";
import {
    buildOutlookSuccessXml,
    buildPoxSuccessXml,
    extractAcceptableResponseSchema,
    extractEmailAddress,
    OUTLOOK_RESPONSE_SCHEMA,
} from "./AutodiscoverXml.js";
import { Mailbox, PluginRegistry } from "@rapidmx/restapi";
const { Config, Init, Logger } = ObjectDecorators;

/** The plugin packages whose endpoints Autodiscover advertises. */
export const ACTIVESYNC_PLUGIN = "@rapidmx/activesync";
export const MAPI_PLUGIN = "@rapidmx/mapi";
const { Get, Param, Post, Query, Request, Response } = RouteDecorators;

/**
 * Abstract base for the two Autodiscover endpoints a real mail client uses to find this deployment's EAS
 * server URL from just an email address - classic POX (`POST /autodiscover/autodiscover.xml`, per
 * `[MS-ASCMD]`'s "MobileSync" response schema) and the modern JSON variant Microsoft calls "Autodiscover v2"
 * (`GET /autodiscover/autodiscover.json/v1.0/<email>?Protocol=ActiveSync`). Like `BaseMailIngestRoute`/
 * `BaseEasRoute`, this class is undecorated; its Mongo/SQL concrete classes are mounted at `/autodiscover`, which
 * composes with this class's own relative method paths to land exactly on the real spec's conventional paths (`/autodiscover/autodiscover.xml`, `/autodiscover/autodiscover.json/v1.0/:email`).
 *
 * **Auth: deliberately none.** Real classic Autodiscover conventionally expects the client to send HTTP Basic
 * credentials (email+password), with the server free to answer `401` and force re-entry - a model this
 * library's JWT-only, no-credential-verification-of-our-own boundary (see `BaseEasRoute`'s own doc comment on
 * why per-request Basic Auth was rejected there) can't and shouldn't absorb; no credential-verification
 * function exists in this codebase or its dependencies. Instead, both endpoints here follow Microsoft's own
 * Autodiscover v2 design intent exactly: answer anonymously, and reveal nothing but a deployment-wide,
 * config-supplied EAS server URL - not a per-mailbox secret - once the request email is confirmed to belong to
 * a real `Mailbox` in this deployment. The actual security boundary is unchanged from Phase 2: real mailbox
 * access still requires a JWT at `BaseEasRoute`, exactly as today.
 *
 * **Known gap, deliberately out of scope**: `Action.Redirect` (for multi-tenant hosted providers whose mailbox
 * moved to a different domain) is not implemented - this library serves exactly one EAS URL for its whole
 * deployment, so there's never a different domain to redirect to. The client-side well-known-URL discovery
 * sequence (`MS-OXDISCO`: root domain -> `autodiscover.` subdomain -> unauthenticated HTTP redirect probe ->
 * DNS SRV record -> cache) and the DNS `CNAME`/`SRV` record setup it depends on are also out of scope here -
 * both are client/deployment concerns, not application code; this class only needs to answer correctly once a
 * request actually arrives at one of its two paths.
 *
 * `mailboxClass` is supplied by the Mongo/SQL concrete subclasses following the exact one-line-per-backend
 * pattern used throughout this library. The advertised URLs are built from the `mail:autodiscover:public_url`
 * setting, and each protocol is only advertised while its plugin (`@rapidmx/activesync`, `@rapidmx/mapi`) is
 * loaded, so a client is never pointed at an endpoint this deployment doesn't serve.
 *
 * @author Jean-Philippe Steinmetz
 */
export abstract class BaseAutodiscoverRoute<M extends Mailbox> {
    protected abstract mailboxClass: any;

    /** The public base URL clients reach this deployment at, e.g. `https://mail.example.com`. */
    @Config("mail:autodiscover:public_url", "")
    protected publicUrl: string = "";

    /** The EAS endpoint URL to report, or `undefined` when ActiveSync isn't available here. */
    protected get easUrl(): string | undefined {
        return this.endpointUrl(ACTIVESYNC_PLUGIN, "/Microsoft-Server-ActiveSync");
    }

    /** The MAPI/HTTP `emsmdb` endpoint URL to report to a real Outlook desktop client, or `undefined` when MAPI
     * isn't available here. Only used by `pox()`'s Outlook/EXCH response branch. */
    protected get mapiUrl(): string | undefined {
        return this.endpointUrl(MAPI_PLUGIN, "/mapi/emsmdb");
    }

    // Automatically injected by ObjectFactory on instantiation
    private _objectFactory?: ObjectFactory;

    private mailboxRepo?: RepoUtils<M>;

    @Logger
    private logger: any;

    /**
     * Builds the query value used to match `Mailbox.aliasAddresses` against the given address. See
     * `BaseMailIngestRoute.aliasQueryValue()`'s identical doc comment - same Mongo-array vs.
     * SQL-`simple-json`-column backend split, same override point (`AutodiscoverRouteSQL` overrides this
     * identically to `MailIngestRouteSQL`).
     */
    protected aliasQueryValue(address: string): any {
        return address;
    }

    private endpointUrl(plugin: string, path: string): string | undefined {
        if (!this.publicUrl || !PluginRegistry.isActive(plugin)) {
            return undefined;
        }
        return `${this.publicUrl.replace(/\/+$/, "")}${path}`;
    }

    @Init
    public async init(): Promise<void> {
        if (!this.publicUrl) {
            this.logger?.warn("Autodiscover has no mail:autodiscover:public_url set, so it can't point clients at any protocol.");
        }
        this.mailboxRepo = await this._objectFactory!.newInstance(RepoUtils, {
            name: this.mailboxClass.name,
            args: [this.mailboxClass],
        });
    }

    private async resolveMailbox(email: string): Promise<M | undefined> {
        const address: string = email.trim().toLowerCase();
        const [byPrimary, byAlias] = await Promise.all([
            this.mailboxRepo!.find({ primarySmtpAddress: address }, { ignoreACL: true, limit: 1 }),
            this.mailboxRepo!.find({ aliasAddresses: this.aliasQueryValue(address) }, { ignoreACL: true, limit: 1 }),
        ]);
        return byPrimary[0] ?? byAlias[0];
    }

    /**
     * Classic POX Autodiscover. Per Microsoft's own client-behavior documentation, a permanent failure is just
     * as validly conveyed via a plain HTTP status as via an inner `Error` element, so this uses the HTTP-status
     * form for both error cases rather than inventing values for `[MS-ASCMD]`'s provider-specific numeric
     * error-code table.
     *
     * Branches on the request's `AcceptableResponseSchema` to decide which response shape to build: a real
     * Outlook desktop client sends `[MS-OXDSCLI]`'s Outlook/EXCH schema (`OUTLOOK_RESPONSE_SCHEMA`) to locate
     * this deployment's MAPI/HTTP endpoint (`buildOutlookSuccessXml`); any other value (or a mobile/EAS-only
     * client that omits the field entirely) gets the original MobileSync/EAS response (`buildPoxSuccessXml`).
     */
    @Post("/autodiscover.xml")
    public async pox(@Request req: HttpRequest, @Response res: HttpResponse): Promise<void> {
        if (!this.mailboxRepo) {
            res.status(500).send();
            return;
        }

        const body: string = req.rawBody ? req.rawBody.toString("utf-8") : "";
        const email: string | undefined = extractEmailAddress(body);
        if (!email) {
            res.status(400).send();
            return;
        }

        const mailbox: M | undefined = await this.resolveMailbox(email);
        if (!mailbox) {
            res.status(404).send();
            return;
        }

        const outlook: boolean = extractAcceptableResponseSchema(body) === OUTLOOK_RESPONSE_SCHEMA;
        const url: string | undefined = outlook ? this.mapiUrl : this.easUrl;
        if (!url) {
            // The protocol this client asked about isn't served here - a 404 lets it move on to its next discovery step.
            res.status(404).send();
            return;
        }
        const xml: string = outlook
            ? buildOutlookSuccessXml({ emailAddress: email, displayName: mailbox.displayName, mapiUrl: url })
            : buildPoxSuccessXml({ emailAddress: email, displayName: mailbox.displayName, easUrl: url });
        res.setHeader("Content-Type", "application/xml; charset=utf-8").status(200).send(xml);
    }

    /**
     * Autodiscover v2 (JSON). Only the `ActiveSync` protocol is served - this library has no EWS/other
     * protocol surface to advertise - so any other `Protocol` value is rejected outright rather than silently
     * answered with an EAS URL under the wrong protocol name.
     */
    @Get("/autodiscover.json/v1.0/:email")
    public async v2(
        @Param("email") email: string,
        @Query("Protocol") protocol: string | undefined,
        @Response res: HttpResponse,
    ): Promise<void> {
        if (!this.mailboxRepo) {
            res.status(500).send();
            return;
        }
        // An unsupported protocol and ActiveSync not being available here get the same answer.
        const easUrl: string | undefined = this.easUrl;
        if (protocol !== "ActiveSync" || !easUrl) {
            res.status(400).json({
                ErrorCode: "ProtocolNotSupported",
                ErrorMessage: `Unsupported protocol: ${protocol ?? ""}`,
            });
            return;
        }

        const mailbox: M | undefined = await this.resolveMailbox(decodeURIComponent(email));
        if (!mailbox) {
            res.status(404).json({
                ErrorCode: "UserNotFound",
                ErrorMessage: "No mailbox exists for the given address.",
            });
            return;
        }

        res.status(200).json({ Protocol: "ActiveSync", Url: easUrl });
    }
}
