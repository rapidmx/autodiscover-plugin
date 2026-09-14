///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { Raw } from "typeorm";
import { MailboxSQL } from "@rapidmx/restapi/sql";
import { RouteDecorators } from "@rapidrest/service-core";
import { BaseAutodiscoverRoute } from "../BaseAutodiscoverRoute.js";
const { Route } = RouteDecorators;

/**
 * SQL-backed concrete `BaseAutodiscoverRoute`, mounted at `/autodiscover` and exported from this plugin's `./sql`
 * entry point. `aliasQueryValue()` is overridden identically to `MailIngestRouteSQL` - see
 * that class's doc comment for the full vulnerability rationale (unescaped `%`/`_` enabling alias-enumeration
 * and cross-mailbox matches) this same escaping closes here too, since `resolveMailbox()` is reached by
 * unauthenticated callers exactly like `MailIngestRoute.resolve()` is.
 *
 * @author Jean-Philippe Steinmetz
 */
@Route("/autodiscover")
export class AutodiscoverRouteSQL extends BaseAutodiscoverRoute<MailboxSQL> {
    protected mailboxClass: any = MailboxSQL;

    protected aliasQueryValue(address: string): any {
        const escaped: string = address.replace(/[\\%_]/g, (ch) => `\\${ch}`);
        return Raw((alias) => `${alias} LIKE :pattern ESCAPE '\\'`, { pattern: `%"${escaped}"%` });
    }
}
