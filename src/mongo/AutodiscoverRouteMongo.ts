///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { MailboxMongo } from "@rapidmx/restapi/mongo";
import { RouteDecorators } from "@rapidrest/service-core";
import { BaseAutodiscoverRoute } from "../BaseAutodiscoverRoute.js";
const { Route } = RouteDecorators;

/**
 * Mongo-backed concrete `BaseAutodiscoverRoute`, mounted at `/autodiscover`. Exported from this plugin's `./mongo`
 * entry point, so the server host mounts it without a wrapper class of its own.
 *
 * @author Jean-Philippe Steinmetz
 */
@Route("/autodiscover")
export class AutodiscoverRouteMongo extends BaseAutodiscoverRoute<MailboxMongo> {
    protected mailboxClass: any = MailboxMongo;
}
