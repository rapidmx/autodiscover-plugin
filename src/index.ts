///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/**
 * Autodiscover protocol compatibility: the mechanism a real mail client uses to find this deployment's EAS
 * server URL from just an email address, before it has any credentials or JWT of its own. Exports the
 * backend-agnostic surface - the request/response XML helpers and the abstract route base class. The concrete
 * Mongo/SQL classes a deployment actually instantiates (`AutodiscoverRouteMongo`/`AutodiscoverRouteSQL`) are
 * exported from this package's `./mongo`/`./sql` subpaths instead, alongside every other entity/route/job this
 * library defines - see `src/eas/index.ts`'s identical doc comment for the same convention.
 *
 * A deployment mounts Autodiscover by loading the ready-to-mount `AutodiscoverRouteMongo`/`AutodiscoverRouteSQL`
 * class from this plugin's `./mongo`/`./sql` entry point directly - no subclass or URLs to supply, since
 * `BaseAutodiscoverRoute` builds the EAS (`easUrl`) and MAPI (`mapiUrl`) endpoint URLs itself from the
 * `mail:autodiscover:public_url` config setting (also exposed as this plugin's own "Public server URL"
 * admin-console setting - see `package.json`'s `rapidmx.plugin.settings`), and only advertises each protocol
 * while its own plugin (`@rapidmx/activesync-plugin`, `@rapidmx/mapi-plugin`) is loaded. Set
 * `mail:autodiscover:public_url` to this deployment's externally-reachable `https://` base URL (e.g.
 * `https://mail.example.com`) for either endpoint to be advertised at all - see `BaseAutodiscoverRoute`'s own
 * doc comment for the exact validation rules.
 */
export * from "./AutodiscoverXml.js";
export * from "./BaseAutodiscoverRoute.js";
