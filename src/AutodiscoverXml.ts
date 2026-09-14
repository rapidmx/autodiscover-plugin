///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/**
 * Minimal, purpose-built XML handling for the classic POX Autodiscover request/response (`[MS-ASCMD]`'s
 * "MobileSync" schema) - deliberately NOT a general-purpose XML/DOM parser. `BaseAutodiscoverRoute`'s POX
 * endpoint is unauthenticated and internet-facing (see its own doc comment for why), and the real request
 * schema has exactly one field this library ever reads (`EMailAddress`), so pulling in a general XML parser -
 * especially one not hardened against XXE/entity-expansion - would be real attack surface for zero benefit.
 * This follows the same "hand-build the wire format precisely, no library" approach Phase 2's WBXML codec
 * already took, just for text XML instead of binary.
 */

/** Reads an XML name (`[\w:.-]*`) starting at `start`, returning it lower-cased plus the index just past it. */
function readTagName(xml: string, start: number): { name: string; end: number } {
    let end = start;
    while (end < xml.length && /[\w:.-]/.test(xml[end])) {
        end++;
    }
    return { name: xml.slice(start, end).toLowerCase(), end };
}

/** Whether a (lower-cased) tag name is `localName`, optionally namespace-prefixed (`ns:localName`). */
function isElement(name: string, localName: string): boolean {
    return name === localName || name.endsWith(`:${localName}`);
}

/**
 * Returns the raw text content of the first `<localName>text</localName>` element (namespace prefix and
 * attributes allowed, case-insensitive) whose content is plain text. A single forward, linear-time scan -
 * deliberately NOT a regex: this runs on an unauthenticated request body, and the previous
 * `\s*([^<]*?)\s*<\/...>` pattern backtracked cubically on an unterminated element padded with whitespace.
 * Every `indexOf` below starts past the previous one, so no character is rescanned more than a constant
 * number of times.
 */
function extractElementText(xml: string, localName: string): string | undefined {
    const target = localName.toLowerCase();
    let pos = 0;
    while ((pos = xml.indexOf("<", pos)) !== -1) {
        const open = readTagName(xml, pos + 1);
        pos = open.end;
        if (!isElement(open.name, target)) {
            continue;
        }
        const gt = xml.indexOf(">", pos);
        if (gt === -1) {
            return undefined;
        }
        pos = gt + 1;
        if (xml[gt - 1] === "/") {
            continue; // Self-closing, so no text content.
        }
        const lt = xml.indexOf("<", pos);
        if (lt === -1) {
            return undefined;
        }
        if (xml[lt + 1] === "/") {
            const close = readTagName(xml, lt + 2);
            if (isElement(close.name, target) && xml[close.end] === ">") {
                return xml.slice(pos, lt);
            }
        }
        pos = lt;
    }
    return undefined;
}

/**
 * Extracts the request's `<EMailAddress>` text content via a small, linear-time tag scan rather than a DOM
 * parse. Never resolves entities/DTDs - only the literal text between the open/close tags is ever inspected.
 */
export function extractEmailAddress(xml: string): string | undefined {
    const value = extractElementText(xml, "EMailAddress")?.trim();
    return value ? decodeXmlEntities(value) : undefined;
}

/** Escapes the 5 XML-predefined-entity characters for safe inclusion as element text content. */
export function escapeXml(value: string): string {
    return value
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function decodeXmlEntities(value: string): string {
    return value
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&");
}

export interface AutodiscoverPoxSuccess {
    emailAddress: string;
    displayName?: string;
    easUrl: string;
}

/** The `AcceptableResponseSchema` value a real Outlook desktop client sends to request the Outlook/EXCH
 * response shape (`buildOutlookSuccessXml`) instead of the EAS-only MobileSync one - confirmed against
 * `[MS-OXDSCLI]`'s own Autodiscover Response XSD, whose target namespace is exactly this value. */
export const OUTLOOK_RESPONSE_SCHEMA = "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a";

/**
 * Extracts the request's `<AcceptableResponseSchema>` text content via the same linear-time tag scan as
 * `extractEmailAddress` - never a DOM parse, same rationale (unauthenticated, internet-facing endpoint).
 */
export function extractAcceptableResponseSchema(xml: string): string | undefined {
    return extractElementText(xml, "AcceptableResponseSchema")?.trim() || undefined;
}

/**
 * Builds a real `mobilesync:Response` success document per `[MS-ASCMD]`'s `AutodiscoverMobileSync.xsd` -
 * `Culture`/`User`/`Action.Settings.Server{Type=MobileSync,Url,Name}`, matching the shape confirmed against
 * Microsoft's own "Autodiscover for Exchange ActiveSync developers" example response. A real response can also
 * carry a `CertEnroll` server block for client-certificate enrollment - deliberately omitted, same "pragmatic
 * subset" scoping as Phase 2's own documented gaps, since this library has no certificate-enrollment support.
 */
export function buildPoxSuccessXml({ emailAddress, displayName, easUrl }: AutodiscoverPoxSuccess): string {
    const email = escapeXml(emailAddress);
    const url = escapeXml(easUrl);
    const displayNameElement = displayName
        ? `\n            <autodiscover:DisplayName>${escapeXml(displayName)}</autodiscover:DisplayName>`
        : "";
    return `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns:autodiscover="https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006">
    <autodiscover:Response>
        <autodiscover:Culture>en:us</autodiscover:Culture>
        <autodiscover:User>${displayNameElement}
            <autodiscover:EMailAddress>${email}</autodiscover:EMailAddress>
        </autodiscover:User>
        <autodiscover:Action>
            <autodiscover:Settings>
                <autodiscover:Server>
                    <autodiscover:Type>MobileSync</autodiscover:Type>
                    <autodiscover:Url>${url}</autodiscover:Url>
                    <autodiscover:Name>${url}</autodiscover:Name>
                </autodiscover:Server>
            </autodiscover:Settings>
        </autodiscover:Action>
    </autodiscover:Response>
</Autodiscover>`;
}

export interface AutodiscoverOutlookSuccess {
    emailAddress: string;
    displayName?: string;
    mapiUrl: string;
}

/**
 * Builds a real Outlook/EXCH `2006a` response per `[MS-OXDSCLI]`'s confirmed Autodiscover Response XSD -
 * an `Autodiscover`(namespace `.../responseschema/2006`) element wrapping a `Response`
 * (namespace `.../outlook/responseschema/2006a`) with `User`/`Account` children.
 *
 * Since this library only speaks MAPI/HTTP (no classic RPC/TCP MAPI transport), the response always
 * advertises exactly one `Protocol`, using `Type`/`Version` as XML ATTRIBUTES rather than the classic
 * `<Type>EXCH</Type>` child element - `[MS-OXDSCLI]`'s "Processing the X-MapiHttpCapability Header" section
 * confirms a mapiHttp-capable client's response "MUST include a Protocol element that contains a Type
 * attribute set to 'mapiHttp' and a Version attribute" and "MUST NOT include a Protocol element that contains
 * a Type element set to 'EXCH' or 'EXPR'" - the two forms are mutually exclusive by design, not merely
 * alternatives. This deployment has no classic RPC/TCP MAPI endpoint to fall back to, so every Outlook-schema
 * request always gets the mapiHttp form - the real `X-MapiHttpCapability` request-header negotiation a full
 * implementation would consult to choose between EXCH/mapiHttp is not implemented, a documented gap.
 * `MailStore.InternalUrl`/`ExternalUrl` are both set to the same `mapiUrl`, since this library serves one URL
 * per deployment with no separate internal/external network split.
 *
 * `LegacyDN`/`DeploymentId` are schema-required fields this library has no real backing value for (no X.500
 * DN resolution, no multi-tenant deployment identity) - both are synthesized placeholders. Real Outlook does
 * not validate their exact content for MAPI/HTTP connectivity; they matter for classic RPC/TCP MAPI
 * free-busy/permissions lookups this library doesn't implement anyway.
 */
export function buildOutlookSuccessXml({ emailAddress, displayName, mapiUrl }: AutodiscoverOutlookSuccess): string {
    const email = escapeXml(emailAddress);
    const url = escapeXml(mapiUrl);
    const name = escapeXml(displayName ?? emailAddress);
    const legacyDn = escapeXml(
        `/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=${emailAddress}`,
    );
    return `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006">
    <Response xmlns="${OUTLOOK_RESPONSE_SCHEMA}">
        <User>
            <DisplayName>${name}</DisplayName>
            <LegacyDN>${legacyDn}</LegacyDN>
            <AutoDiscoverSMTPAddress>${email}</AutoDiscoverSMTPAddress>
            <DeploymentId>00000000-0000-0000-0000-000000000000</DeploymentId>
        </User>
        <Account>
            <AccountType>email</AccountType>
            <Action>settings</Action>
            <MicrosoftOnline>False</MicrosoftOnline>
            <Protocol Type="mapiHttp" Version="1">
                <MailStore>
                    <InternalUrl>${url}</InternalUrl>
                    <ExternalUrl>${url}</ExternalUrl>
                </MailStore>
            </Protocol>
        </Account>
    </Response>
</Autodiscover>`;
}
