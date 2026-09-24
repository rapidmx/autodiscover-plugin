///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// AutodiscoverXml is pure string logic with no DI/DB dependency, so it's tested directly here rather than only
// indirectly through a real-server HTTP round trip - the same precedent test/eas/codec/WbxmlCodec.test.ts sets
// for the WBXML codec.
import {
    buildOutlookSuccessXml,
    buildPoxSuccessXml,
    escapeXml,
    extractAcceptableResponseSchema,
    extractEmailAddress,
    OUTLOOK_RESPONSE_SCHEMA,
} from "../src/AutodiscoverXml.js";

describe("AutodiscoverXml Tests", () => {
    describe("extractEmailAddress", () => {
        it("Extracts the EMailAddress field from a real POX request body (no namespace prefix).", () => {
            const xml = `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="https://schemas.microsoft.com/exchange/autodiscover/mobilesync/requestschema/2006">
    <Request>
        <EMailAddress>chris@woodgrovebank.com</EMailAddress>
        <AcceptableResponseSchema>https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006</AcceptableResponseSchema>
    </Request>
</Autodiscover>`;
            expect(extractEmailAddress(xml)).toBe("chris@woodgrovebank.com");
        });

        it("Extracts the field even when it is namespace-prefixed.", () => {
            const xml = `<autodiscover:Request><autodiscover:EMailAddress>a@example.com</autodiscover:EMailAddress></autodiscover:Request>`;
            expect(extractEmailAddress(xml)).toBe("a@example.com");
        });

        it("Trims surrounding whitespace/newlines around the value.", () => {
            const xml = `<EMailAddress>\n   spaced@example.com   \n</EMailAddress>`;
            expect(extractEmailAddress(xml)).toBe("spaced@example.com");
        });

        it("Decodes XML entities in the extracted value.", () => {
            const xml = `<EMailAddress>a&amp;b@example.com</EMailAddress>`;
            expect(extractEmailAddress(xml)).toBe("a&b@example.com");
        });

        it("Returns undefined when no EMailAddress element is present.", () => {
            expect(extractEmailAddress(`<Autodiscover><Request></Request></Autodiscover>`)).toBeUndefined();
        });

        it("Returns undefined when the element is present but empty.", () => {
            expect(extractEmailAddress(`<EMailAddress></EMailAddress>`)).toBeUndefined();
        });

        it("Returns undefined for a garbage/non-XML body.", () => {
            expect(extractEmailAddress("not xml at all")).toBeUndefined();
        });

        it("Ignores self-closing, differently-named and mismatched-close elements.", () => {
            expect(extractEmailAddress(`<EMailAddress/><EMailAddressX>x@example.com</EMailAddressX>`)).toBeUndefined();
            expect(extractEmailAddress(`<EMailAddress>a@example.com</Other><EMailAddress>b@example.com</EMailAddress>`)).toBe(
                "b@example.com",
            );
            expect(extractEmailAddress(`<EMailAddress attr="1">c@example.com</ns:EMailAddress>`)).toBe("c@example.com");
        });

        it("Allows whitespace before the closing tag's '>'.", () => {
            expect(extractEmailAddress(`<EMailAddress>a@example.com</EMailAddress \r\n\t>`)).toBe("a@example.com");
            expect(extractEmailAddress(`<EMailAddress>a@example.com</EMailAddress x>`)).toBeUndefined();
        });

        it("Doesn't stop at a literal '>' inside a quoted attribute value on the opening tag.", () => {
            expect(extractEmailAddress(`<EMailAddress xmlns:x="a>b">victim@example.com</EMailAddress>`)).toBe(
                "victim@example.com",
            );
            expect(extractEmailAddress(`<EMailAddress xmlns:x='a>b'>victim@example.com</EMailAddress>`)).toBe(
                "victim@example.com",
            );
            // A self-closing tag with the same hazard still correctly finds no text content.
            expect(extractEmailAddress(`<EMailAddress xmlns:x="a>b"/><EMailAddress>real@example.com</EMailAddress>`)).toBe(
                "real@example.com",
            );
        });

        it("Skips comments, both inside the element and around it.", () => {
            expect(extractEmailAddress(`<EMailAddress>a<!-- note -->@example.com<!----></EMailAddress>`)).toBe("a@example.com");
            expect(
                extractEmailAddress(`<!-- <EMailAddress>old@example.com</EMailAddress> --><EMailAddress>new@example.com</EMailAddress>`),
            ).toBe("new@example.com");
            expect(extractEmailAddress(`<!-- <EMailAddress>old@example.com</EMailAddress>`)).toBeUndefined();
            expect(extractEmailAddress(`<EMailAddress>a@example.com<!-- unterminated`)).toBeUndefined();
        });

        it("Unwraps CDATA sections verbatim (no entity decoding inside them).", () => {
            expect(extractEmailAddress(`<EMailAddress><![CDATA[ a@example.com ]]></EMailAddress>`)).toBe("a@example.com");
            expect(extractEmailAddress(`<EMailAddress>a&amp;<![CDATA[&amp;<b>]]>@example.com</EMailAddress>`)).toBe(
                "a&&amp;<b>@example.com",
            );
            expect(extractEmailAddress(`<![CDATA[<EMailAddress>x@example.com</EMailAddress>]]>`)).toBeUndefined();
            expect(extractEmailAddress(`<EMailAddress><![CDATA[a@example.com</EMailAddress>`)).toBeUndefined();
        });

        it("Scans pathological unterminated/whitespace-padded bodies in linear time.", () => {
            const bodies = [
                `<EMailAddress>${" ".repeat(1_000_000)}`,
                `<EMailAddress>${" ".repeat(500_000)}<${" ".repeat(500_000)}`,
                "<EMailAddress".repeat(100_000),
                "<EMailAddress>".repeat(100_000),
                `<a:EMailAddress ${"<EMailAddress ".repeat(100_000)}>`,
                `<AcceptableResponseSchema>${"\t".repeat(1_000_000)}`,
            ];
            const started = Date.now();
            for (const body of bodies) {
                expect(extractEmailAddress(body)).toBeUndefined();
                expect(extractAcceptableResponseSchema(body)).toBeUndefined();
            }
            expect(Date.now() - started).toBeLessThan(2000);
        });
    });

    describe("escapeXml", () => {
        it("Escapes all 5 XML-predefined-entity characters.", () => {
            expect(escapeXml(`a&b<c>d"e'f`)).toBe("a&amp;b&lt;c&gt;d&quot;e&apos;f");
        });

        it("Leaves an already-safe string unchanged.", () => {
            expect(escapeXml("plain@example.com")).toBe("plain@example.com");
        });
    });

    describe("buildPoxSuccessXml", () => {
        it("Produces a well-formed mobilesync:Response with the expected structure and values.", () => {
            const xml = buildPoxSuccessXml({
                emailAddress: "chris@woodgrovebank.com",
                displayName: "Chris Gray",
                easUrl: "https://mail.example.com/Microsoft-Server-ActiveSync",
            });

            expect(xml).toContain(
                'xmlns:autodiscover="https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006"',
            );
            expect(xml).toContain("<autodiscover:DisplayName>Chris Gray</autodiscover:DisplayName>");
            expect(xml).toContain("<autodiscover:EMailAddress>chris@woodgrovebank.com</autodiscover:EMailAddress>");
            expect(xml).toContain("<autodiscover:Type>MobileSync</autodiscover:Type>");
            expect(xml).toContain(
                "<autodiscover:Url>https://mail.example.com/Microsoft-Server-ActiveSync</autodiscover:Url>",
            );
            expect(xml).toContain(
                "<autodiscover:Name>https://mail.example.com/Microsoft-Server-ActiveSync</autodiscover:Name>",
            );
            expect(xml).not.toContain("CertEnroll");
        });

        it("Omits the DisplayName element entirely when no display name is given.", () => {
            const xml = buildPoxSuccessXml({
                emailAddress: "noname@example.com",
                easUrl: "https://mail.example.com/Microsoft-Server-ActiveSync",
            });
            expect(xml).not.toContain("DisplayName");
        });

        it("Escapes an XML-special character in the display name/email/url values.", () => {
            const xml = buildPoxSuccessXml({
                emailAddress: "a+tag@example.com",
                displayName: `Bob & "The Builder"`,
                easUrl: "https://mail.example.com/Microsoft-Server-ActiveSync?x=1&y=2",
            });
            expect(xml).toContain("<autodiscover:EMailAddress>a+tag@example.com</autodiscover:EMailAddress>");
            expect(xml).toContain("Bob &amp; &quot;The Builder&quot;");
            expect(xml).toContain("?x=1&amp;y=2");
        });

        it("Round-trips through extractEmailAddress for the email it embeds.", () => {
            const xml = buildPoxSuccessXml({
                emailAddress: "round@example.com",
                easUrl: "https://mail.example.com/Microsoft-Server-ActiveSync",
            });
            expect(extractEmailAddress(xml)).toBe("round@example.com");
        });
    });

    describe("extractAcceptableResponseSchema", () => {
        it("Extracts the AcceptableResponseSchema field from a real Outlook request body.", () => {
            const xml = `<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
    <Request>
        <EMailAddress>chris@woodgrovebank.com</EMailAddress>
        <AcceptableResponseSchema>${OUTLOOK_RESPONSE_SCHEMA}</AcceptableResponseSchema>
    </Request>
</Autodiscover>`;
            expect(extractAcceptableResponseSchema(xml)).toBe(OUTLOOK_RESPONSE_SCHEMA);
        });

        it("Extracts the field even when it is namespace-prefixed.", () => {
            const xml = `<autodiscover:AcceptableResponseSchema>https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006</autodiscover:AcceptableResponseSchema>`;
            expect(extractAcceptableResponseSchema(xml)).toBe(
                "https://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006",
            );
        });

        it("Returns undefined when no AcceptableResponseSchema element is present.", () => {
            expect(extractAcceptableResponseSchema(`<Autodiscover><Request></Request></Autodiscover>`)).toBeUndefined();
        });

        it("Returns undefined when the element is present but empty.", () => {
            expect(extractAcceptableResponseSchema(`<AcceptableResponseSchema></AcceptableResponseSchema>`)).toBeUndefined();
        });
    });

    describe("buildOutlookSuccessXml", () => {
        it("Produces a well-formed Outlook/EXCH 2006a Response with the expected structure and values.", () => {
            const xml = buildOutlookSuccessXml({
                emailAddress: "chris@woodgrovebank.com",
                displayName: "Chris Gray",
                mapiUrl: "https://mail.example.com/mapi/emsmdb",
                nspiUrl: "https://mail.example.com/mapi/nspi",
            });

            expect(xml).toContain('xmlns="http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006"');
            expect(xml).toContain(`<Response xmlns="${OUTLOOK_RESPONSE_SCHEMA}">`);
            expect(xml).toContain("<DisplayName>Chris Gray</DisplayName>");
            expect(xml).toContain("<AutoDiscoverSMTPAddress>chris@woodgrovebank.com</AutoDiscoverSMTPAddress>");
            expect(xml).toContain('<Protocol Type="mapiHttp" Version="1">');
            expect(xml).toContain("<InternalUrl>https://mail.example.com/mapi/emsmdb</InternalUrl>");
            expect(xml).toContain("<ExternalUrl>https://mail.example.com/mapi/emsmdb</ExternalUrl>");
            expect(xml).not.toContain("<Type>EXCH</Type>");
            expect(xml).not.toContain("<Type>EXPR</Type>");
            // The address book is advertised in the same Protocol as the mail store: Outlook needs both to set an account up.
            const protocol = xml.slice(xml.indexOf("<Protocol"), xml.indexOf("</Protocol>"));
            expect(protocol).toMatch(/<MailStore>\s*<InternalUrl>https:\/\/mail\.example\.com\/mapi\/emsmdb<\/InternalUrl>\s*<ExternalUrl>https:\/\/mail\.example\.com\/mapi\/emsmdb<\/ExternalUrl>\s*<\/MailStore>\s*<AddressBook>\s*<InternalUrl>https:\/\/mail\.example\.com\/mapi\/nspi<\/InternalUrl>\s*<ExternalUrl>https:\/\/mail\.example\.com\/mapi\/nspi<\/ExternalUrl>\s*<\/AddressBook>/);
        });

        it("Leaves out the AddressBook block when no address book URL is given, and escapes it when there is one.", () => {
            expect(buildOutlookSuccessXml({ emailAddress: "a@example.com", mapiUrl: "https://mail.example.com/mapi/emsmdb" })).not.toContain("AddressBook");
            const xml = buildOutlookSuccessXml({ emailAddress: "a@example.com", mapiUrl: "https://mail.example.com/mapi/emsmdb", nspiUrl: "https://mail.example.com/mapi/nspi?x=1&y=2" });
            expect(xml).toContain("<AddressBook>");
            expect(xml).toContain("<InternalUrl>https://mail.example.com/mapi/nspi?x=1&amp;y=2</InternalUrl>");
        });

        it("Falls back to the email address as DisplayName when no display name is given.", () => {
            const xml = buildOutlookSuccessXml({
                emailAddress: "noname@example.com",
                mapiUrl: "https://mail.example.com/mapi/emsmdb",
            });
            expect(xml).toContain("<DisplayName>noname@example.com</DisplayName>");
        });

        it("Escapes an XML-special character in the display name/email/url values.", () => {
            const xml = buildOutlookSuccessXml({
                emailAddress: "a+tag@example.com",
                displayName: `Bob & "The Builder"`,
                mapiUrl: "https://mail.example.com/mapi/emsmdb?x=1&y=2",
            });
            expect(xml).toContain("<AutoDiscoverSMTPAddress>a+tag@example.com</AutoDiscoverSMTPAddress>");
            expect(xml).toContain("Bob &amp; &quot;The Builder&quot;");
            expect(xml).toContain("?x=1&amp;y=2");
        });
    });
});
