///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
// The plugin contract: a server host registers every export of `./mongo`/`./sql` and mounts, connects or starts
// it, so each entry point must export only the mounted route, and package.json must carry a valid manifest.
import "reflect-metadata";
import fs from "fs";
import { parsePluginManifest } from "@rapidmx/restapi";
import * as MongoEntry from "../src/mongo.js";
import * as SqlEntry from "../src/sql.js";

describe("plugin entry points", () => {
    it.each([
        ["mongo", MongoEntry, "Mongo"],
        ["sql", SqlEntry, "SQL"],
    ])("./%s exports only the mounted Autodiscover route", (_name, entry, suffix) => {
        const paths = Object.fromEntries(Object.entries(entry).map(([name, clazz]) => [name, Reflect.getMetadata("rrst:routePaths", (clazz as any).prototype)]));
        expect(paths).toEqual({ [`AutodiscoverRoute${suffix}`]: ["/autodiscover"] });
    });

    it("declares a valid plugin manifest", () => {
        const pkg = JSON.parse(fs.readFileSync(new URL("../package.json", import.meta.url), "utf8"));
        expect(parsePluginManifest(pkg)).toEqual(expect.objectContaining({ displayName: "Autodiscover" }));
    });
});
