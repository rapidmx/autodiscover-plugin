import { RouteDecorators } from "@rapidrest/service-core";
import { AutodiscoverRouteMongo } from "../../../src/mongo/AutodiscoverRouteMongo.js";
const { Route } = RouteDecorators;

@Route("/mongo/autodiscover")
export class AutodiscoverRoute extends AutodiscoverRouteMongo {}
