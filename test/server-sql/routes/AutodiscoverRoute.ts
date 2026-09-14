import { RouteDecorators } from "@rapidrest/service-core";
import { AutodiscoverRouteSQL } from "../../../src/sql/AutodiscoverRouteSQL.js";
const { Route } = RouteDecorators;

@Route("/sql/autodiscover")
export class AutodiscoverRoute extends AutodiscoverRouteSQL {}
