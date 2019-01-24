import { TypedHash } from "@pnp/common";
import { HandlerBase } from "./handlerbase";
import { IProvisioningConfig } from "../provisioningconfig";
export declare const DefaultHandlerMap: (config: IProvisioningConfig) => TypedHash<HandlerBase>;
export declare const DefaultHandlerSort: TypedHash<number>;
