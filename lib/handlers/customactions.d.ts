import { HandlerBase } from "./handlerbase";
import { ICustomAction } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the Custom Actions Object Handler
 */
export declare class CustomActions extends HandlerBase {
    /**
     * Creates a new instance of the ObjectCustomActions class
     */
    constructor();
    /**
     * Provisioning Custom Actions
     *
     * @param {Web} web The web
     * @param {Array<ICustomAction>} customactions The Custom Actions to provision
     */
    ProvisionObjects(web: Web, customActions: ICustomAction[]): Promise<void>;
}
