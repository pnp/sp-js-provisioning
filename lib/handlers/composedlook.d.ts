import { IComposedLook } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web } from "sp-pnp-js";
/**
 * Describes the Composed Look Object Handler
 */
export declare class ComposedLook extends HandlerBase {
    /**
     * Creates a new instance of the ObjectComposedLook class
     */
    constructor();
    /**
     * Provisioning Composed Look
     *
     * @param {Web} web The web
     * @param {IComposedLook} object The Composed Look to provision
     */
    ProvisionObjects(web: Web, composedLook: IComposedLook): Promise<void>;
}
