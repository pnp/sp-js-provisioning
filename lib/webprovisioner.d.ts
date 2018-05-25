import { Schema } from "./schema";
import { HandlerBase } from "./handlers/handlerbase";
import { TypedHash, Web } from "sp-pnp-js";
/**
 * Root class of Provisioning
 */
export declare class WebProvisioner {
    private web;
    handlerMap: TypedHash<HandlerBase>;
    handlerSort: TypedHash<number>;
    /**
     * Creates a new instance of the Provisioner class
     *
     * @param {Web} web The Web instance to which we want to apply templates
     * @param {TypedHash<HandlerBase>} handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    constructor(web: Web, handlerMap?: TypedHash<HandlerBase>, handlerSort?: TypedHash<number>);
    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param {Schema} template The template to apply
     * @param {Function} progressCallback Callback for progress updates
     */
    applyTemplate(template: Schema, progressCallback?: (msg: string) => void): Promise<void>;
}
