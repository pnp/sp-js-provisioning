import { HandlerBase } from "./handlerbase";
import { IPropertyBagEntry } from "../schema";
import { Web } from "sp-pnp-js";
/**
 * Describes the PropertyBagEntries Object Handler
 */
export declare class PropertyBagEntries extends HandlerBase {
    /**
     * Creates a new instance of the PropertyBagEntries class
     */
    constructor();
    /**
     * Provisioning property bag entries
     *
     * @param {Web} web The web
     * @param {Array<IPropertyBagEntry>} entries The property bag entries to provision
     */
    ProvisionObjects(web: Web, entries: IPropertyBagEntry[]): Promise<void>;
}
