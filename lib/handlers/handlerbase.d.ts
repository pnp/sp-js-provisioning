import { Web } from "sp-pnp-js";
import { ProvisioningContext } from "../provisioningcontext";
/**
 * Describes the Object Handler Base
 */
export declare class HandlerBase {
    private name;
    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(name: string);
    /**
     * Provisioning objects
     */
    ProvisionObjects(web: Web, templatePart: any, _context?: ProvisioningContext): Promise<void>;
    /**
     * Writes to Logger when scope has started
     */
    scope_started(): void;
    /**
     * Writes to Logger when scope has stopped
     */
    scope_ended(): void;
}
