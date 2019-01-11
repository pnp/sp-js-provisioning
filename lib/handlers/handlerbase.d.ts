import { Web } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Object Handler Base
 */
export declare class HandlerBase {
    private name;
    config: IProvisioningConfig;
    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(name: string, config?: IProvisioningConfig);
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
    /**
     * Writes to Logger
     *
     * @param {string} scope Scope
     * @param {string} message Message
     */
    log_info(scope: string, message: string): void;
}
