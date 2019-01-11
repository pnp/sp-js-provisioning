import { Web } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Object Handler Base
 */
export class HandlerBase {

    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(private name: string, public config: IProvisioningConfig = {}) { }

    /**
     * Provisioning objects
     */
    public ProvisionObjects(web: Web, templatePart: any, _context?: ProvisioningContext): Promise<void> {
        Logger.log({ data: templatePart, level: LogLevel.Warning, message: `Handler ${this.name} for web [${web.toUrl()}] does not override ProvisionObjects.` });
        return Promise.resolve();
    }

    /**
     * Writes to Logger when scope has started
     */
    public scope_started() {
        this.log_info("ProvisionObjects", "Code execution scope started");
    }

    /**
     * Writes to Logger when scope has stopped
     */
    public scope_ended() {
        this.log_info("ProvisionObjects", "Code execution scope ended");
    }

    /**
     * Writes to Logger
     *
     * @param {string} scope Scope
     * @param {string} message Message
     */
    public log_info(scope: string, message: string) {
        let prefix = (this.config.logging && this.config.logging.prefix) ? `(${this.config.logging.prefix}) ` : "";
        Logger.write(`${prefix}${this.name}: ${scope}: ${message}`);
    }
}
