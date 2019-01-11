import { Web } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Object Handler Base
 */
export class HandlerBase {
    public config: IProvisioningConfig = {};
    private name: string;

    /**
     * Creates a new instance of the ObjectHandlerBase class
     *
     * @param {string} name Name
     * @param {IProvisioningConfig} config Config
     */
    constructor(name: string, config: IProvisioningConfig = {}) {
        this.name = name;
        this.config = config;
    }

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
     * @param {any} data Data
     */
    public log_info(scope: string, message: string, data?: any) {
        let prefix = (this.config.logging && this.config.logging.prefix) ? `${this.config.logging.prefix} ` : "";
        Logger.log({ message: `${prefix}(${this.name}): ${scope}: ${message}`, data: data, level: LogLevel.Info });
    }
}
