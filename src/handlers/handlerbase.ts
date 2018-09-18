import { Web, Logger, LogLevel } from "sp-pnp-js";
import { ProvisioningContext } from "../provisioningcontext";

/**
 * Describes the Object Handler Base
 */
export class HandlerBase {

    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(private name: string) {}

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
        Logger.write(`${this.name}: Code execution scope started`);
    }

    /**
     * Writes to Logger when scope has stopped
     */
    public scope_ended() {
        Logger.write(`${this.name}: Code execution scope stopped`);
    }
}
