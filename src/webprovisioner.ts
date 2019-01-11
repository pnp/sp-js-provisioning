// we need to import HandlerBase & TypedHash to avoid naming issues in ts transpile
import { Schema } from "./schema";
import { HandlerBase } from "./handlers/handlerbase";
import { Web } from "@pnp/sp";
import { TypedHash } from "@pnp/common";
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";
import { DefaultHandlerMap, DefaultHandlerSort } from "./handlers/exports";
import { ProvisioningContext } from "./provisioningcontext";
import { IProvisioningConfig } from "./provisioningconfig";

/**
 * Root class of Provisioning
 */
export class WebProvisioner {
    public handlerMap: TypedHash<HandlerBase>;
    /**
     * Creates a new instance of the Provisioner class
     *
     * @param {Web} web The Web instance to which we want to apply templates
     * @param {TypedHash<HandlerBase>} handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    constructor(
        private web: Web,
        private config?: IProvisioningConfig,
        private context: ProvisioningContext = new ProvisioningContext(),
        public handlerSort: TypedHash<number> = DefaultHandlerSort) {
        this.handlerMap = DefaultHandlerMap(this.config);
    }

    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param {Schema} template The template to apply
     * @param {Function} progressCallback Callback for progress updates
     */
    public applyTemplate(template: Schema, progressCallback?: (msg: string) => void): Promise<void> {
        Logger.write(`Beginning processing of web [${this.web.toUrl()}]`, LogLevel.Info);

        let operations = Object.getOwnPropertyNames(template).sort((name1: string, name2: string) => {
            let sort1 = this.handlerSort.hasOwnProperty(name1) ? this.handlerSort[name1] : 99;
            let sort2 = this.handlerSort.hasOwnProperty(name2) ? this.handlerSort[name2] : 99;
            return sort1 - sort2;
        });

        return operations.reduce((chain, name) => {
            let handler = this.handlerMap[name];
            return chain.then(_ => {
                if (progressCallback) {
                    progressCallback(name);
                }
                return handler.ProvisionObjects(this.web, template[name], this.context);
            });
        }, Promise.resolve())
            .then(_ => {
                Logger.write(`Done processing of web [${this.web.toUrl()}]`, LogLevel.Info);
            })
            .catch(_ => {
                Logger.write(`Processing of web [${this.web.toUrl()}] failed`, LogLevel.Error);
            });
    }

    /**
    * Sets up the web provisioner
    *
    * @param {IProvisioningConfig} config Provisioning config
    */
    public setup(config: IProvisioningConfig) {
        this.config = config;
        if (this.config && this.config.activeLogLevel) {
            Logger.subscribe(new ConsoleListener());
            Logger.activeLogLevel = this.config.activeLogLevel;
        }
    }
}
