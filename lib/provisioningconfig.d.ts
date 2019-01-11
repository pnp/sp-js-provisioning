import { LogLevel } from "@pnp/logging";
export declare type ProvisioningParameters = {
    [key: string]: string;
};
export interface IProvisioningLogging {
    activeLogLevel?: LogLevel;
    prefix?: string;
}
/**
 * Describes the Provisioning Config
 */
export interface IProvisioningConfig {
    parameters?: ProvisioningParameters;
    spfxContext?: any;
    logging?: IProvisioningLogging;
}
