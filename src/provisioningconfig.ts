import { LogLevel } from "@pnp/logging";

/**
 * Describes the Provisioning Config
 */
export interface IProvisioningConfig {
    spfxContext?;
    activeLogLevel?: LogLevel;
}
