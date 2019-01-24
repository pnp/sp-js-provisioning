import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Token Helper
 */
export declare class TokenHelper {
    context: ProvisioningContext;
    config: IProvisioningConfig;
    private tokenRegex;
    /**
     * Creates a new instance of the TokenHelper class
     */
    constructor(context: ProvisioningContext, config: IProvisioningConfig);
    replaceTokens(str: string): string;
}
