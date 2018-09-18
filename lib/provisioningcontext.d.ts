/**
 * Describes the Provisioning Context
 */
export declare class ProvisioningContext {
    lists: Array<any>;
    tokenRegex: RegExp;
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor(lists?: Array<any>, tokenRegex?: RegExp);
    replaceTokens(str: string): string;
}
