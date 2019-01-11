/**
 * Describes the Provisioning Context
 */
export declare class ProvisioningContext {
    lists: {
        [key: string]: string;
    };
    tokenRegex: RegExp;
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor(lists?: {
        [key: string]: string;
    }, tokenRegex?: RegExp);
    replaceTokens(str: string): string;
}
