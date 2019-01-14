/**
 * Describes the Provisioning Context
 */
export declare class ProvisioningContext {
    web: any;
    lists: {
        [key: string]: string;
    };
    siteFields: {
        [key: string]: string;
    };
    contentTypes: {
        [key: string]: string;
    };
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor();
}
