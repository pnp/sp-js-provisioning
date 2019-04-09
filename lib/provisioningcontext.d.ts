import { IContentType } from "./schema";
/**
 * Describes the Provisioning Context
 */
export declare class ProvisioningContext {
    web: any;
    lists: {
        [key: string]: string;
    };
    listViews: {
        [key: string]: string;
    };
    siteFields: {
        [key: string]: string;
    };
    contentTypes: {
        [key: string]: IContentType;
    };
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor();
}
