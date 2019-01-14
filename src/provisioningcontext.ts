import {IContentType} from "./schema";

/**
 * Describes the Provisioning Context
 */
export class ProvisioningContext {
    public web = null;
    public lists: { [key: string]: string } = {};
    public siteFields: { [key: string]: string } = {};
    public contentTypes: { [key: string]: IContentType } = {};

    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor() { }
}
