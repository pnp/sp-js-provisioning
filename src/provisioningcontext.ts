/**
 * Describes the Provisioning Context
 */
export class ProvisioningContext {
    public web = null;
    public lists: { [key: string]: string } = {};

    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor() { }
}
