/**
 * Describes the Provisioning Context
 */
var ProvisioningContext = (function () {
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    function ProvisioningContext() {
        this.web = null;
        this.lists = {};
        this.siteFields = {};
        this.contentTypes = {};
    }
    return ProvisioningContext;
}());
export { ProvisioningContext };
