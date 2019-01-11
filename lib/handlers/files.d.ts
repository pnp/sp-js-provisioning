import { HandlerBase } from "./handlerbase";
import { IFile } from "../schema";
import { Web } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Features Object Handler
 */
export declare class Files extends HandlerBase {
    private context;
    /**
     * Creates a new instance of the Files class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning Files
     *
     * @param {Web} web The web
     * @param {IFile[]} files The files  to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web: Web, files: IFile[], context: ProvisioningContext): Promise<void>;
    /**
     * Get blob for a file
     *
     * @param {IFile} file The file
     */
    private getFileBlob;
    /**
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The fileAddError
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     */
    private processFile;
    /**
     * Remove exisiting webparts if specified
     *
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     * @param {boolean} shouldRemove Should web parts be removed
     */
    private removeExistingWebParts;
    /**
     * Processes web parts
     *
     * @param {IFile} file The file
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processWebParts;
    /**
     * Fetches web part contents
     *
     * @param {Array<IWebPart>} webParts Web parts
     * @param {Function} callbackFunc Callback function that takes index of the the webpart and the retrieved XML
     */
    private fetchWebPartContents;
    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {Array<IWebPart>} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processPageListViews;
    /**
     * Processes page list view
     *
     * @param {Web} web The web
     * @param {any} listView List view
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processPageListView;
    /**
     * Process list item properties for the file
     *
     * @param {Web} web The web
     * @param {File} pnpFile The PnP file
     * @param {Object} properties The properties to set
     */
    private processProperties;
    /**
    * Replaces tokens in a string, e.g. {site}
    *
    * @param {string} str The string
    * @param {SP.ClientContext} ctx Client context
    */
    private replaceWebPartXmlTokens;
}
