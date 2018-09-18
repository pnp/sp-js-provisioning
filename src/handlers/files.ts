import * as xmljs from "xml-js";
import { HandlerBase } from "./handlerbase";
import { IFile, IWebPart } from "../schema";
import { Web, File, Util, FileAddResult, Logger, LogLevel } from "sp-pnp-js";
import { replaceUrlTokens } from "../util";
import { ProvisioningContext } from "../provisioningcontext";

/**
 * Describes the Features Object Handler
 */
export class Files extends HandlerBase {
    private context: ProvisioningContext;

    /**
     * Creates a new instance of the Files class
     */
    constructor() {
        super("Files");
    }

    /**
     * Provisioning Files
     *
     * @param {Web} web The web
     * @param {IFile[]} files The files  to provision
     */
    public async ProvisionObjects(web: Web, files: IFile[], context: ProvisioningContext): Promise<void> {
        this.context = context;
        super.scope_started();
        if (typeof window === "undefined") {
            throw "Files Handler not supported in Node.";
        }
        const { ServerRelativeUrl } = await web.get();
        try {
            await files.reduce((chain, file) => chain.then(_ => this.processFile(web, file, ServerRelativeUrl)), Promise.resolve());
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
        }
    }

    /**
     * Get blob for a file
     *
     * @param {IFile} file The file
     */
    private async getFileBlob(file: IFile): Promise<Blob> {
        const fileSrcWithoutTokens = replaceUrlTokens(this.context.replaceTokens(file.Src));
        const response = await fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" });
        const fileContents = await response.text();
        const blob = new Blob([fileContents], { type: "text/plain" });
        return blob;
    }

    /**
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The file
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     */
    private async processFile(web: Web, file: IFile, webServerRelativeUrl: string): Promise<void> {
        Logger.log({ level: LogLevel.Info, message: `Processing file ${file.Folder}/${file.Url}` });
        try {
            const blob = await this.getFileBlob(file);
            const folderServerRelativeUrl = Util.combinePaths("/", webServerRelativeUrl, file.Folder);
            const pnpFolder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
            let fileServerRelativeUrl = Util.combinePaths("/", folderServerRelativeUrl, file.Url);
            let fileAddResult: FileAddResult;
            let pnpFile: File;
            try {
                fileAddResult = await pnpFolder.files.add(file.Url, blob, file.Overwrite);
                pnpFile = fileAddResult.file;
                fileServerRelativeUrl = fileAddResult.data.ServerRelativeUrl;
            } catch (fileAddError) {
                pnpFile = web.getFileByServerRelativePath(fileServerRelativeUrl);
            }
            await this.processProperties(web, pnpFile, file);
            await this.processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl);
            await this.processPageListViews(web, file.WebParts, fileServerRelativeUrl);
        } catch (err) {
            throw err;
        }

    }

    /**
     * Remove exisiting webparts if specified
     *
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     * @param {boolean} shouldRemove Should web parts be removed
     */
    private removeExistingWebParts(webServerRelativeUrl: string, fileServerRelativeUrl: string, shouldRemove: boolean) {
        return new Promise((resolve, reject) => {
            if (shouldRemove) {
                Logger.log({ level: LogLevel.Info, message: `Deleting existing webpart from file ${fileServerRelativeUrl}` });
                const clientContext = new SP.ClientContext(webServerRelativeUrl);
                const spFile = clientContext.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl);
                const webPartManager = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                const webParts = webPartManager.get_webParts();
                clientContext.load(webParts);
                clientContext.executeQueryAsync(() => {
                    webParts.get_data().forEach(wp => wp.deleteWebPart());
                    clientContext.executeQueryAsync(resolve, reject);
                }, reject);
            } else {
                Logger.log({ level: LogLevel.Info, message: `Web parts should not be removed from file ${fileServerRelativeUrl}.` });
                resolve();
            }
        });
    }

    /**
     * Processes web parts
     *
     * @param {IFile} file The file
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processWebParts(file: IFile, webServerRelativeUrl: string, fileServerRelativeUrl: string) {
        return new Promise(async (resolve, reject) => {
            Logger.log({ level: LogLevel.Info, message: `Processing webparts for file ${file.Folder}/${file.Url}` });
            await this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts);
            if (file.WebParts && file.WebParts.length > 0) {
                let clientContext = new SP.ClientContext(webServerRelativeUrl),
                    spFile = clientContext.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                    webPartManager = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                await this.fetchWebPartContents(file.WebParts, (index, xml) => { file.WebParts[index].Contents.Xml = xml; });
                file.WebParts.forEach(wp => {
                    const webPartXml = this.context.replaceTokens(this.replaceWebPartXmlTokens(wp.Contents.Xml, clientContext));
                    const webPartDef = webPartManager.importWebPart(webPartXml);
                    const webPartInstance = webPartDef.get_webPart();
                    Logger.log({ data: { webPartXml }, level: LogLevel.Info, message: `Processing webpart ${wp.Title} for file ${file.Folder}/${file.Url}` });
                    webPartManager.addWebPart(webPartInstance, wp.Zone, wp.Order);
                    clientContext.load(webPartInstance);
                });
                clientContext.executeQueryAsync(resolve, (sender, args) => {
                    Logger.log({
                        data: { error: args.get_message() },
                        level: LogLevel.Error,
                        message: `Failed to process webparts for file ${file.Folder}/${file.Url}`,
                    });
                    reject({ sender, args });
                });
            } else {
                resolve();
            }
        });
    }


    /**
     * Fetches web part contents
     *
     * @param {Array<IWebPart>} webParts Web parts
     * @param {Function} callbackFunc Callback function that takes index of the the webpart and the retrieved XML
     */
    private fetchWebPartContents = (webParts: Array<IWebPart>, callbackFunc: (index, xml) => void) => {
        return new Promise<any>((resolve, reject) => {
            let fileFetchPromises = webParts.map((wp, index) => {
                return (() => {
                    return new Promise<any>(async (_res, _rej) => {
                        if (wp.Contents.FileSrc) {
                            const fileSrc = replaceUrlTokens(this.context.replaceTokens(wp.Contents.FileSrc));
                            Logger.log({ data: null, level: LogLevel.Info, message: `Retrieving contents from file '${fileSrc}'.` });
                            const response = await fetch(fileSrc, { credentials: "include", method: "GET" });
                            const xml = await response.text();
                            if (Util.isArray(wp.PropertyOverrides)) {
                                let obj: any = xmljs.xml2js(xml);
                                if (obj.elements[0].name === "webParts") {
                                    const existingProperties = obj.elements[0].elements[0].elements[1].elements[0].elements;
                                    let updatedProperties = [];
                                    existingProperties.forEach(prop => {
                                        let hasOverride = wp.PropertyOverrides.filter(po => po.name === prop.attributes.name).length > 0;
                                        if (!hasOverride) {
                                            updatedProperties.push(prop);
                                        }
                                    });
                                    wp.PropertyOverrides.forEach(({ name, type, value }) => {
                                        updatedProperties.push({ attributes: { name, type }, elements: [{ text: value, type: "text" }], name: "property", type: "element" });
                                    });
                                    obj.elements[0].elements[0].elements[1].elements[0].elements = updatedProperties;
                                    callbackFunc(index, xmljs.js2xml(obj));
                                    _res();
                                } else {
                                    callbackFunc(index, xml);
                                    _res();
                                }
                            } else {
                                callbackFunc(index, xml);
                                _res();
                            }
                        } else {
                            _res();
                        }
                    });
                })();
            });
            Promise.all(fileFetchPromises)
                .then(resolve)
                .catch(reject);
        });
    }

    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {Array<IWebPart>} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processPageListViews(web: Web, webParts: Array<IWebPart>, fileServerRelativeUrl: string): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            if (webParts) {
                Logger.log({
                    data: { webParts, fileServerRelativeUrl },
                    level: LogLevel.Info,
                    message: `Processing page list views for file ${fileServerRelativeUrl}`,
                });
                let listViewWebParts = webParts.filter(wp => wp.ListView);
                if (listViewWebParts.length > 0) {
                    listViewWebParts
                        .reduce((chain, wp) => chain.then(_ => this.processPageListView(web, wp.ListView, fileServerRelativeUrl)), Promise.resolve())
                        .then(() => {
                            Logger.log({
                                data: {},
                                level: LogLevel.Info,
                                message: `Successfully processed page list views for file ${fileServerRelativeUrl}`,
                            });
                            resolve();
                        })
                        .catch(err => {
                            Logger.log({
                                data: { err, fileServerRelativeUrl },
                                level: LogLevel.Error,
                                message: `Failed to process page list views for file ${fileServerRelativeUrl}`,
                            });
                            reject(err);
                        });
                } else {
                    resolve();
                }
            } else {
                resolve();
            }
        });
    }

    /**
     * Processes page list view
     *
     * @param {Web} web The web
     * @param {any} listView List view
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processPageListView(web: Web, listView, fileServerRelativeUrl: string) {
        return new Promise<void>((resolve, reject) => {
            let views = web.lists.getByTitle(listView.List).views;
            views.get()
                .then(listViews => {
                    let wpView = listViews.filter(v => v.ServerRelativeUrl === fileServerRelativeUrl);
                    if (wpView.length === 1) {
                        let view = views.getById(wpView[0].Id);
                        let settings = listView.View.AdditionalSettings || {};
                        view.update(settings)
                            .then(() => {
                                view.fields.removeAll()
                                    .then(_ => {
                                        listView.View.ViewFields.reduce((chain, viewField) => chain.then(() => view.fields.add(viewField)), Promise.resolve())
                                            .then(resolve)
                                            .catch(err => {
                                                Logger.log({
                                                    data: { fileServerRelativeUrl, listView, err },
                                                    level: LogLevel.Error,
                                                    message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                                                });
                                                reject(err);
                                            });
                                    })
                                    .catch(err => {
                                        Logger.log({
                                            data: { fileServerRelativeUrl, listView, err },
                                            level: LogLevel.Error,
                                            message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                                        });
                                        reject(err);
                                    });
                            })
                            .catch(err => {
                                Logger.log({
                                    data: { fileServerRelativeUrl, listView, err },
                                    level: LogLevel.Error,
                                    message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                                });
                                reject(err);
                            });
                    } else {
                        resolve();
                    }
                })
                .catch(err => {
                    Logger.log({
                        data: { fileServerRelativeUrl, listView, err },
                        level: LogLevel.Error,
                        message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                    });
                    reject(err);
                });
        });
    }

    /**
     * Process list item properties for the file
     *
     * @param {Web} web The web
     * @param {File} pnpFile The PnP file
     * @param {Object} properties The properties to set
     */
    private async processProperties(web: Web, pnpFile: File, file: IFile) {
        const hasProperties = file.Properties && Object.keys(file.Properties).length > 0;
        if (hasProperties) {
            Logger.log({ level: LogLevel.Info, message: `Processing properties for ${file.Folder}/${file.Url}` });
            const listItemAllFields = await pnpFile.listItemAllFields.select("ID", "ParentList/ID", "ParentList/Title").expand("ParentList").get();
            await web.lists.getById(listItemAllFields.ParentList.Id).items.getById(listItemAllFields.ID).update(file.Properties);
            Logger.log({ level: LogLevel.Info, message: `Successfully processed properties for ${file.Folder}/${file.Url}` });
        }
    }

    /**
    * Replaces tokens in a string, e.g. {site}
    *
    * @param {string} str The string
    * @param {SP.ClientContext} ctx Client context
    */
    private replaceWebPartXmlTokens(str: string, ctx: SP.ClientContext): string {
        let site = Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    }
}
