import * as xmljs from "xml-js";
import { HandlerBase } from "./handlerbase";
import { IFile, IWebPart } from "../schema";
import { Web, Util, FileAddResult, Logger, LogLevel } from "sp-pnp-js";
import { ReplaceTokens } from "../util";

/**
 * Describes the Features Object Handler
 */
export class Files extends HandlerBase {
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
    public async ProvisionObjects(web: Web, files: IFile[]): Promise<void> {
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
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The file
     * @param {string} serverRelativeUrl ServerRelativeUrl for the web
     */
    private async processFile(web: Web, file: IFile, serverRelativeUrl: string): Promise<void> {
        Logger.log({ data: file, level: LogLevel.Info, message: `Processing file ${file.Folder}/${file.Url}` });
        const fileSrcWithoutTokens = ReplaceTokens(file.Src);
        try {
            const response = await fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" })
            const fileContents = await response.text();
            const blob = new Blob([fileContents], { type: "text/plain" });
            const folderServerRelativeUrl = Util.combinePaths("/", serverRelativeUrl, file.Folder);
            const folder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
            const result = await folder.files.add(file.Url, blob, file.Overwrite);
            await Promise.all([
                this.processWebParts(file, serverRelativeUrl, result.data.ServerRelativeUrl),
                this.processProperties(web, result, file.Properties),
            ])
            await this.processPageListViews(web, file.WebParts, result.data.ServerRelativeUrl);
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
                Logger.log({
                    data: { webServerRelativeUrl, fileServerRelativeUrl },
                    level: LogLevel.Info,
                    message: `Deleting existing webpart from file ${fileServerRelativeUrl}`,
                });
                let ctx = new SP.ClientContext(webServerRelativeUrl),
                    spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                    lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared),
                    webParts = lwpm.get_webParts();
                ctx.load(webParts);
                ctx.executeQueryAsync(() => {
                    webParts.get_data().forEach(wp => wp.deleteWebPart());
                    ctx.executeQueryAsync(resolve, reject);
                }, reject);
            } else {
                Logger.log({
                    data: { webServerRelativeUrl, fileServerRelativeUrl },
                    level: LogLevel.Info,
                    message: `Web parts should not be removed from file ${fileServerRelativeUrl}. Resolving.`,
                });
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
        return new Promise((resolve, reject) => {
            Logger.log({ data: file.WebParts, level: LogLevel.Info, message: `Processing webparts for file ${file.Folder}/${file.Url}` });
            this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts).then(() => {
                if (file.WebParts && file.WebParts.length > 0) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl),
                        spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl),
                        lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                    this.fetchWebPartContents(file.WebParts, (index, xml) => {
                        file.WebParts[index].Contents.Xml = xml;
                    })
                        .then(() => {
                            file.WebParts.forEach(wp => {
                                let def = lwpm.importWebPart(this.replaceXmlTokens(wp.Contents.Xml, ctx)),
                                    inst = def.get_webPart();
                                Logger.log({ data: wp, level: LogLevel.Info, message: `Processing webpart ${wp.Title} for file ${file.Folder}/${file.Url}` });
                                lwpm.addWebPart(inst, wp.Zone, wp.Order);
                                ctx.load(inst);
                            });
                            ctx.executeQueryAsync(resolve, (sender, args) => {
                                Logger.log({
                                    data: { error: args.get_message() },
                                    level: LogLevel.Error,
                                    message: `Failed to process webparts for file ${file.Folder}/${file.Url}`,
                                });
                                reject({ sender, args });
                            });
                        })
                        .catch(reject);
                } else {
                    resolve();
                }
            }, reject);
        });
    }

    /**
     * Fetches web part contents
     *
     * @param {IWebPart[]} webParts Web parts
     * @param {Function} cb Callback function that takes index of the the webpart and the retrieved XML
     */
    private fetchWebPartContents = (webParts: IWebPart[], cb: (index, xml) => void) => new Promise<any>((resolve, reject) => {
        let fileFetchPromises = webParts.map((wp, index) => {
            return (() => {
                return new Promise<any>((_res, _rej) => {
                    if (wp.Contents.FileSrc) {
                        const fileSrc = ReplaceTokens(wp.Contents.FileSrc);
                        Logger.log({ data: null, level: LogLevel.Info, message: `Retrieving contents from file '${fileSrc}'.` });
                        fetch(fileSrc, { credentials: "include", method: "GET" })
                            .then(res => {
                                res.text()
                                    .then(xml => {
                                        if (Util.isArray(wp.PropertyOverrides)) {
                                            let obj = xmljs.xml2js(xml);
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
                                                    updatedProperties.push({
                                                        attributes: {
                                                            name,
                                                            type,
                                                        },
                                                        elements: [
                                                            {
                                                                text: value,
                                                                type: "text",
                                                            },
                                                        ],
                                                        name: "property",
                                                        type: "element",
                                                    });
                                                });
                                                obj.elements[0].elements[0].elements[1].elements[0].elements = updatedProperties;
                                                cb(index, xmljs.js2xml(obj));
                                                _res();
                                            } else {
                                                cb(index, xml);
                                                _res();
                                            }
                                        } else {
                                            cb(index, xml);
                                            _res();
                                        }
                                    })
                                    .catch(err => {
                                        Logger.log({ data: err, level: LogLevel.Error, message: `Failed to retrieve contents from file '${fileSrc}'.` });
                                        reject(err);
                                    });
                            })
                            .catch(err => {
                                Logger.log({ data: err, level: LogLevel.Error, message: `Failed to retrieve contents from file '${fileSrc}'.` });
                                reject(err);
                            });
                    } else {
                        _res();
                    }
                });
            })();
        });
        Promise.all(fileFetchPromises)
            .then(resolve)
            .catch(reject);
    })

    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {IWebPart[]} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    private processPageListViews(web: Web, webParts: IWebPart[], fileServerRelativeUrl: string): Promise<void> {
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
     * @param {FileAddResuylt} result The file add result
     * @param {Object} properties The properties to set
     */
    private processProperties(web: Web, result: FileAddResult, properties: { [key: string]: string | number }) {
        return new Promise((resolve, reject) => {
            if (properties && Object.keys(properties).length > 0) {
                result.file.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get()
                    .then(({ ID, ParentList }) => {
                        web.lists.getById(ParentList.Id).items.getById(ID).update(properties)
                            .then(resolve)
                            .catch(reject);
                    })
                    .catch(reject);
            } else {
                resolve();
            }
        });
    }

    /**
     * Replaces tokens in a string, e.g. {site}
     *
     * @param {string} str The string
     * @param {SP.ClientContext} ctx Client context
     */
    private replaceXmlTokens(str: string, ctx: SP.ClientContext): string {
        let site = Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    }
}
