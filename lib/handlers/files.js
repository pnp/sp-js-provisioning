"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const xmljs = require("xml-js");
const handlerbase_1 = require("./handlerbase");
const common_1 = require("@pnp/common");
const logging_1 = require("@pnp/logging");
const util_1 = require("../util");
/**
 * Describes the Features Object Handler
 */
class Files extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Files class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config) {
        super("Files", config);
        /**
         * Fetches web part contents
         *
         * @param {Array<IWebPart>} webParts Web parts
         * @param {Function} callbackFunc Callback function that takes index of the the webpart and the retrieved XML
         */
        this.fetchWebPartContents = (webParts, callbackFunc) => {
            return new Promise((resolve, reject) => {
                let fileFetchPromises = webParts.map((wp, index) => {
                    return (() => {
                        return new Promise((_res, _rej) => __awaiter(this, void 0, void 0, function* () {
                            if (wp.Contents.FileSrc) {
                                const fileSrc = util_1.replaceUrlTokens(this.context.replaceTokens(wp.Contents.FileSrc), this.config);
                                logging_1.Logger.log({ data: null, level: 1 /* Info */, message: `Retrieving contents from file '${fileSrc}'.` });
                                const response = yield fetch(fileSrc, { credentials: "include", method: "GET" });
                                const xml = yield response.text();
                                if (common_1.isArray(wp.PropertyOverrides)) {
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
                                            updatedProperties.push({ attributes: { name, type }, elements: [{ text: value, type: "text" }], name: "property", type: "element" });
                                        });
                                        obj.elements[0].elements[0].elements[1].elements[0].elements = updatedProperties;
                                        callbackFunc(index, xmljs.js2xml(obj));
                                        _res();
                                    }
                                    else {
                                        callbackFunc(index, xml);
                                        _res();
                                    }
                                }
                                else {
                                    callbackFunc(index, xml);
                                    _res();
                                }
                            }
                            else {
                                _res();
                            }
                        }));
                    })();
                });
                Promise.all(fileFetchPromises)
                    .then(resolve)
                    .catch(reject);
            });
        };
    }
    /**
     * Provisioning Files
     *
     * @param {Web} web The web
     * @param {IFile[]} files The files  to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web, files, context) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            this.context = context;
            _super.scope_started.call(this);
            if (typeof window === "undefined") {
                throw "Files Handler not supported in Node.";
            }
            if (this.config.spfxContext) {
                throw "Files Handler not supported in SPFx.";
            }
            const { ServerRelativeUrl } = yield web.get();
            try {
                yield files.reduce((chain, file) => chain.then(_ => this.processFile(web, file, ServerRelativeUrl)), Promise.resolve());
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
            }
        });
    }
    /**
     * Get blob for a file
     *
     * @param {IFile} file The file
     */
    getFileBlob(file) {
        return __awaiter(this, void 0, void 0, function* () {
            const fileSrcWithoutTokens = util_1.replaceUrlTokens(this.context.replaceTokens(file.Src), this.config);
            const response = yield fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" });
            const fileContents = yield response.text();
            const blob = new Blob([fileContents], { type: "text/plain" });
            return blob;
        });
    }
    /**
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The fileAddError
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     */
    processFile(web, file, webServerRelativeUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            logging_1.Logger.log({ level: 1 /* Info */, message: `Processing file ${file.Folder}/${file.Url}` });
            try {
                const blob = yield this.getFileBlob(file);
                const folderServerRelativeUrl = common_1.combine("/", webServerRelativeUrl, file.Folder);
                const pnpFolder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
                let fileServerRelativeUrl = common_1.combine("/", folderServerRelativeUrl, file.Url);
                let fileAddResult;
                let pnpFile;
                try {
                    fileAddResult = yield pnpFolder.files.add(file.Url, blob, file.Overwrite);
                    pnpFile = fileAddResult.file;
                    fileServerRelativeUrl = fileAddResult.data.ServerRelativeUrl;
                }
                catch (fileAddError) {
                    pnpFile = web.getFileByServerRelativePath(fileServerRelativeUrl);
                }
                yield this.processProperties(web, pnpFile, file);
                yield this.processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl);
                yield this.processPageListViews(web, file.WebParts, fileServerRelativeUrl);
            }
            catch (err) {
                throw err;
            }
        });
    }
    /**
     * Remove exisiting webparts if specified
     *
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     * @param {boolean} shouldRemove Should web parts be removed
     */
    removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, shouldRemove) {
        return new Promise((resolve, reject) => {
            if (shouldRemove) {
                logging_1.Logger.log({ level: 1 /* Info */, message: `Deleting existing webpart from file ${fileServerRelativeUrl}` });
                const clientContext = new SP.ClientContext(webServerRelativeUrl);
                const spFile = clientContext.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl);
                const webPartManager = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                const webParts = webPartManager.get_webParts();
                clientContext.load(webParts);
                clientContext.executeQueryAsync(() => {
                    webParts.get_data().forEach(wp => wp.deleteWebPart());
                    clientContext.executeQueryAsync(resolve, reject);
                }, reject);
            }
            else {
                logging_1.Logger.log({ level: 1 /* Info */, message: `Web parts should not be removed from file ${fileServerRelativeUrl}.` });
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
    processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => __awaiter(this, void 0, void 0, function* () {
            logging_1.Logger.log({ level: 1 /* Info */, message: `Processing webparts for file ${file.Folder}/${file.Url}` });
            yield this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts);
            if (file.WebParts && file.WebParts.length > 0) {
                let clientContext = new SP.ClientContext(webServerRelativeUrl), spFile = clientContext.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), webPartManager = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                yield this.fetchWebPartContents(file.WebParts, (index, xml) => { file.WebParts[index].Contents.Xml = xml; });
                file.WebParts.forEach(wp => {
                    const webPartXml = this.context.replaceTokens(this.replaceWebPartXmlTokens(wp.Contents.Xml, clientContext));
                    const webPartDef = webPartManager.importWebPart(webPartXml);
                    const webPartInstance = webPartDef.get_webPart();
                    logging_1.Logger.log({ data: { webPartXml }, level: 1 /* Info */, message: `Processing webpart ${wp.Title} for file ${file.Folder}/${file.Url}` });
                    webPartManager.addWebPart(webPartInstance, wp.Zone, wp.Order);
                    clientContext.load(webPartInstance);
                });
                clientContext.executeQueryAsync(resolve, (sender, args) => {
                    logging_1.Logger.log({
                        data: { error: args.get_message() },
                        level: 3 /* Error */,
                        message: `Failed to process webparts for file ${file.Folder}/${file.Url}`,
                    });
                    reject({ sender, args });
                });
            }
            else {
                resolve();
            }
        }));
    }
    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {Array<IWebPart>} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    processPageListViews(web, webParts, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            if (webParts) {
                logging_1.Logger.log({
                    data: { webParts, fileServerRelativeUrl },
                    level: 1 /* Info */,
                    message: `Processing page list views for file ${fileServerRelativeUrl}`,
                });
                let listViewWebParts = webParts.filter(wp => wp.ListView);
                if (listViewWebParts.length > 0) {
                    listViewWebParts
                        .reduce((chain, wp) => chain.then(_ => this.processPageListView(web, wp.ListView, fileServerRelativeUrl)), Promise.resolve())
                        .then(() => {
                        logging_1.Logger.log({
                            data: {},
                            level: 1 /* Info */,
                            message: `Successfully processed page list views for file ${fileServerRelativeUrl}`,
                        });
                        resolve();
                    })
                        .catch(err => {
                        logging_1.Logger.log({
                            data: { err, fileServerRelativeUrl },
                            level: 3 /* Error */,
                            message: `Failed to process page list views for file ${fileServerRelativeUrl}`,
                        });
                        reject(err);
                    });
                }
                else {
                    resolve();
                }
            }
            else {
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
    processPageListView(web, listView, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
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
                                logging_1.Logger.log({
                                    data: { fileServerRelativeUrl, listView, err },
                                    level: 3 /* Error */,
                                    message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                                });
                                reject(err);
                            });
                        })
                            .catch(err => {
                            logging_1.Logger.log({
                                data: { fileServerRelativeUrl, listView, err },
                                level: 3 /* Error */,
                                message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                            });
                            reject(err);
                        });
                    })
                        .catch(err => {
                        logging_1.Logger.log({
                            data: { fileServerRelativeUrl, listView, err },
                            level: 3 /* Error */,
                            message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                        });
                        reject(err);
                    });
                }
                else {
                    resolve();
                }
            })
                .catch(err => {
                logging_1.Logger.log({
                    data: { fileServerRelativeUrl, listView, err },
                    level: 3 /* Error */,
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
    processProperties(web, pnpFile, file) {
        return __awaiter(this, void 0, void 0, function* () {
            const hasProperties = file.Properties && Object.keys(file.Properties).length > 0;
            if (hasProperties) {
                logging_1.Logger.log({ level: 1 /* Info */, message: `Processing properties for ${file.Folder}/${file.Url}` });
                const listItemAllFields = yield pnpFile.listItemAllFields.select("ID", "ParentList/ID", "ParentList/Title").expand("ParentList").get();
                yield web.lists.getById(listItemAllFields.ParentList.Id).items.getById(listItemAllFields.ID).update(file.Properties);
                logging_1.Logger.log({ level: 1 /* Info */, message: `Successfully processed properties for ${file.Folder}/${file.Url}` });
            }
        });
    }
    /**
    * Replaces tokens in a string, e.g. {site}
    *
    * @param {string} str The string
    * @param {SP.ClientContext} ctx Client context
    */
    replaceWebPartXmlTokens(str, ctx) {
        let site = common_1.combine(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    }
}
exports.Files = Files;
