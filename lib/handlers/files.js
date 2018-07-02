"use strict";
<<<<<<< HEAD
=======
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
Object.defineProperty(exports, "__esModule", { value: true });
const xmljs = require("xml-js");
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
const util_1 = require("../util");
/**
 * Describes the Features Object Handler
 */
class Files extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Files class
     */
    constructor() {
        super("Files");
        /**
         * Fetches web part contents
         *
         * @param {IWebPart[]} webParts Web parts
         * @param {Function} cb Callback function that takes index of the the webpart and the retrieved XML
         */
        this.fetchWebPartContents = (webParts, cb) => new Promise((resolve, reject) => {
            let fileFetchPromises = webParts.map((wp, index) => {
                return (() => {
<<<<<<< HEAD
                    return new Promise((_res, _rej) => {
                        if (wp.Contents.FileSrc) {
                            const fileSrc = util_1.ReplaceTokens(wp.Contents.FileSrc);
                            sp_pnp_js_1.Logger.log({ data: null, level: sp_pnp_js_1.LogLevel.Info, message: `Retrieving contents from file '${fileSrc}'.` });
                            fetch(fileSrc, { credentials: "include", method: "GET" })
                                .then(res => {
                                res.text()
                                    .then(xml => {
                                    if (sp_pnp_js_1.Util.isArray(wp.PropertyOverrides)) {
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
                                        }
                                        else {
                                            cb(index, xml);
                                            _res();
                                        }
                                    }
                                    else {
                                        cb(index, xml);
                                        _res();
                                    }
                                })
                                    .catch(err => {
                                    sp_pnp_js_1.Logger.log({ data: err, level: sp_pnp_js_1.LogLevel.Error, message: `Failed to retrieve contents from file '${fileSrc}'.` });
                                    reject(err);
                                });
                            })
                                .catch(err => {
                                sp_pnp_js_1.Logger.log({ data: err, level: sp_pnp_js_1.LogLevel.Error, message: `Failed to retrieve contents from file '${fileSrc}'.` });
                                reject(err);
                            });
=======
                    return new Promise((_res, _rej) => __awaiter(this, void 0, void 0, function* () {
                        if (wp.Contents.FileSrc) {
                            const fileSrc = util_1.ReplaceTokens(wp.Contents.FileSrc);
                            sp_pnp_js_1.Logger.log({ data: null, level: sp_pnp_js_1.LogLevel.Info, message: `Retrieving contents from file '${fileSrc}'.` });
                            const response = yield fetch(fileSrc, { credentials: "include", method: "GET" });
                            const xml = yield response.text();
                            if (sp_pnp_js_1.Util.isArray(wp.PropertyOverrides)) {
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
                                }
                                else {
                                    cb(index, xml);
                                    _res();
                                }
                            }
                            else {
                                cb(index, xml);
                                _res();
                            }
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
                        }
                        else {
                            _res();
                        }
<<<<<<< HEAD
                    });
=======
                    }));
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
                })();
            });
            Promise.all(fileFetchPromises)
                .then(resolve)
                .catch(reject);
        });
    }
    /**
     * Provisioning Files
     *
     * @param {Web} web The web
     * @param {IFile[]} files The files  to provision
     */
    ProvisionObjects(web, files) {
<<<<<<< HEAD
        super.scope_started();
        return new Promise((resolve, reject) => {
            if (typeof window === "undefined") {
                reject("Files Handler not supported in Node.");
            }
            web.get().then(({ ServerRelativeUrl }) => {
                files.reduce((chain, file) => chain.then(_ => this.processFile(web, file, ServerRelativeUrl)), Promise.resolve()).then(() => {
                    super.scope_ended();
                    resolve();
                }).catch(e => {
                    super.scope_ended();
                    reject(e);
                });
            });
=======
        const _super = name => super[name];
        return __awaiter(this, void 0, void 0, function* () {
            _super("scope_started").call(this);
            if (typeof window === "undefined") {
                throw "Files Handler not supported in Node.";
            }
            const { ServerRelativeUrl } = yield web.get();
            try {
                yield files.reduce((chain, file) => chain.then(_ => this.processFile(web, file, ServerRelativeUrl)), Promise.resolve());
                _super("scope_ended").call(this);
            }
            catch (err) {
                _super("scope_ended").call(this);
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
            const fileSrcWithoutTokens = util_1.ReplaceTokens(file.Src);
            const response = yield fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" });
            const fileContents = yield response.text();
            const blob = new Blob([fileContents], { type: "text/plain" });
            return blob;
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
        });
    }
    /**
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The file
<<<<<<< HEAD
     * @param {string} serverRelativeUrl ServerRelativeUrl for the web
     */
    processFile(web, file, serverRelativeUrl) {
        return new Promise((resolve, reject) => {
            sp_pnp_js_1.Logger.log({ data: file, level: sp_pnp_js_1.LogLevel.Info, message: `Processing file ${file.Folder}/${file.Url}` });
            const fileSrcWithoutTokens = util_1.ReplaceTokens(file.Src);
            fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" })
                .then(res => {
                res.text()
                    .then(fileContents => {
                    const blob = new Blob([fileContents], {
                        type: "text/plain",
                    });
                    const folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("/", serverRelativeUrl, file.Folder);
                    const folder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
                    folder.files.add(file.Url, blob, file.Overwrite)
                        .then(result => {
                        Promise.all([
                            this.processWebParts(file, serverRelativeUrl, result.data.ServerRelativeUrl),
                            this.processProperties(web, result, file.Properties),
                        ])
                            .then(_ => {
                            this.processPageListViews(web, file.WebParts, result.data.ServerRelativeUrl)
                                .then(resolve)
                                .catch(reject);
                        })
                            .catch(reject);
                    })
                        .catch(reject);
                })
                    .catch(reject);
            })
                .catch(reject);
=======
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     */
    processFile(web, file, webServerRelativeUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            sp_pnp_js_1.Logger.log({ data: file, level: sp_pnp_js_1.LogLevel.Info, message: `Processing file ${file.Folder}/${file.Url}` });
            try {
                const blob = yield this.getFileBlob(file);
                const folderServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("/", webServerRelativeUrl, file.Folder);
                const pnpFolder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
                let fileServerRelativeUrl = sp_pnp_js_1.Util.combinePaths("/", folderServerRelativeUrl, file.Url);
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
                yield Promise.all([
                    this.processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl),
                    this.processProperties(web, pnpFile, file.Properties),
                ]);
                yield this.processPageListViews(web, file.WebParts, fileServerRelativeUrl);
            }
            catch (err) {
                throw err;
            }
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
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
                sp_pnp_js_1.Logger.log({
                    data: { webServerRelativeUrl, fileServerRelativeUrl },
                    level: sp_pnp_js_1.LogLevel.Info,
                    message: `Deleting existing webpart from file ${fileServerRelativeUrl}`,
                });
                let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared), webParts = lwpm.get_webParts();
                ctx.load(webParts);
                ctx.executeQueryAsync(() => {
                    webParts.get_data().forEach(wp => wp.deleteWebPart());
                    ctx.executeQueryAsync(resolve, reject);
                }, reject);
            }
            else {
                sp_pnp_js_1.Logger.log({
                    data: { webServerRelativeUrl, fileServerRelativeUrl },
                    level: sp_pnp_js_1.LogLevel.Info,
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
    processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            sp_pnp_js_1.Logger.log({ data: file.WebParts, level: sp_pnp_js_1.LogLevel.Info, message: `Processing webparts for file ${file.Folder}/${file.Url}` });
            this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts).then(() => {
                if (file.WebParts && file.WebParts.length > 0) {
                    let ctx = new SP.ClientContext(webServerRelativeUrl), spFile = ctx.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), lwpm = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                    this.fetchWebPartContents(file.WebParts, (index, xml) => {
                        file.WebParts[index].Contents.Xml = xml;
                    })
                        .then(() => {
                        file.WebParts.forEach(wp => {
<<<<<<< HEAD
                            let def = lwpm.importWebPart(this.replaceXmlTokens(wp.Contents.Xml, ctx)), inst = def.get_webPart();
=======
                            let def = lwpm.importWebPart(util_1.ReplaceTokens(wp.Contents.Xml)), inst = def.get_webPart();
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
                            sp_pnp_js_1.Logger.log({ data: wp, level: sp_pnp_js_1.LogLevel.Info, message: `Processing webpart ${wp.Title} for file ${file.Folder}/${file.Url}` });
                            lwpm.addWebPart(inst, wp.Zone, wp.Order);
                            ctx.load(inst);
                        });
                        ctx.executeQueryAsync(resolve, (sender, args) => {
                            sp_pnp_js_1.Logger.log({
                                data: { error: args.get_message() },
                                level: sp_pnp_js_1.LogLevel.Error,
                                message: `Failed to process webparts for file ${file.Folder}/${file.Url}`,
                            });
                            reject({ sender, args });
                        });
                    })
                        .catch(reject);
                }
                else {
                    resolve();
                }
            }, reject);
        });
    }
    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {IWebPart[]} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    processPageListViews(web, webParts, fileServerRelativeUrl) {
        return new Promise((resolve, reject) => {
            if (webParts) {
                sp_pnp_js_1.Logger.log({
                    data: { webParts, fileServerRelativeUrl },
                    level: sp_pnp_js_1.LogLevel.Info,
                    message: `Processing page list views for file ${fileServerRelativeUrl}`,
                });
                let listViewWebParts = webParts.filter(wp => wp.ListView);
                if (listViewWebParts.length > 0) {
                    listViewWebParts
                        .reduce((chain, wp) => chain.then(_ => this.processPageListView(web, wp.ListView, fileServerRelativeUrl)), Promise.resolve())
                        .then(() => {
                        sp_pnp_js_1.Logger.log({
                            data: {},
                            level: sp_pnp_js_1.LogLevel.Info,
                            message: `Successfully processed page list views for file ${fileServerRelativeUrl}`,
                        });
                        resolve();
                    })
                        .catch(err => {
                        sp_pnp_js_1.Logger.log({
                            data: { err, fileServerRelativeUrl },
                            level: sp_pnp_js_1.LogLevel.Error,
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
                                sp_pnp_js_1.Logger.log({
                                    data: { fileServerRelativeUrl, listView, err },
                                    level: sp_pnp_js_1.LogLevel.Error,
                                    message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                                });
                                reject(err);
                            });
                        })
                            .catch(err => {
                            sp_pnp_js_1.Logger.log({
                                data: { fileServerRelativeUrl, listView, err },
                                level: sp_pnp_js_1.LogLevel.Error,
                                message: `Failed to process page list view for file ${fileServerRelativeUrl}`,
                            });
                            reject(err);
                        });
                    })
                        .catch(err => {
                        sp_pnp_js_1.Logger.log({
                            data: { fileServerRelativeUrl, listView, err },
                            level: sp_pnp_js_1.LogLevel.Error,
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
                sp_pnp_js_1.Logger.log({
                    data: { fileServerRelativeUrl, listView, err },
                    level: sp_pnp_js_1.LogLevel.Error,
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
<<<<<<< HEAD
     * @param {FileAddResuylt} result The file add result
     * @param {Object} properties The properties to set
     */
    processProperties(web, result, properties) {
        return new Promise((resolve, reject) => {
            if (properties && Object.keys(properties).length > 0) {
                result.file.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get()
                    .then(({ ID, ParentList }) => {
                    web.lists.getById(ParentList.Id).items.getById(ID).update(properties)
                        .then(resolve)
                        .catch(reject);
                })
                    .catch(reject);
            }
            else {
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
    replaceXmlTokens(str, ctx) {
        let site = sp_pnp_js_1.Util.combinePaths(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    }
=======
     * @param {File} pnpFile The PnP file
     * @param {Object} properties The properties to set
     */
    processProperties(web, pnpFile, properties) {
        return __awaiter(this, void 0, void 0, function* () {
            if (properties && Object.keys(properties).length > 0) {
                const listItemAllFields = yield pnpFile.listItemAllFields.select("ID", "ParentList/ID").expand("ParentList").get();
                yield web.lists.getById(listItemAllFields.ParentList.Id).items.getById(listItemAllFields.ID).update(properties);
            }
        });
    }
>>>>>>> cd8cc30c728a0a98cea63fe411d9a6d8b23c4308
}
exports.Files = Files;
