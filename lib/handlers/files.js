var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as xmljs from "xml-js";
import { HandlerBase } from "./handlerbase";
import { combine, isArray } from "@pnp/common";
import { Logger } from "@pnp/logging";
import { replaceUrlTokens } from "../util";
/**
 * Describes the Features Object Handler
 */
var Files = (function (_super) {
    __extends(Files, _super);
    /**
     * Creates a new instance of the Files class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    function Files(config) {
        var _this = _super.call(this, "Files", config) || this;
        /**
         * Fetches web part contents
         *
         * @param {Array<IWebPart>} webParts Web parts
         * @param {Function} callbackFunc Callback function that takes index of the the webpart and the retrieved XML
         */
        _this.fetchWebPartContents = function (webParts, callbackFunc) {
            return new Promise(function (resolve, reject) {
                var fileFetchPromises = webParts.map(function (wp, index) {
                    return (function () {
                        return new Promise(function (_res, _rej) { return __awaiter(_this, void 0, void 0, function () {
                            var fileSrc, response, xml, obj, existingProperties, updatedProperties_1;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        if (!wp.Contents.FileSrc) return [3 /*break*/, 3];
                                        fileSrc = replaceUrlTokens(this.context.replaceTokens(wp.Contents.FileSrc), this.config);
                                        Logger.log({ data: null, level: 1 /* Info */, message: "Retrieving contents from file '" + fileSrc + "'." });
                                        return [4 /*yield*/, fetch(fileSrc, { credentials: "include", method: "GET" })];
                                    case 1:
                                        response = _a.sent();
                                        return [4 /*yield*/, response.text()];
                                    case 2:
                                        xml = _a.sent();
                                        if (isArray(wp.PropertyOverrides)) {
                                            obj = xmljs.xml2js(xml);
                                            if (obj.elements[0].name === "webParts") {
                                                existingProperties = obj.elements[0].elements[0].elements[1].elements[0].elements;
                                                updatedProperties_1 = [];
                                                existingProperties.forEach(function (prop) {
                                                    var hasOverride = wp.PropertyOverrides.filter(function (po) { return po.name === prop.attributes.name; }).length > 0;
                                                    if (!hasOverride) {
                                                        updatedProperties_1.push(prop);
                                                    }
                                                });
                                                wp.PropertyOverrides.forEach(function (_a) {
                                                    var name = _a.name, type = _a.type, value = _a.value;
                                                    updatedProperties_1.push({ attributes: { name: name, type: type }, elements: [{ text: value, type: "text" }], name: "property", type: "element" });
                                                });
                                                obj.elements[0].elements[0].elements[1].elements[0].elements = updatedProperties_1;
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
                                        return [3 /*break*/, 4];
                                    case 3:
                                        _res();
                                        _a.label = 4;
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); });
                    })();
                });
                Promise.all(fileFetchPromises)
                    .then(resolve)
                    .catch(reject);
            });
        };
        return _this;
    }
    /**
     * Provisioning Files
     *
     * @param {Web} web The web
     * @param {IFile[]} files The files  to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    Files.prototype.ProvisionObjects = function (web, files, context) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var ServerRelativeUrl, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.context = context;
                        _super.prototype.scope_started.call(this);
                        if (typeof window === "undefined") {
                            throw "Files Handler not supported in Node.";
                        }
                        if (this.config.spfxContext) {
                            throw "Files Handler not supported in SPFx.";
                        }
                        return [4 /*yield*/, web.get()];
                    case 1:
                        ServerRelativeUrl = (_a.sent()).ServerRelativeUrl;
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, files.reduce(function (chain, file) { return chain.then(function () { return _this.processFile(web, file, ServerRelativeUrl); }); }, Promise.resolve())];
                    case 3:
                        _a.sent();
                        _super.prototype.scope_ended.call(this);
                        return [3 /*break*/, 5];
                    case 4:
                        err_1 = _a.sent();
                        _super.prototype.scope_ended.call(this);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get blob for a file
     *
     * @param {IFile} file The file
     */
    Files.prototype.getFileBlob = function (file) {
        return __awaiter(this, void 0, void 0, function () {
            var fileSrcWithoutTokens, response, fileContents, blob;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        fileSrcWithoutTokens = replaceUrlTokens(this.context.replaceTokens(file.Src), this.config);
                        return [4 /*yield*/, fetch(fileSrcWithoutTokens, { credentials: "include", method: "GET" })];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.text()];
                    case 2:
                        fileContents = _a.sent();
                        blob = new Blob([fileContents], { type: "text/plain" });
                        return [2 /*return*/, blob];
                }
            });
        });
    };
    /**
     * Procceses a file
     *
     * @param {Web} web The web
     * @param {IFile} file The fileAddError
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     */
    Files.prototype.processFile = function (web, file, webServerRelativeUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var blob, folderServerRelativeUrl, pnpFolder, fileServerRelativeUrl, fileAddResult, pnpFile, fileAddError_1, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        Logger.log({ level: 1 /* Info */, message: "Processing file " + file.Folder + "/" + file.Url });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 10, , 11]);
                        return [4 /*yield*/, this.getFileBlob(file)];
                    case 2:
                        blob = _a.sent();
                        folderServerRelativeUrl = combine("/", webServerRelativeUrl, file.Folder);
                        pnpFolder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
                        fileServerRelativeUrl = combine("/", folderServerRelativeUrl, file.Url);
                        fileAddResult = void 0;
                        pnpFile = void 0;
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, pnpFolder.files.add(file.Url, blob, file.Overwrite)];
                    case 4:
                        fileAddResult = _a.sent();
                        pnpFile = fileAddResult.file;
                        fileServerRelativeUrl = fileAddResult.data.ServerRelativeUrl;
                        return [3 /*break*/, 6];
                    case 5:
                        fileAddError_1 = _a.sent();
                        pnpFile = web.getFileByServerRelativePath(fileServerRelativeUrl);
                        return [3 /*break*/, 6];
                    case 6: return [4 /*yield*/, this.processProperties(web, pnpFile, file)];
                    case 7:
                        _a.sent();
                        return [4 /*yield*/, this.processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl)];
                    case 8:
                        _a.sent();
                        return [4 /*yield*/, this.processPageListViews(web, file.WebParts, fileServerRelativeUrl)];
                    case 9:
                        _a.sent();
                        return [3 /*break*/, 11];
                    case 10:
                        err_2 = _a.sent();
                        throw err_2;
                    case 11: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Remove exisiting webparts if specified
     *
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     * @param {boolean} shouldRemove Should web parts be removed
     */
    Files.prototype.removeExistingWebParts = function (webServerRelativeUrl, fileServerRelativeUrl, shouldRemove) {
        return new Promise(function (resolve, reject) {
            if (shouldRemove) {
                Logger.log({ level: 1 /* Info */, message: "Deleting existing webpart from file " + fileServerRelativeUrl });
                var clientContext_1 = new SP.ClientContext(webServerRelativeUrl);
                var spFile = clientContext_1.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl);
                var webPartManager = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                var webParts_1 = webPartManager.get_webParts();
                clientContext_1.load(webParts_1);
                clientContext_1.executeQueryAsync(function () {
                    webParts_1.get_data().forEach(function (wp) { return wp.deleteWebPart(); });
                    clientContext_1.executeQueryAsync(resolve, reject);
                }, reject);
            }
            else {
                Logger.log({ level: 1 /* Info */, message: "Web parts should not be removed from file " + fileServerRelativeUrl + "." });
                resolve();
            }
        });
    };
    /**
     * Processes web parts
     *
     * @param {IFile} file The file
     * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    Files.prototype.processWebParts = function (file, webServerRelativeUrl, fileServerRelativeUrl) {
        var _this = this;
        return new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            var clientContext_2, spFile, webPartManager_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        Logger.log({ level: 1 /* Info */, message: "Processing webparts for file " + file.Folder + "/" + file.Url });
                        return [4 /*yield*/, this.removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, file.RemoveExistingWebParts)];
                    case 1:
                        _a.sent();
                        if (!(file.WebParts && file.WebParts.length > 0)) return [3 /*break*/, 3];
                        clientContext_2 = new SP.ClientContext(webServerRelativeUrl), spFile = clientContext_2.get_web().getFileByServerRelativeUrl(fileServerRelativeUrl), webPartManager_1 = spFile.getLimitedWebPartManager(SP.WebParts.PersonalizationScope.shared);
                        return [4 /*yield*/, this.fetchWebPartContents(file.WebParts, function (index, xml) { file.WebParts[index].Contents.Xml = xml; })];
                    case 2:
                        _a.sent();
                        file.WebParts.forEach(function (wp) {
                            var webPartXml = _this.context.replaceTokens(_this.replaceWebPartXmlTokens(wp.Contents.Xml, clientContext_2));
                            var webPartDef = webPartManager_1.importWebPart(webPartXml);
                            var webPartInstance = webPartDef.get_webPart();
                            Logger.log({ data: { webPartXml: webPartXml }, level: 1 /* Info */, message: "Processing webpart " + wp.Title + " for file " + file.Folder + "/" + file.Url });
                            webPartManager_1.addWebPart(webPartInstance, wp.Zone, wp.Order);
                            clientContext_2.load(webPartInstance);
                        });
                        clientContext_2.executeQueryAsync(resolve, function (sender, args) {
                            Logger.log({
                                data: { error: args.get_message() },
                                level: 3 /* Error */,
                                message: "Failed to process webparts for file " + file.Folder + "/" + file.Url,
                            });
                            reject({ sender: sender, args: args });
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        resolve();
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        }); });
    };
    /**
     * Processes page list views
     *
     * @param {Web} web The web
     * @param {Array<IWebPart>} webParts Web parts
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    Files.prototype.processPageListViews = function (web, webParts, fileServerRelativeUrl) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (webParts) {
                Logger.log({
                    data: { webParts: webParts, fileServerRelativeUrl: fileServerRelativeUrl },
                    level: 1 /* Info */,
                    message: "Processing page list views for file " + fileServerRelativeUrl,
                });
                var listViewWebParts = webParts.filter(function (wp) { return wp.ListView; });
                if (listViewWebParts.length > 0) {
                    listViewWebParts
                        .reduce(function (chain, wp) { return chain.then(function () { return _this.processPageListView(web, wp.ListView, fileServerRelativeUrl); }); }, Promise.resolve())
                        .then(function () {
                        Logger.log({
                            data: {},
                            level: 1 /* Info */,
                            message: "Successfully processed page list views for file " + fileServerRelativeUrl,
                        });
                        resolve();
                    })
                        .catch(function (err) {
                        Logger.log({
                            data: { err: err, fileServerRelativeUrl: fileServerRelativeUrl },
                            level: 3 /* Error */,
                            message: "Failed to process page list views for file " + fileServerRelativeUrl,
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
    };
    /**
     * Processes page list view
     *
     * @param {Web} web The web
     * @param {any} listView List view
     * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
     */
    Files.prototype.processPageListView = function (web, listView, fileServerRelativeUrl) {
        return new Promise(function (resolve, reject) {
            var views = web.lists.getByTitle(listView.List).views;
            views.get()
                .then(function (listViews) {
                var wpView = listViews.filter(function (v) { return v.ServerRelativeUrl === fileServerRelativeUrl; });
                if (wpView.length === 1) {
                    var view_1 = views.getById(wpView[0].Id);
                    var settings = listView.View.AdditionalSettings || {};
                    view_1.update(settings)
                        .then(function () {
                        view_1.fields.removeAll()
                            .then(function () {
                            listView.View.ViewFields.reduce(function (chain, viewField) { return chain.then(function () { return view_1.fields.add(viewField); }); }, Promise.resolve())
                                .then(resolve)
                                .catch(function (err) {
                                Logger.log({
                                    data: { fileServerRelativeUrl: fileServerRelativeUrl, listView: listView, err: err },
                                    level: 3 /* Error */,
                                    message: "Failed to process page list view for file " + fileServerRelativeUrl,
                                });
                                reject(err);
                            });
                        })
                            .catch(function (err) {
                            Logger.log({
                                data: { fileServerRelativeUrl: fileServerRelativeUrl, listView: listView, err: err },
                                level: 3 /* Error */,
                                message: "Failed to process page list view for file " + fileServerRelativeUrl,
                            });
                            reject(err);
                        });
                    })
                        .catch(function (err) {
                        Logger.log({
                            data: { fileServerRelativeUrl: fileServerRelativeUrl, listView: listView, err: err },
                            level: 3 /* Error */,
                            message: "Failed to process page list view for file " + fileServerRelativeUrl,
                        });
                        reject(err);
                    });
                }
                else {
                    resolve();
                }
            })
                .catch(function (err) {
                Logger.log({
                    data: { fileServerRelativeUrl: fileServerRelativeUrl, listView: listView, err: err },
                    level: 3 /* Error */,
                    message: "Failed to process page list view for file " + fileServerRelativeUrl,
                });
                reject(err);
            });
        });
    };
    /**
     * Process list item properties for the file
     *
     * @param {Web} web The web
     * @param {File} pnpFile The PnP file
     * @param {Object} properties The properties to set
     */
    Files.prototype.processProperties = function (web, pnpFile, file) {
        return __awaiter(this, void 0, void 0, function () {
            var hasProperties, listItemAllFields;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        hasProperties = file.Properties && Object.keys(file.Properties).length > 0;
                        if (!hasProperties) return [3 /*break*/, 3];
                        Logger.log({ level: 1 /* Info */, message: "Processing properties for " + file.Folder + "/" + file.Url });
                        return [4 /*yield*/, pnpFile.listItemAllFields.select("ID", "ParentList/ID", "ParentList/Title").expand("ParentList").get()];
                    case 1:
                        listItemAllFields = _a.sent();
                        return [4 /*yield*/, web.lists.getById(listItemAllFields.ParentList.Id).items.getById(listItemAllFields.ID).update(file.Properties)];
                    case 2:
                        _a.sent();
                        Logger.log({ level: 1 /* Info */, message: "Successfully processed properties for " + file.Folder + "/" + file.Url });
                        _a.label = 3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
    * Replaces tokens in a string, e.g. {site}
    *
    * @param {string} str The string
    * @param {SP.ClientContext} ctx Client context
    */
    Files.prototype.replaceWebPartXmlTokens = function (str, ctx) {
        var site = combine(document.location.protocol, "//", document.location.host, ctx.get_url());
        return str.replace(/{site}/g, site);
    };
    return Files;
}(HandlerBase));
export { Files };
