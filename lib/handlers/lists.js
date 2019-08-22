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
import * as xmljs from 'xml-js';
import { HandlerBase } from './handlerbase';
import { TokenHelper } from '../util/tokenhelper';
/**
 * Describes the Lists Object Handler
 */
var Lists = (function (_super) {
    __extends(Lists, _super);
    /**
     * Creates a new instance of the Lists class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    function Lists(config) {
        return _super.call(this, 'Lists', config) || this;
    }
    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    Lists.prototype.ProvisionObjects = function (web, lists, context) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var _a, _b, err_1;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        this.context = context;
                        this.tokenHelper = new TokenHelper(this.context, this.config);
                        _super.prototype.scope_started.call(this);
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 8, , 9]);
                        _a = this.context;
                        return [4 /*yield*/, web.lists.select('Id', 'Title').get()];
                    case 2:
                        _a.lists = (_c.sent()).reduce(function (obj, l) {
                            obj[l.Title] = l.Id;
                            return obj;
                        }, {});
                        return [4 /*yield*/, lists.reduce(function (chain, list) { return chain.then(function () { return _this.processList(web, list); }); }, Promise.resolve())];
                    case 3:
                        _c.sent();
                        return [4 /*yield*/, lists.reduce(function (chain, list) { return chain.then(function () { return _this.processListFields(web, list); }); }, Promise.resolve())];
                    case 4:
                        _c.sent();
                        return [4 /*yield*/, lists.reduce(function (chain, list) { return chain.then(function () { return _this.processListFieldRefs(web, list); }); }, Promise.resolve())];
                    case 5:
                        _c.sent();
                        return [4 /*yield*/, lists.reduce(function (chain, list) { return chain.then(function () { return _this.processListViews(web, list); }); }, Promise.resolve())];
                    case 6:
                        _c.sent();
                        _b = this.context;
                        return [4 /*yield*/, web.lists.select('Id', 'Title').get()];
                    case 7:
                        _b.lists = (_c.sent()).reduce(function (obj, l) {
                            obj[l.Title] = l.Id;
                            return obj;
                        }, {});
                        _super.prototype.scope_ended.call(this);
                        return [3 /*break*/, 9];
                    case 8:
                        err_1 = _c.sent();
                        _super.prototype.scope_ended.call(this);
                        throw err_1;
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    Lists.prototype.processList = function (web, lc) {
        return __awaiter(this, void 0, void 0, function () {
            var list, listEnsure, listAdd;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _super.prototype.log_info.call(this, 'processList', "Processing list " + lc.Title);
                        if (!this.context.lists[lc.Title]) return [3 /*break*/, 2];
                        _super.prototype.log_info.call(this, 'processList', "List " + lc.Title + " already exists. Ensuring...");
                        return [4 /*yield*/, web.lists.ensure(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings)];
                    case 1:
                        listEnsure = _a.sent();
                        list = listEnsure.list;
                        return [3 /*break*/, 4];
                    case 2:
                        _super.prototype.log_info.call(this, 'processList', "List " + lc.Title + " doesn't exist. Creating...");
                        return [4 /*yield*/, web.lists.add(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings)];
                    case 3:
                        listAdd = _a.sent();
                        list = listAdd.list;
                        this.context.lists[listAdd.data.Title] = listAdd.data.Id;
                        _a.label = 4;
                    case 4:
                        if (!lc.ContentTypeBindings) return [3 /*break*/, 6];
                        return [4 /*yield*/, this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes)];
                    case 5:
                        _a.sent();
                        _a.label = 6;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    Lists.prototype.processContentTypeBindings = function (lc, list, contentTypeBindings, removeExisting) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var promises_1, contentTypes;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _super.prototype.log_info.call(this, 'processContentTypeBindings', "Processing content types for list " + lc.Title + ".");
                        return [4 /*yield*/, contentTypeBindings.reduce(function (chain, ct) { return chain.then(function () { return _this.processContentTypeBinding(lc, list, ct.ContentTypeID); }); }, Promise.resolve())];
                    case 1:
                        _a.sent();
                        if (!removeExisting) return [3 /*break*/, 4];
                        promises_1 = [];
                        return [4 /*yield*/, list.contentTypes.get()];
                    case 2:
                        contentTypes = _a.sent();
                        contentTypes.forEach(function (_a) {
                            var ContentTypeId = _a.Id.StringValue;
                            var shouldRemove = (contentTypeBindings.filter(function (ctb) { return ContentTypeId.indexOf(ctb.ContentTypeID) !== -1; }).length === 0)
                                && (ContentTypeId.indexOf('0x0120') === -1);
                            if (shouldRemove) {
                                _super.prototype.log_info.call(_this, 'processContentTypeBindings', "Removing content type " + ContentTypeId + " from list " + lc.Title);
                                promises_1.push(list.contentTypes.getById(ContentTypeId).delete());
                            }
                            else {
                                _super.prototype.log_info.call(_this, 'processContentTypeBindings', "Skipping removal of content type " + ContentTypeId + " from list " + lc.Title);
                            }
                        });
                        return [4 /*yield*/, Promise.all(promises_1)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    Lists.prototype.processContentTypeBinding = function (lc, list, contentTypeID) {
        return __awaiter(this, void 0, void 0, function () {
            var err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        _super.prototype.log_info.call(this, 'processContentTypeBinding', "Adding content Type " + contentTypeID + " to list " + lc.Title + ".");
                        return [4 /*yield*/, list.contentTypes.addAvailableContentType(contentTypeID)];
                    case 1:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processContentTypeBinding', "Content Type " + contentTypeID + " added successfully to list " + lc.Title + ".");
                        return [3 /*break*/, 3];
                    case 2:
                        err_2 = _a.sent();
                        _super.prototype.log_info.call(this, 'processContentTypeBinding', "Failed to add Content Type " + contentTypeID + " to list " + lc.Title + ".");
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    Lists.prototype.processListFields = function (web, list) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!list.Fields) return [3 /*break*/, 2];
                        return [4 /*yield*/, list.Fields.reduce(function (chain, field) { return chain.then(function () { return _this.processField(web, list, field); }); }, Promise.resolve())];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    Lists.prototype.processField = function (web, lc, fieldXml) {
        return __awaiter(this, void 0, void 0, function () {
            var list, fXmlJson, fieldAttr, fieldName, fieldDisplayName, field, err_3, fieldAddResult, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = web.lists.getByTitle(lc.Title);
                        fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
                        fieldAttr = fXmlJson.elements[0].attributes;
                        fieldName = fieldAttr.Name;
                        fieldDisplayName = fieldAttr.DisplayName;
                        _super.prototype.log_info.call(this, 'processField', "Processing field " + fieldName + " (" + fieldDisplayName + ") for list " + lc.Title + ".");
                        fXmlJson.elements[0].attributes.DisplayName = fieldName;
                        fieldXml = xmljs.json2xml(fXmlJson);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, list.fields.getById(fieldAttr.ID)];
                    case 2:
                        field = _a.sent();
                        return [4 /*yield*/, field.delete()];
                    case 3:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processField', "Field " + fieldName + " (" + fieldDisplayName + ") successfully deleted from list " + lc.Title + ".");
                        return [3 /*break*/, 5];
                    case 4:
                        err_3 = _a.sent();
                        _super.prototype.log_info.call(this, 'processField', "Field " + fieldName + " (" + fieldDisplayName + ") does not exist in list " + lc.Title + ".");
                        return [3 /*break*/, 5];
                    case 5:
                        _a.trys.push([5, 8, , 9]);
                        return [4 /*yield*/, list.fields.createFieldAsXml(this.tokenHelper.replaceTokens(fieldXml))];
                    case 6:
                        fieldAddResult = _a.sent();
                        return [4 /*yield*/, fieldAddResult.field.update({ Title: fieldDisplayName })];
                    case 7:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processField', "Field '" + fieldDisplayName + "' added successfully to list " + lc.Title + ".");
                        return [3 /*break*/, 9];
                    case 8:
                        err_4 = _a.sent();
                        _super.prototype.log_info.call(this, 'processField', "Failed to add field '" + fieldDisplayName + "' to list " + lc.Title + ".");
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    Lists.prototype.processListFieldRefs = function (web, list) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!list.FieldRefs) return [3 /*break*/, 2];
                        return [4 /*yield*/, list.FieldRefs.reduce(function (chain, fieldRef) { return chain.then(function () { return _this.processFieldRef(web, list, fieldRef); }); }, Promise.resolve())];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    Lists.prototype.processFieldRef = function (web, lc, fieldRef) {
        return __awaiter(this, void 0, void 0, function () {
            var list, err_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        list = web.lists.getByTitle(lc.Title);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName })];
                    case 2:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processFieldRef', "Field '" + fieldRef.ID + "' updated for list " + lc.Title + ".");
                        return [3 /*break*/, 4];
                    case 3:
                        err_5 = _a.sent();
                        _super.prototype.log_info.call(this, 'processFieldRef', "Failed to update field '" + fieldRef.ID + "' for list " + lc.Title + ".");
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    Lists.prototype.processListViews = function (web, lc) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!lc.Views) return [3 /*break*/, 2];
                        return [4 /*yield*/, lc.Views.reduce(function (chain, view) { return chain.then(function () { return _this.processView(web, lc, view); }); }, Promise.resolve())];
                    case 1:
                        _b.sent();
                        _b.label = 2;
                    case 2:
                        _a = this.context;
                        return [4 /*yield*/, web.lists.getByTitle(lc.Title).views.select('Id', 'Title').get()];
                    case 3:
                        _a.listViews = (_b.sent()).reduce(function (obj, view) {
                            obj[lc.Title + "|" + view.Title] = view.Id;
                            return obj;
                        }, this.context.listViews);
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    Lists.prototype.processView = function (web, lc, lvc) {
        return __awaiter(this, void 0, void 0, function () {
            var existingView, viewExists, err_6, result, err_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _super.prototype.log_info.call(this, 'processView', "Processing view " + lvc.Title + " for list " + lc.Title + ".");
                        existingView = web.lists.getByTitle(lc.Title).views.getByTitle(lvc.Title);
                        viewExists = false;
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, existingView.get()];
                    case 2:
                        _a.sent();
                        viewExists = true;
                        return [3 /*break*/, 4];
                    case 3:
                        err_6 = _a.sent();
                        return [3 /*break*/, 4];
                    case 4:
                        _a.trys.push([4, 11, , 12]);
                        if (!viewExists) return [3 /*break*/, 7];
                        _super.prototype.log_info.call(this, 'processView', "View " + lvc.Title + " for list " + lc.Title + " already exists, updating.");
                        return [4 /*yield*/, existingView.update(lvc.AdditionalSettings)];
                    case 5:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processView', "View " + lvc.Title + " successfully updated for list " + lc.Title + ".");
                        return [4 /*yield*/, this.processViewFields(existingView, lvc)];
                    case 6:
                        _a.sent();
                        return [3 /*break*/, 10];
                    case 7:
                        _super.prototype.log_info.call(this, 'processView', "View " + lvc.Title + " for list " + lc.Title + " doesn't exists, creating.");
                        return [4 /*yield*/, web.lists.getByTitle(lc.Title).views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings)];
                    case 8:
                        result = _a.sent();
                        _super.prototype.log_info.call(this, 'processView', "View " + lvc.Title + " added successfully to list " + lc.Title + ".");
                        return [4 /*yield*/, this.processViewFields(result.view, lvc)];
                    case 9:
                        _a.sent();
                        _a.label = 10;
                    case 10: return [3 /*break*/, 12];
                    case 11:
                        err_7 = _a.sent();
                        _super.prototype.log_info.call(this, 'processViewFields', "Failed to process view for view " + lvc.Title + ".");
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {IListView} lvc The view configuration
     */
    Lists.prototype.processViewFields = function (view, lvc) {
        return __awaiter(this, void 0, void 0, function () {
            var err_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        _super.prototype.log_info.call(this, 'processViewFields', "Processing view fields for view " + lvc.Title + ".");
                        return [4 /*yield*/, view.fields.removeAll()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, lvc.ViewFields.reduce(function (chain, viewField) { return chain.then(function () { return view.fields.add(viewField); }); }, Promise.resolve())];
                    case 2:
                        _a.sent();
                        _super.prototype.log_info.call(this, 'processViewFields', "View fields successfully processed for view " + lvc.Title + ".");
                        return [3 /*break*/, 4];
                    case 3:
                        err_8 = _a.sent();
                        _super.prototype.log_info.call(this, 'processViewFields', "Failed to process view fields for view " + lvc.Title + ".");
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return Lists;
}(HandlerBase));
export { Lists };
