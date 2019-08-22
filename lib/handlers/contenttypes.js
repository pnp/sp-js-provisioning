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
import initSpfxJsom, { ExecuteJsomQuery } from "spfx-jsom";
import { HandlerBase } from "./handlerbase";
/**
 * Describes the Content Types Object Handler
 */
var ContentTypes = (function (_super) {
    __extends(ContentTypes, _super);
    /**
     * Creates a new instance of the ObjectSiteFields class
     */
    function ContentTypes(config) {
        return _super.call(this, "ContentTypes", config) || this;
    }
    /**
     * Provisioning Content Types
     *
     * @param {Web} web The web
     * @param {IContentType[]} contentTypes The content types
     * @param {ProvisioningContext} context Provisioning context
     */
    ContentTypes.prototype.ProvisionObjects = function (web, contentTypes, context) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var _a, _b, err_1;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _a = this;
                        return [4 /*yield*/, initSpfxJsom(context.web.ServerRelativeUrl)];
                    case 1:
                        _a.jsomContext = (_c.sent()).jsomContext;
                        this.context = context;
                        _super.prototype.scope_started.call(this);
                        _c.label = 2;
                    case 2:
                        _c.trys.push([2, 5, , 6]);
                        _b = this.context;
                        return [4 /*yield*/, web.contentTypes.select('Id', 'Name', 'FieldLinks').expand('FieldLinks').get()];
                    case 3:
                        _b.contentTypes = (_c.sent()).reduce(function (obj, contentType) {
                            obj[contentType.Name] = {
                                ID: contentType.Id.StringValue,
                                Name: contentType.Name,
                                FieldRefs: contentType.FieldLinks.map(function (fl) { return ({
                                    ID: fl.Id,
                                    Name: fl.Name,
                                    Required: fl.Required,
                                    Hidden: fl.Hidden,
                                }); }),
                            };
                            return obj;
                        }, {});
                        return [4 /*yield*/, contentTypes
                                .sort(function (a, b) {
                                if (a.ID < b.ID) {
                                    return -1;
                                }
                                if (a.ID > b.ID) {
                                    return 1;
                                }
                                return 0;
                            })
                                .reduce(function (chain, contentType) { return chain.then(function () { return _this.processContentType(web, contentType); }); }, Promise.resolve())];
                    case 4:
                        _c.sent();
                        return [3 /*break*/, 6];
                    case 5:
                        err_1 = _c.sent();
                        _super.prototype.scope_ended.call(this);
                        throw err_1;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Provision a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    ContentTypes.prototype.processContentType = function (web, contentType) {
        return __awaiter(this, void 0, void 0, function () {
            var contentTypeId, contentTypeAddResult, spContentType, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        contentTypeId = this.context.contentTypes[contentType.Name].ID;
                        if (!!contentTypeId) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.addContentType(web, contentType)];
                    case 1:
                        contentTypeAddResult = _a.sent();
                        contentTypeId = contentTypeAddResult.data.Id;
                        _a.label = 2;
                    case 2:
                        _super.prototype.log_info.call(this, "processContentType", "Processing content type [" + contentType.Name + "] (" + contentTypeId + ")");
                        spContentType = this.jsomContext.web.get_contentTypes().getById(contentTypeId);
                        if (contentType.Description) {
                            spContentType.set_description(contentType.Description);
                        }
                        if (contentType.Group) {
                            spContentType.set_group(contentType.Group);
                        }
                        spContentType.update(true);
                        return [4 /*yield*/, ExecuteJsomQuery(this.jsomContext)];
                    case 3:
                        _a.sent();
                        if (!contentType.FieldRefs) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.processContentTypeFieldRefs(contentType, spContentType)];
                    case 4:
                        _a.sent();
                        _a.label = 5;
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        err_2 = _a.sent();
                        throw err_2;
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Add a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    ContentTypes.prototype.addContentType = function (web, contentType) {
        return __awaiter(this, void 0, void 0, function () {
            var err_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        _super.prototype.log_info.call(this, "addContentType", "Adding content type [" + contentType.Name + "] (" + contentType.ID + ")");
                        return [4 /*yield*/, web.contentTypes.add(contentType.ID, contentType.Name, contentType.Description, contentType.Group)];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        err_3 = _a.sent();
                        throw err_3;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
 * Adding content type field refs
 *
 * @param {IContentType} contentType Content type
 * @param {SP.ContentType} spContentType Content type
 */
    ContentTypes.prototype.processContentTypeFieldRefs = function (contentType, spContentType) {
        return __awaiter(this, void 0, void 0, function () {
            var _loop_1, this_1, i, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        _loop_1 = function (i) {
                            var fieldRef = contentType.FieldRefs[i];
                            var existingFieldLink = this_1.context.contentTypes[contentType.Name].FieldRefs.filter(function (fr) { return fr.Name === fieldRef.Name; })[0];
                            var fieldLink = void 0;
                            if (existingFieldLink) {
                                fieldLink = spContentType.get_fieldLinks().getById(new SP.Guid(existingFieldLink.ID));
                            }
                            else {
                                _super.prototype.log_info.call(this_1, "processContentTypeFieldRefs", "Adding field ref " + fieldRef.Name + " to content type [" + contentType.Name + "] (" + contentType.ID + ")");
                                var siteField = this_1.jsomContext.web.get_fields().getByInternalNameOrTitle(fieldRef.Name);
                                var fieldLinkCreationInformation = new SP.FieldLinkCreationInformation();
                                fieldLinkCreationInformation.set_field(siteField);
                                fieldLink = spContentType.get_fieldLinks().add(fieldLinkCreationInformation);
                            }
                            if (contentType.FieldRefs[i].hasOwnProperty("Required")) {
                                fieldLink.set_required(contentType.FieldRefs[i].Required);
                            }
                            if (contentType.FieldRefs[i].hasOwnProperty("Hidden")) {
                                fieldLink.set_hidden(contentType.FieldRefs[i].Hidden);
                            }
                        };
                        this_1 = this;
                        for (i = 0; i < contentType.FieldRefs.length; i++) {
                            _loop_1(i);
                        }
                        spContentType.update(true);
                        return [4 /*yield*/, ExecuteJsomQuery(this.jsomContext)];
                    case 1:
                        _a.sent();
                        _super.prototype.log_info.call(this, "processContentTypeFieldRefs", "Successfully processed field refs for content type [" + contentType.Name + "] (" + contentType.ID + ")");
                        return [3 /*break*/, 3];
                    case 2:
                        error_1 = _a.sent();
                        _super.prototype.log_info.call(this, "processContentTypeFieldRefs", "Failed to process field refs for content type [" + contentType.Name + "] (" + contentType.ID + ")", { error: error_1.args && error_1.args.get_message() });
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return ContentTypes;
}(HandlerBase));
export { ContentTypes };
