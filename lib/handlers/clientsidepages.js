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
import { HandlerBase } from "./handlerbase";
import { ClientSideWebpart } from "@pnp/sp";
/**
 * Describes the Composed Look Object Handler
 */
var ClientSidePages = (function (_super) {
    __extends(ClientSidePages, _super);
    /**
     * Creates a new instance of the ObjectClientSidePages class
     */
    function ClientSidePages(config) {
        return _super.call(this, "ClientSidePages", config) || this;
    }
    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {IClientSidePage[]} clientSidePages The client side pages to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    ClientSidePages.prototype.ProvisionObjects = function (web, clientSidePages, context) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.context = context;
                        _super.prototype.scope_started.call(this);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, clientSidePages.reduce(function (chain, clientSidePage) { return chain.then(function () { return _this.processClientSidePage(web, clientSidePage); }); }, Promise.resolve())];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        err_1 = _a.sent();
                        _super.prototype.scope_ended.call(this);
                        throw err_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Provision a client side page
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    ClientSidePages.prototype.processClientSidePage = function (web, clientSidePage) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var page, pageItem;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _super.prototype.log_info.call(this, "processClientSidePage", "Processing client side page " + clientSidePage.Name);
                        return [4 /*yield*/, web.addClientSidePage(clientSidePage.Name, clientSidePage.Title, clientSidePage.LibraryTitle)];
                    case 1:
                        page = _a.sent();
                        if (!clientSidePage.Sections) return [3 /*break*/, 3];
                        clientSidePage.Sections.forEach(function (s) {
                            var section = page.addSection();
                            s.Columns.forEach(function (col) {
                                var column = section.addColumn(col.Factor);
                                col.Controls.forEach(function (control) {
                                    var controlJsonString = _this.context.replaceTokens(JSON.stringify(control));
                                    control = JSON.parse(controlJsonString);
                                    _super.prototype.log_info.call(_this, "processClientSidePage", "Adding " + control.Title + " to client side page " + clientSidePage.Name, { controlJsonString: controlJsonString });
                                    column.addControl(new ClientSideWebpart(control.Title, control.Description, control.ClientSideComponentProperties, control.ClientSideComponentId));
                                });
                            });
                        });
                        return [4 /*yield*/, page.save()];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3: return [4 /*yield*/, page.publish()];
                    case 4:
                        _a.sent();
                        if (!clientSidePage.CommentsDisabled) return [3 /*break*/, 6];
                        _super.prototype.log_info.call(this, "processClientSidePage", "Disabling comments for client side page " + clientSidePage.Name);
                        return [4 /*yield*/, page.disableComments()];
                    case 5:
                        _a.sent();
                        _a.label = 6;
                    case 6:
                        if (!clientSidePage.PageLayoutType) return [3 /*break*/, 9];
                        _super.prototype.log_info.call(this, "processClientSidePage", "Setting page layout " + clientSidePage.PageLayoutType + " for client side page " + clientSidePage.Name);
                        return [4 /*yield*/, page.getItem()];
                    case 7:
                        pageItem = _a.sent();
                        return [4 /*yield*/, pageItem.update({ PageLayoutType: clientSidePage.PageLayoutType })];
                    case 8:
                        _a.sent();
                        _a.label = 9;
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    return ClientSidePages;
}(HandlerBase));
export { ClientSidePages };
