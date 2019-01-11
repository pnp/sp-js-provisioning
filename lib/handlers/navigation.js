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
import { isArray } from "@pnp/common";
import { replaceUrlTokens } from "../util";
/**
 * Describes the Navigation Object Handler
 */
var Navigation = (function (_super) {
    __extends(Navigation, _super);
    /**
     * Creates a new instance of the Navigation class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    function Navigation(config) {
        return _super.call(this, "Navigation", config) || this;
    }
    /**
     * Provisioning navigation
     *
     * @param {Navigation} navigation The navigation to provision
     */
    Navigation.prototype.ProvisionObjects = function (web, navigation) {
        return __awaiter(this, void 0, void 0, function () {
            var promises, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _super.prototype.scope_started.call(this);
                        promises = [];
                        if (isArray(navigation.QuickLaunch)) {
                            promises.push(this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
                        }
                        if (isArray(navigation.TopNavigationBar)) {
                            promises.push(this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, Promise.all(promises)];
                    case 2:
                        _a.sent();
                        _super.prototype.scope_ended.call(this);
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
    Navigation.prototype.processNavTree = function (target, nodes) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var existingNodes_1, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        return [4 /*yield*/, target.get()];
                    case 1:
                        existingNodes_1 = _a.sent();
                        return [4 /*yield*/, this.deleteExistingNodes(target)];
                    case 2:
                        _a.sent();
                        return [4 /*yield*/, nodes.reduce(function (chain, node) { return chain.then(function () { return _this.processNode(target, node, existingNodes_1); }); }, Promise.resolve())];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        err_2 = _a.sent();
                        throw err_2;
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    Navigation.prototype.processNode = function (target, node, existingNodes) {
        return __awaiter(this, void 0, void 0, function () {
            var existingNode, result, err_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        existingNode = existingNodes.filter(function (n) { return n.Title === node.Title; });
                        if (existingNode.length === 1 && node.IgnoreExisting !== true) {
                            node.Url = existingNode[0].Url;
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 5, , 6]);
                        return [4 /*yield*/, target.add(node.Title, replaceUrlTokens(node.Url, this.config))];
                    case 2:
                        result = _a.sent();
                        if (!isArray(node.Children)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.processNavTree(result.node.children, node.Children)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        err_3 = _a.sent();
                        throw err_3;
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    Navigation.prototype.deleteExistingNodes = function (target) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var existingNodes, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, target.get()];
                    case 1:
                        existingNodes = _a.sent();
                        return [4 /*yield*/, existingNodes.reduce(function (chain, n) { return chain.then(function (_) { return _this.deleteNode(target, n.Id); }); }, Promise.resolve())];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        err_4 = _a.sent();
                        throw err_4;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    Navigation.prototype.deleteNode = function (target, id) {
        return __awaiter(this, void 0, void 0, function () {
            var err_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, target.getById(id).delete()];
                    case 1:
                        _a.sent();
                        return [3 /*break*/, 3];
                    case 2:
                        err_5 = _a.sent();
                        throw err_5;
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    return Navigation;
}(HandlerBase));
export { Navigation };
