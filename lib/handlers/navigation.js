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
const handlerbase_1 = require("./handlerbase");
const common_1 = require("@pnp/common");
const util_1 = require("../util");
/**
 * Describes the Navigation Object Handler
 */
class Navigation extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Navigation class
     */
    constructor() {
        super("Navigation");
    }
    /**
     * Provisioning navigation
     *
     * @param {Navigation} navigation The navigation to provision
     */
    ProvisionObjects(web, navigation) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.scope_started.call(this);
            let promises = [];
            if (common_1.isArray(navigation.QuickLaunch)) {
                promises.push(this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
            }
            if (common_1.isArray(navigation.TopNavigationBar)) {
                promises.push(this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
            }
            try {
                yield Promise.all(promises);
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
    processNavTree(target, nodes) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const existingNodes = yield target.get();
                yield this.deleteExistingNodes(target);
                yield nodes.reduce((chain, node) => chain.then(_ => this.processNode(target, node, existingNodes)), Promise.resolve());
            }
            catch (err) {
                throw err;
            }
        });
    }
    processNode(target, node, existingNodes) {
        return __awaiter(this, void 0, void 0, function* () {
            let existingNode = existingNodes.filter(n => n.Title === node.Title);
            if (existingNode.length === 1 && node.IgnoreExisting !== true) {
                node.Url = existingNode[0].Url;
            }
            try {
                const result = yield target.add(node.Title, util_1.replaceUrlTokens(node.Url));
                if (common_1.isArray(node.Children)) {
                    yield this.processNavTree(result.node.children, node.Children);
                }
            }
            catch (err) {
                throw err;
            }
        });
    }
    deleteExistingNodes(target) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const existingNodes = yield target.get();
                yield existingNodes.reduce((chain, n) => chain.then(_ => this.deleteNode(target, n.Id)), Promise.resolve());
            }
            catch (err) {
                throw err;
            }
        });
    }
    deleteNode(target, id) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                yield target.getById(id).delete();
            }
            catch (err) {
                throw err;
            }
        });
    }
}
exports.Navigation = Navigation;
