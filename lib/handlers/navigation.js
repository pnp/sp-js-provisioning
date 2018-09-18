"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const sp_pnp_js_1 = require("sp-pnp-js");
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
        super.scope_started();
        return new Promise((resolve, reject) => {
            let promises = [];
            if (sp_pnp_js_1.Util.isArray(navigation.QuickLaunch)) {
                promises.push(this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
            }
            if (sp_pnp_js_1.Util.isArray(navigation.TopNavigationBar)) {
                promises.push(this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
            }
            Promise.all(promises).then(() => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
    processNavTree(target, nodes) {
        return new Promise((resolve, reject) => {
            target.get().then(existingNodes => {
                this.deleteExistingNodes(target).then(() => {
                    nodes.reduce((chain, node) => chain.then(_ => this.processNode(target, node, existingNodes)), Promise.resolve()).then(resolve, reject);
                }, reject);
            });
        });
    }
    processNode(target, node, existingNodes) {
        return new Promise((resolve, reject) => {
            let existingNode = existingNodes.filter(n => n.Title === node.Title);
            if (existingNode.length === 1 && node.IgnoreExisting !== true) {
                node.Url = existingNode[0].Url;
            }
            target.add(node.Title, util_1.replaceUrlTokens(node.Url)).then(result => {
                if (sp_pnp_js_1.Util.isArray(node.Children)) {
                    this.processNavTree(result.node.children, node.Children).then(resolve, reject);
                }
                else {
                    resolve();
                }
            }, reject);
        });
    }
    deleteExistingNodes(target) {
        return new Promise((resolve, reject) => {
            target.get().then(existingNodes => {
                existingNodes.reduce((chain, n) => chain.then(_ => this.deleteNode(target, n.Id)), Promise.resolve()).then(() => {
                    resolve();
                }, reject);
            });
        });
    }
    deleteNode(target, id) {
        return new Promise((resolve, reject) => {
            target.getById(id).delete().then(resolve, reject);
        });
    }
}
exports.Navigation = Navigation;
