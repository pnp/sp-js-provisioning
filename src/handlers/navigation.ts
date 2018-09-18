import { HandlerBase } from "./handlerbase";
import { INavigation, INavigationNode } from "../schema";
import { Web, NavigationNodes, Util } from "sp-pnp-js";
import { replaceUrlTokens } from "../util";

/**
 * Describes the Navigation Object Handler
 */
export class Navigation extends HandlerBase {
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
    public ProvisionObjects(web: Web, navigation: INavigation): Promise<void> {
        super.scope_started();
        return new Promise<void>((resolve, reject) => {
            let promises = [];
            if (Util.isArray(navigation.QuickLaunch)) {
                promises.push(this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
            }
            if (Util.isArray(navigation.TopNavigationBar)) {
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

    private processNavTree(target: NavigationNodes, nodes: INavigationNode[]): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            target.get().then(existingNodes => {
                this.deleteExistingNodes(target).then(() => {
                    nodes.reduce((chain, node) => chain.then(_ => this.processNode(target, node, existingNodes)), Promise.resolve()).then(resolve, reject);
                }, reject);
            });
        });
    }

    private processNode(target: NavigationNodes, node: INavigationNode, existingNodes: any[]): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let existingNode = existingNodes.filter(n => n.Title === node.Title);
            if (existingNode.length === 1 && node.IgnoreExisting !== true) {
                node.Url = existingNode[0].Url;
            }
            target.add(node.Title, replaceUrlTokens(node.Url)).then(result => {
                if (Util.isArray(node.Children)) {
                    this.processNavTree(result.node.children, node.Children).then(resolve, reject);
                } else {
                    resolve();
                }
            }, reject);
        });
    }

    private deleteExistingNodes(target: NavigationNodes) {
        return new Promise<void>((resolve, reject) => {
            target.get().then(existingNodes => {
                existingNodes.reduce((chain: Promise<void>, n: any) => chain.then(_ => this.deleteNode(target, n.Id)), Promise.resolve()).then(() => {
                    resolve();
                }, reject);
            });
        });
    }

    private deleteNode(target: NavigationNodes, id: number): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            target.getById(id).delete().then(resolve, reject);
        });
    }
}
