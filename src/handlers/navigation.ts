import { HandlerBase } from "./handlerbase";
import { INavigation, INavigationNode } from "../schema";
import { Web, NavigationNodes } from "@pnp/sp";
import { isArray } from "@pnp/common";
import { replaceUrlTokens } from "../util";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Navigation Object Handler
 */
export class Navigation extends HandlerBase {
    /**
     * Creates a new instance of the Navigation class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig) {
        super("Navigation", config);
    }

    /**
     * Provisioning navigation
     *
     * @param {Navigation} navigation The navigation to provision
     */
    public async ProvisionObjects(web: Web, navigation: INavigation): Promise<void> {
        super.scope_started();
        let promises = [];
        if (isArray(navigation.QuickLaunch)) {
            promises.push(this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch));
        }
        if (isArray(navigation.TopNavigationBar)) {
            promises.push(this.processNavTree(web.navigation.topNavigationBar, navigation.TopNavigationBar));
        }
        try {
            await Promise.all(promises);
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    private async processNavTree(target: NavigationNodes, nodes: INavigationNode[]): Promise<void> {
        try {
            const existingNodes = await target.get();
            await this.deleteExistingNodes(target);
            await nodes.reduce((chain: any, node) => chain.then(() => this.processNode(target, node, existingNodes)), Promise.resolve());
        } catch (err) {
            throw err;
        }
    }

    private async processNode(target: NavigationNodes, node: INavigationNode, existingNodes: any[]): Promise<void> {
        let existingNode = existingNodes.filter(n => n.Title === node.Title);
        if (existingNode.length === 1 && node.IgnoreExisting !== true) {
            node.Url = existingNode[0].Url;
        }
        try {
            const result = await target.add(node.Title, replaceUrlTokens(node.Url, this.config));
            if (isArray(node.Children)) {
                await this.processNavTree(result.node.children, node.Children);
            }
        } catch (err) {
            throw err;
        }
    }

    private async deleteExistingNodes(target: NavigationNodes) {
        try {
            const existingNodes = await target.get();
            await existingNodes.reduce((chain: Promise<void>, n: any) => chain.then(_ => this.deleteNode(target, n.Id)), Promise.resolve());
        } catch (err) {
            throw err;
        }
    }

    private async deleteNode(target: NavigationNodes, id: number): Promise<void> {
        try {
            await target.getById(id).delete();
        } catch (err) {
            throw err;
        }
    }
}
