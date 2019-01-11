import { HandlerBase } from "./handlerbase";
import { INavigation } from "../schema";
import { Web } from "@pnp/sp";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Navigation Object Handler
 */
export declare class Navigation extends HandlerBase {
    /**
     * Creates a new instance of the Navigation class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning navigation
     *
     * @param {Navigation} navigation The navigation to provision
     */
    ProvisionObjects(web: Web, navigation: INavigation): Promise<void>;
    private processNavTree(target, nodes);
    private processNode(target, node, existingNodes);
    private deleteExistingNodes(target);
    private deleteNode(target, id);
}
