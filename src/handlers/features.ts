import { HandlerBase } from "./handlerbase";
import { IFeature } from "../schema";
import { Web } from "@pnp/sp";
import { IProvisioningConfig} from "../provisioningconfig";

/**
 * Describes the Features Object Handler
 */
export class Features extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig) {
        super("Features", config);
    }

    /**
     * Provisioning features
     *
     * @param {Web} web The web
     * @param {Array<IFeature>} features The features to provision
     */
    public async ProvisionObjects(web: Web, features: IFeature[]): Promise<void> {
        super.scope_started();
        try {
            await features.reduce((chain, feature) => {
                if (feature.deactivate) {
                    return chain.then(() => web.features.remove(feature.id, feature.force));
                } else {
                    return chain.then(() => web.features.add(feature.id, feature.force));
                }
            }, Promise.resolve<any>({}));
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }
}
