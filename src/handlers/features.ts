import { HandlerBase } from "./handlerbase";
import { IFeature } from "../schema";
import { Web } from "sp-pnp-js";

/**
 * Describes the Features Object Handler
 */
export class Features extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Features");
    }

    /**
     * Provisioning features
     *
     * @param {Web} web The web
     * @param {Array<IFeature>} features The features to provision
     */
    public ProvisionObjects(web: Web, features: IFeature[]): Promise<void> {
        super.scope_started();
        return new Promise<any>((resolve, reject) => {
            features.reduce((chain, feature) => {
                if (feature.deactivate) {
                    return chain.then(() => web.features.remove(feature.id, feature.force));
                } else {
                    return chain.then(() => web.features.add(feature.id, feature.force));
                }
            }, Promise.resolve<any>({})).then(() => {
                super.scope_ended();
                resolve();
            }).catch((e: Error) => {
                super.scope_ended();
                reject(e);
            });
        });
    }
}
