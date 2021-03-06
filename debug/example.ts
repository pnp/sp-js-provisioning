// use of relative paths to the modules
import { sp } from "@pnp/sp";
import { getGUID } from "@pnp/common";
import { WebProvisioner } from "../src/webprovisioner";
import { default as template } from "../sample-schemas/all-simple";

export function Example() {

    sp.web.webs.add(`Provisioning Debug ${Date.now().toLocaleString()}`, getGUID()).then(war => {
        let provisioner = new WebProvisioner(war.web);

        provisioner.applyTemplate(template).then(() => {

            console.log("Template Applied, checking work...");

            // review the user custom actions
            war.web.userCustomActions.get().then(uca => {
                console.log(`UCA: ${JSON.stringify(uca)}`);
            });

            // review the features to see if ours was deactivated
            war.web.features.get().then(features => {
                console.log(`Features: ${JSON.stringify(features)}`);
            });

            // review the quick launch
            war.web.navigation.quicklaunch.get().then(ql => {
                console.log(`Quick Launch: ${JSON.stringify(ql)}`);
            });

            // review the top navigation
            war.web.navigation.topNavigationBar.get().then(tb => {
                console.log(`Top Navigation Bar: ${JSON.stringify(tb)}`);
            });

            // review the lists
            war.web.lists.get().then(tb => {
                console.log(`Lists: ${JSON.stringify(tb)}`);
            });
        });
    });
}
