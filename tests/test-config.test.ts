declare var global: any;
import * as chai from "chai";
import "mocha";
import { sp, Web } from "@pnp/sp";
import { extend, combine, getGUID } from "@pnp/common";
import { SPFetchClient } from "@pnp/nodejs";
import * as chaiAsPromised from "chai-as-promised";
chai.use(chaiAsPromised);

export let testSettings = extend(global.settings.testing, { webUrl: "" });

before(function (done: MochaDone) {

    // this may take some time, don't timeout early
    this.timeout(90000);

    // establish the connection to sharepoint
    if (testSettings.enableWebTests) {

        sp.setup({
            sp: {
                fetchClientFactory: () => {
                    return new SPFetchClient(testSettings.siteUrl, testSettings.clientId, testSettings.clientSecret);
                },
            },
        });

        // comment this out to keep older subsites
        // cleanUpAllSubsites();

        // create the web in which we will test
        let d = new Date();
        let g = getGUID();

        sp.web.webs.add(`PnP-JS-Provisioning Testing ${d.toDateString()}`, g).then(() => {

            let url = combine(testSettings.siteUrl, g);

            // set the testing web url so our tests have access if needed
            testSettings.webUrl = url;

            // re-setup the node client to use the new web
            sp.setup({
                // headers: {
                //     "Accept": "application/json;odata=verbose",
                // },
                sp: {
                    fetchClientFactory: () => {
                        return new SPFetchClient(url, testSettings.clientId, testSettings.clientSecret);
                    },
                },
            });

            done();
        });
    } else {
        done();
    }
});

after(() => {

    // could remove the sub web here?
    // clean up other stuff?
    // write some logging?
});

// this can be used to clean up lots of test sub webs :)
export function cleanUpAllSubsites() {
    sp.site.rootWeb.webs.select("Title").get().then((w) => {
        w.forEach((element: any) => {
            let web = new Web(element["odata.id"], "");
            web.webs.select("Title").get().then((sw: any[]) => {
                return Promise.all(sw.map((value) => {
                    let web2 = new Web(value["odata.id"], "");
                    return web2.delete();
                }));
            }).then(() => { web.delete(); });
        });
    }).catch(e => console.log("Error: " + JSON.stringify(e)));
}
