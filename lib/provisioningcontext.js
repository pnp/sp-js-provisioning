"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Describes the Provisioning Context
 */
class ProvisioningContext {
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor(lists = [], tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g) {
        this.lists = lists;
        this.tokenRegex = tokenRegex;
    }
    replaceTokens(str) {
        let m;
        while ((m = this.tokenRegex.exec(str)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [tokenType, tokenValue] = match.replace(/[\{\}]/g, "").split(":");
                switch (tokenType) {
                    case "listid": {
                        let [list] = this.lists.filter(lst => lst.Title === tokenValue);
                        if (list) {
                            str = str.replace(match, list.Id);
                        }
                    }
                }
            });
        }
        return str;
    }
}
exports.ProvisioningContext = ProvisioningContext;
