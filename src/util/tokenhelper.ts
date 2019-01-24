
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Token Helper
 */
export class TokenHelper {
    private tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g;

    /**
     * Creates a new instance of the TokenHelper class
     */
    constructor(
        public context: ProvisioningContext,
        public config: IProvisioningConfig
    ) { }

    public replaceTokens(str: string) {
        let m;
        while ((m = this.tokenRegex.exec(str)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [tokenKey, tokenValue] = match.replace(/[\{\}]/g, "").split(":");
                switch (tokenKey) {
                    case "listid": {
                        if (this.context.lists[tokenValue]) {
                            str = str.replace(match, this.context.lists[tokenValue]);
                        }
                    }
                        break;
                    case "siteid": {
                        if (this.context.web.Id) {
                            str = str.replace(match, this.context.web.Id);
                        }
                    }
                        break;
                    case "parameter": {
                        if (this.config.parameters) {
                            const paramValue = this.config.parameters[tokenValue];
                            if (paramValue) {
                                str = str.replace(match, paramValue);
                            }
                        }
                    }
                        break;
                }
            });
        }
        return str;
    }
}
