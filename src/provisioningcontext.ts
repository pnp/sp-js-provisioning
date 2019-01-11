/**
 * Describes the Provisioning Context
 */
export class ProvisioningContext {
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    constructor(
        public lists: { [key: string]: string } = {},
        public tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g,
    ) { }

    public replaceTokens(str: string) {
        let m;
        while ((m = this.tokenRegex.exec(str)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach((match) => {
                let [tokenType, tokenValue] = match.replace(/[\{\}]/g, "").split(":");
                switch (tokenType) {
                    case "listid": {
                        if (this.lists[tokenValue]) {
                            str = str.replace(match, this.lists[tokenValue]);
                        }
                    }
                }
            });
        }
        return str;
    }
}
