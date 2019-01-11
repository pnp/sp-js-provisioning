/**
 * Describes the Provisioning Context
 */
var ProvisioningContext = (function () {
    /**
     * Creates a new instance of the ProvisioningContext class
     */
    function ProvisioningContext(lists, tokenRegex) {
        if (lists === void 0) { lists = []; }
        if (tokenRegex === void 0) { tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g; }
        this.lists = lists;
        this.tokenRegex = tokenRegex;
    }
    ProvisioningContext.prototype.replaceTokens = function (str) {
        var _this = this;
        var m;
        while ((m = this.tokenRegex.exec(str)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach(function (match) {
                var _a = match.replace(/[\{\}]/g, "").split(":"), tokenType = _a[0], tokenValue = _a[1];
                switch (tokenType) {
                    case "listid": {
                        var list = _this.lists.filter(function (lst) { return lst.Title === tokenValue; })[0];
                        if (list) {
                            str = str.replace(match, list.Id);
                        }
                    }
                }
            });
        }
        return str;
    };
    return ProvisioningContext;
}());
export { ProvisioningContext };
