/**
 * Describes the Token Helper
 */
var TokenHelper = (function () {
    /**
     * Creates a new instance of the TokenHelper class
     */
    function TokenHelper(context, config) {
        this.context = context;
        this.config = config;
        this.tokenRegex = /{[a-z]*:[ÆØÅæøåA-za-z ]*}/g;
    }
    TokenHelper.prototype.replaceTokens = function (str) {
        var _this = this;
        var m;
        while ((m = this.tokenRegex.exec(str)) !== null) {
            if (m.index === this.tokenRegex.lastIndex) {
                this.tokenRegex.lastIndex++;
            }
            m.forEach(function (match) {
                var _a = match.replace(/[\{\}]/g, "").split(":"), tokenKey = _a[0], tokenValue = _a[1];
                switch (tokenKey) {
                    case "listid":
                        {
                            if (_this.context.lists[tokenValue]) {
                                str = str.replace(match, _this.context.lists[tokenValue]);
                            }
                        }
                        break;
                    case "webid":
                        {
                            if (_this.context.web.Id) {
                                str = str.replace(match, _this.context.web.Id);
                            }
                        }
                        break;
                    case "siteid":
                        {
                            if (_this.context.web.Id) {
                                str = str.replace(match, _this.context.web.Id);
                            }
                        }
                        break;
                    case "sitecollectionid":
                        {
                            if (_this.context.web.Id) {
                                str = str.replace(match, _this.context.web.Id);
                            }
                        }
                        break;
                    case "parameter":
                        {
                            if (_this.config.parameters) {
                                var paramValue = _this.config.parameters[tokenValue];
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
    };
    return TokenHelper;
}());
export { TokenHelper };
