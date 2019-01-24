import { IProvisioningConfig } from "../provisioningconfig";
export declare function replaceUrlTokens(str: string, config: IProvisioningConfig): string;
export declare function makeUrlRelative(absUrl: string): string;
export declare function base64EncodeString(str: string): string;
export declare function isNode(): boolean;
