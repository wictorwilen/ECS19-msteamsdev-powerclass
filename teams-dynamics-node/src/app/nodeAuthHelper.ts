import * as adal from "adal-node";
import debug = require("debug");
const log = debug("msteams");

// tslint:disable-next-line: class-name
export default class nodeAuthHelper {
    private authContext: adal.AuthenticationContext;
    private readonly cache: any;

    constructor() {
        this.cache = new adal.MemoryCache();
        this.authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/${process.env.TENANT_NAME}.onmicrosoft.com`, true, this.cache);
    }
    public async getTokenWithSecret(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            this.authContext.acquireTokenWithClientCredentials(
                `https://${process.env.TENANT_NAME}.crm4.dynamics.com`,
                process.env.CLIENT_APP_ID as string,
                process.env.CLIENT_APP_PASSWORD as string,
                (error: Error, response: adal.TokenResponse) => {
                    if (!error) {
                        resolve(response.accessToken);
                    } else {
                        reject(error.message);
                    }
                });
        });
    }

    public async getTokenWithUsernamePassword(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            this.authContext.acquireTokenWithUsernamePassword(
                `https://${process.env.TENANT_NAME}.crm4.dynamics.com`,
                process.env.DYNAMICS_USER as string,
                process.env.DYNAMICS_PASSWORD as string,
                process.env.CLIENT_APP_ID as string,
                (error: Error, response: adal.TokenResponse) => {
                    if (!error) {
                        resolve(response.accessToken);
                    } else {
                        log(error);
                        reject(error.message);
                    }
                });
        });
    }
}
