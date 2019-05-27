import * as microsoftTeams from "@microsoft/teams-js";
import * as AuthenticationContext from "adal-angular";

// tslint:disable-next-line: class-name
export class silentLoginTab {
    public static getConfig(tid: string): Promise<AuthenticationContext.Options> {
        return new Promise((resolve, reject) => {
            resolve({
                tenant: `${process.env.TENANT_NAME}.onmicrosoft.com`,
                clientId: `${process.env.CLIENT_APP_ID}`,
                endpoints: {
                    orgUri: `https://${process.env.TENANT_NAME}.crm4.dynamics.com/`,
                    settings: "api://21e4bc8f-e7e6-4f13-a590-2db1d70d66e9"
                },
                redirectUri: window.location.origin + "/silentEnd.html",
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
            });
        });

    }

    public static Start() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            silentLoginTab.getConfig(context.tid as string).then(config => {
                if (context.upn) {
                    config.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(context.upn);
                } else {
                    config.extraQueryParameter = "scope=openid+profile";
                }
                const authContext = new AuthenticationContext(config);
                authContext.login();
            });
        });
    }

    public static End() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            silentLoginTab.getConfig(context.tid as string).then(config => {
                const authContext = new AuthenticationContext(config);
                if (authContext.isCallback(window.location.hash)) {
                    authContext.handleWindowCallback(window.location.hash);
                    if (authContext.getCachedUser()) {
                        authContext.acquireToken(config.endpoints!.orgUri, (errorDesc, token, error) => {
                            microsoftTeams.authentication.notifySuccess(token as string);
                        });

                    } else {
                        microsoftTeams.authentication.notifyFailure(authContext.getLoginError());
                    }
                }
            });
        });

    }
}
