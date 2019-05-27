import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import * as AuthenticationContext from "adal-angular";


export interface ISecuredTeamsPageState extends ITeamsBaseComponentState {
    showLoginButton: boolean;
    errorMessage: string;
    host: string;
    status: string;
    user: AuthenticationContext.UserInfo | null;
}

export interface ISecuredTeamsPageProps extends ITeamsBaseComponentProps {
    endpoints?: { [resource: string]: string };
}

export abstract class SecuredTeamsPage<P extends ISecuredTeamsPageProps, S extends ISecuredTeamsPageState> extends TeamsBaseComponent<P, S> {


    protected authContext: AuthenticationContext;

    public constructor(props: P, state: S) {
        super(props, state);
        // this.login = this.login.bind(this);
        // this.tokenRecieved = this.tokenRecieved.bind(this);
        this.signIn = this.signIn.bind(this);
        // this.onSignedIn = this.onSignedIn.bind(this);
    }
    // public abstract tokenRecieved(token, user?): void;
    public abstract onSignedIn(): void;
    public abstract render();

    public getConfig(tid: string): Promise<AuthenticationContext.Options> {
        return new Promise((resolve, reject) => {
            resolve({
                tenant: `${process.env.TENANT_NAME}.onmicrosoft.com`,
                clientId: `${process.env.CLIENT_APP_ID}`,
                endpoints: {
                    orgUri: `https://${process.env.TENANT_NAME}.crm4.dynamics.com`
                },
                redirectUri: window.location.origin + "/silentEnd.html",
                cacheLocation: "localStorage",
                navigateToLoginRequestUrl: false,
            });
        });

    }
    // public componentDidMount() {
    //     // Start the sign-in process
    //     if (this.inTeams()) {
    //         microsoftTeams.initialize();
    //         microsoftTeams.getContext(context => {
    //             this.signIn();
    //         });
    //     } else {
    //         this.signIn();
    //     }
    // }

    public getToken(endpoint: string): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const token = this.authContext.getCachedToken(endpoint);

            if (token) {
                resolve(token);
                // don't read the data in the adal renew frame
                // if (!this.isInAdalRenewFrame()) {
                //     this.authContext.getUser((err, user) => {
                //         console.warn("Why are we here?");
                //     });
                // }
            } else {
                const ac: any = this.authContext;
                this.authContext.getUser((err, user) => {
                    // if (this.inTeams()) {
                    //     ac._renewToken(endpoint, (errorDesc, token2, error, tokenType) => {
                    //         if (error) {
                    //             reject(error);
                    //         } else {
                    //             const t = this.authContext.getCachedToken(endpoint);
                    //             resolve(t);
                    //         }
                    //     });
                    // } else {
                    //     ac.acquireToken(endpoint, (errorDesc: any, token2: any, error: string, tokenType: any) => {
                    //         if (error) {
                    //             if (error === "login required") {
                    //                 console.error("login required");
                    //                 reject("login required");
                    //             } else {
                    //                 alert("Should we redirect or throw an error");
                    //                 ac.acquireTokenRedirect(endpoint, null, null);
                    //             }
                    //         } else {
                    //             const t = this.authContext.getCachedToken(endpoint);
                    //             resolve(t);
                    //         }
                    //     });
                    // }
                    ac._renewToken(endpoint, (errorDesc, token2, error, tokenType) => {
                        if (error) {
                            if (error === "interaction_required") {
                                this.login();
                            }
                            reject(error);
                        } else {
                            const t = this.authContext.getCachedToken(endpoint);
                            resolve(t);
                        }
                    });
                });
            }
        });

    }

    public login() {
        if (this.inTeams()) {
            this.setState({
                status: "Requiring user input for logging in..."
            });
            microsoftTeams.authentication.authenticate({
                url: window.location.origin + "/silentStart.html",
                width: 600,
                height: 535,
                successCallback: (result) => {
                    this.setState({
                        showLoginButton: false
                    });
                    this.signIn();
                },
                failureCallback: (reason) => {
                    this.setState({
                        errorMessage: "Login failed: " + reason
                    });
                    if (reason === "CancelledByUser" || reason === "FailedToOpenWindow") {
                        this.setState({
                            status: "Login was blocked by popup blocker or canceled by user.",
                            showLoginButton: true
                        });
                    }
                }
            });
        } else {
            this.authContext.login();
        }
    }



    public isInAdalRenewFrame() {
        return window.frameElement &&
            (window.frameElement.id.indexOf("adalRenewFrame") !== -1);
    }

    // #region Login functions
    public signIn() {
        // Sign in
        this.setState({
            status: "signing in"
        });
        // Handle any callbacks from ADAL
        const isCallback = this.authContext.isCallback(window.location.hash);
        if (isCallback) {
            this.authContext.handleWindowCallback();
            // return;
        }

        if (!this.isInAdalRenewFrame()) {
            this.authContext.getUser((err, user) => {
                if (err) {
                    this.setState({
                        errorMessage: err,
                        status: "sign in failed",
                        user: null
                    });
                    this.login();
                } else {
                    this.setState({
                        status: "signed in",
                        user
                    });
                    this.onSignedIn();
                }
            });
        }

    }

    protected abstract redirectUri(): string;


}
