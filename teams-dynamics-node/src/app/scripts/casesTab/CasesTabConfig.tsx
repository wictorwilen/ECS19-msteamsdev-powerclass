import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Input,
    Surface,
    getContext,
    TeamsThemeContext,
    Dropdown
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { silentLoginTab } from "../silentLoginTab";
import { SecuredTeamsPage, ISecuredTeamsPageState, ISecuredTeamsPageProps } from "../SecuredTeamsPage";
import Accounts from "../../accounts";
import { ClientRequestFactory } from "../../ClientRequestFactory";
import { Account, AccountResult } from "../../DynamicsDefinitions";
import * as AuthenticationContext from "adal-angular";

export interface ICasesTabConfigState extends ISecuredTeamsPageState {
    value: string;
    accounts: Account[];
    selectedAccount: any;
    selectedId: any;

}

export interface ICasesTabConfigProps extends ISecuredTeamsPageProps {

}

/**
 * Implementation of Cases configuration page
 */
export class CasesTabConfig extends SecuredTeamsPage<ICasesTabConfigProps, ICasesTabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.setState({
                    value: context.entityId
                });
                this.getConfig(context.tid as string).then( config => {
                    const upn = context.upn;
                    if (upn) {
                        config.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(upn);
                    } else {
                        config.extraQueryParameter = "scope=openid+profile";
                    }
                    this.authContext = new AuthenticationContext(config);
                    this.signIn();
                });

            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                // Calculate host dynamically to enable local debugging
                const host = "https://" + window.location.host;
                microsoftTeams.settings.setSettings({
                    contentUrl: host + "/casesTab/?data=" + this.state.selectedId,
                    suggestedDisplayName: "Cases for "  + this.state.selectedAccount,
                    removeUrl: host + "/casesTab/remove.html",
                    entityId: this.state.selectedId
                });
                saveEvent.notifySuccess();
            });
        } else {
        }
    }
    public redirectUri(): string {
        return "/casesTab/";
    }
    public onSignedIn(): void {

        this.getToken(`https://${process.env.TENANT_NAME}.crm4.dynamics.com`).then(token => {
            if (!token) {
                this.setState({
                    showLoginButton: true,
                    status: "No token",
                });
                return;
            }

            this.setState({
                status: "Fetching data",
            });
            this.retrieveAccounts(token);
        });
    }

    public retrieveAccounts(token: string) {
        if (!token) {
            this.setState({
                showLoginButton: true,
                status: "No token"
            });
            return;
        }

        this.setState({
            status: "Fetching data"
        });
        const accounts = new Accounts<ClientRequestFactory<Account>>(token, new ClientRequestFactory<Account>());
        accounts.getAll().then(accs => {
            this.updateAccounts(accs);
        });


    }
    public updateAccounts(accounts: Account[]) {
        const account = accounts.filter(a => {
            return a.accountid === this.state.selectedId;
        });
        // tslint:disable-next-line: no-console
        console.log(account);
        this.setState({
            accounts,
            status: "",
            showLoginButton: false,
            errorMessage: "",
            selectedAccount: account.length === 0 ? undefined : account[0].name,
            selectedId: account.length === 0 ? undefined : account[0].accountid,

        });
    }


    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Configure your tab</div>
                        </PanelHeader>
                        <PanelBody style={{ height: "300px" }}>
                            <div style={styles.section}>
                                {this.state.status}
                            </div>
                            {this.state.errorMessage &&
                                <div style={styles.section}>
                                    <b>Error: </b>{this.state.errorMessage}
                                </div>
                            }
                            {this.state.showLoginButton &&
                                <div style={styles.section}>
                                    <PrimaryButton onClick={() => this.login()}>Login</PrimaryButton>
                                </div>
                            }

                            {this.state.accounts &&
                                <div style={styles.section}>
                                    <Dropdown
                                        autoFocus
                                        style={{ width: "100%" }}
                                        label="Select account"
                                        mainButtonText={this.state.selectedAccount}
                                        items={this.state.accounts.map(c => {
                                            return {
                                                text: c.name,
                                                onClick: () => {
                                                    this.setState({
                                                        selectedAccount: c.name,
                                                        selectedId: c.accountid
                                                    }, () => {
                                                        this.setValidityState(true);
                                                    });
                                                }
                                            };
                                        })}
                                    />
                                </div>
                            }

                        </PanelBody>
                        <PanelFooter>
                        </PanelFooter>
                    </Panel>
                </Surface>
                `  </TeamsThemeContext.Provider>
        );
    }
}
