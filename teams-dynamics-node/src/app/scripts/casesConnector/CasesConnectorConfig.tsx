import * as React from "react";
import {
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Dropdown,
    IDropdownItemProps,
    Surface,
    TeamsThemeContext,
    PrimaryButton
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { getContext } from "msteams-ui-styles-core";
import { Account, AccountResult } from "../../DynamicsDefinitions";
import { ISecuredTeamsPageState, ISecuredTeamsPageProps, SecuredTeamsPage } from "../SecuredTeamsPage";
import * as AuthenticationContext from "adal-angular";
import Accounts from "../../accounts";
import { ClientRequestFactory } from "../../ClientRequestFactory";
export interface ICasesConnectorConfigState extends ISecuredTeamsPageState {
    submit: boolean;
    webhookUrl: string;
    appType: string;
    groupName: string;
    accounts: Account[];
    selectedAccount: any;
    selectedId: any;
}

export interface ICasesConnectorConfigProps extends ISecuredTeamsPageProps {
}


/**
 * Implementation of the casesConnector Connector connect page
 */
export class CasesConnectorConfig extends SecuredTeamsPage<ICasesConnectorConfigProps, ICasesConnectorConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();

            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.setState({
                    selectedId: context.entityId,
                    selectedAccount: (context as any).configName
                });
                this.setValidityState(this.state.selectedId !== undefined);
                this.getConfig(context.tid as string).then(config => {
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
                // INFO: Should really be of type microsoftTeams.settings.Settings, but configName does not exist in the Teams JS SDK
                const settings: any = {
                    entityId: this.state.selectedId,
                    contentUrl: `https://${process.env.HOSTNAME}/casesConnector/config.html`,
                    configName: this.state.selectedAccount
                };
                window.status = `https://${process.env.HOSTNAME}/casesConnector/config.html`;
                microsoftTeams.settings.setSettings(settings);

                microsoftTeams.settings.getSettings((s: any) => {
                    this.setState({
                        webhookUrl: s.webhookUrl,
                        user: s.userObjectId,
                        appType: s.appType,
                    });

                    fetch("/api/connector/connect", {
                        method: "POST",
                        headers: [
                            ["Content-Type", "application/json"]
                        ],
                        body: JSON.stringify({
                            webhookUrl: this.state.webhookUrl,
                            user: this.state.user,
                            appType: this.state.appType,
                            groupName: this.state.groupName,
                            accountid: this.state.selectedId,
                            state: "myAppsState"
                        })
                    }).then(x => {
                        if (x.status === 200 || x.status === 302) {
                            saveEvent.notifySuccess();
                        } else {
                            saveEvent.notifyFailure(x.statusText);
                        }
                    }).catch(e => {
                        saveEvent.notifyFailure(e);
                    });
                });
            });
        } else {
            // Not in Microsoft Teams
            alert("Operation not supported outside of Microsoft Teams");
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
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4), height: "200px" },
            input: {},
        };

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Configure your Connector</div>
                        </PanelHeader>
                        <PanelBody>
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
            </TeamsThemeContext.Provider >
        );
    }
}
