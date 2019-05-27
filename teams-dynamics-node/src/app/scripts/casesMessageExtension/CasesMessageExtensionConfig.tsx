import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    Checkbox,
    Dropdown,
    TeamsThemeContext,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { ISecuredTeamsPageState, ISecuredTeamsPageProps, SecuredTeamsPage } from "../SecuredTeamsPage";
import * as AuthenticationContext from "adal-angular";
import Accounts from "../../accounts";
import { ClientRequestFactory } from "../../ClientRequestFactory";
import { Account, AccountResult } from "../../DynamicsDefinitions";

/**
 * State for the CasesMessageExtensionConfig React component
 */
export interface ICasesMessageExtensionConfigState extends ISecuredTeamsPageState {
    value: string;
    accounts: Account[];
    selectedAccount: any;
    selectedId: any;
    token: string;
    tenantId?: string;
    teamId?: string;
    channelId?: string;
    saveEnabled: boolean;
}

/**
 * Properties for the CasesMessageExtensionConfig React component
 */
export interface ICasesMessageExtensionConfigProps extends ISecuredTeamsPageProps {

}

/**
 * Implementation of the Cases configuration page
 */
export class CasesMessageExtensionConfig extends SecuredTeamsPage<ICasesMessageExtensionConfigProps, ICasesMessageExtensionConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

        microsoftTeams.getContext((context: microsoftTeams.Context) => {
            this.setState({
                value: context.entityId,
                tenantId: context.tid,
                teamId: context.teamId,
                channelId: context.channelId,
                saveEnabled: false
            });
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
    }

    public redirectUri(): string {
        return "/casesMessageExtension/config.html";
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
                token
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
            const settings = new ClientRequestFactory<any>();
            settings.get(`https://${process.env.HOSTNAME}/api/settings/${this.state.tenantId}/${this.state.teamId}/${this.state.channelId}`, token).then(setting => {
                this.setState({
                    selectedAccount: setting.selectedAccount,
                    selectedId: setting.selectedId,
                    saveEnabled: true
                });
            });
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

    /**
     * The render() method to create the UI of the tab
     */
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
            section2: { height: "250px", ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Cases configuration</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
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
                                                            selectedId: c.accountid,
                                                            saveEnabled: true
                                                        });
                                                    }
                                                };
                                            })}
                                        />
                                    </div>
                                }
                                <div style={styles.section}>
                                    <PrimaryButton
                                        disabled={!this.state.saveEnabled}
                                        onClick={() => {
                                            microsoftTeams.authentication.notifySuccess(JSON.stringify({
                                                selectedAccount: this.state.selectedAccount,
                                                selectedId: this.state.selectedId
                                            }));
                                        }}>Save</PrimaryButton>
                                </div>
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div style={styles.footer}>
                                (C) Copyright Wictor Wilen
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
}
