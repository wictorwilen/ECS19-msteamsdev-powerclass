import * as React from "react";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { ISecuredTeamsPageState, ISecuredTeamsPageProps, SecuredTeamsPage } from "../SecuredTeamsPage";
import { Incident, Account } from "../../DynamicsDefinitions";
import Accounts from "../../accounts";
import { ClientRequestFactory } from "../../ClientRequestFactory";
import * as AuthenticationContext from "adal-angular";
import { IncidentCard } from "./IncidentCard";

/**
 * State for the casesTabTab React component
 */
export interface ICasesTabState extends ISecuredTeamsPageState {
    entityId?: string;
    incidents: Incident[];
}

/**
 * Properties for the casesTabTab React component
 */
export interface ICasesTabProps extends ISecuredTeamsPageProps {

}

/**
 * Implementation of the Cases content page
 */
export class CasesTab extends SecuredTeamsPage<ICasesTabProps, ICasesTabState> {

    public constructor(props: ICasesTabProps, state: ICasesTabState) {
        super(props, state);
        this.retrieveIncidents = this.retrieveIncidents.bind(this);
        this.updateIncidents = this.updateIncidents.bind(this);
    }

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
                // tslint:disable-next-line: no-console
                console.info(context);
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
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
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
            this.retrieveIncidents(token);
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
            footer: { ...sizes.xsmall },
            incidentsArea: {
                display: "flex",
                alignItems: "left",
                justifyContent: "left",
                flexDirection: "row",
                flexWrap: "wrap",
                flexFlow: "row wrap",
                alignContent: "flex-end"
            } as React.CSSProperties,

        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Cases</div>
                        </PanelHeader>
                        <PanelBody>
                            <div style={styles.section}>
                                {this.state.incidents &&
                                    <div style={styles.incidentsArea}>
                                        {this.state.incidents.map(incident => {
                                            return <IncidentCard
                                                title={incident.title}
                                                contacted={incident.customercontacted}
                                                created={incident["createdon@OData.Community.Display.V1.FormattedValue"]}
                                                createdBy={incident["_createdby_value@OData.Community.Display.V1.FormattedValue"]}
                                                origin={incident["caseorigincode@OData.Community.Display.V1.FormattedValue"]}
                                                owner={incident["_ownerid_value@OData.Community.Display.V1.FormattedValue"]}
                                                status={incident["statuscode@OData.Community.Display.V1.FormattedValue"]}
                                                priority={incident["prioritycode@OData.Community.Display.V1.FormattedValue"]}
                                                type={incident["casetypecode@OData.Community.Display.V1.FormattedValue:"]}
                                                caseno={incident.ticketnumber}
                                                contact={incident["_primarycontactid_value@OData.Community.Display.V1.FormattedValue"]}
                                                id={incident.incidentid}
                                            />;
                                        })}
                                    </div>
                                }
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

    public retrieveIncidents(token) {
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
        const incidents = new Accounts<ClientRequestFactory<Account>>(token, new ClientRequestFactory<Account>());
        incidents.getSubEntity<Incident>(`accountid eq ${this.state.entityId}`, "incident_customer_accounts").then(incs => {
            this.updateIncidents(incs);
        });
    }
    public updateIncidents(incidents: Incident[]) {
        this.setState({
            incidents,
            status: "",
            showLoginButton: false,
            errorMessage: "",
        });
    }
}
