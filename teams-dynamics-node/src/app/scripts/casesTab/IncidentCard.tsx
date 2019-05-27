import React = require("react");
import { Panel, connectTeamsComponent, ITeamsThemeContextProps, Checkbox, PanelHeader, PrimaryButton, SecondaryButton, getContext, ThemeStyle, IconButton } from "msteams-ui-components-react";

interface ICaseProps {
    title: string;
    status: string;
    createdBy: string;
    owner: string;
    origin: string;
    created: string;
    contacted: boolean;
    priority: string;
    type: string;
    caseno: string;
    contact: string;
    id: string;
}

type Props = ICaseProps & ITeamsThemeContextProps;

interface IComponentState { }

// tslint:disable-next-line: class-name
class incidentCard extends React.Component<Props, IComponentState> {
    public render() {
        const context = this.props.context;
        const { rem, font, colors } = context;
        const { sizes, weights } = font;
        const styles = {
            title: { ...sizes.title, ...weights.semibold },
            facts: { ...sizes.caption, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall },
            incident: {
                flex: "0 1 320px",
                height: "80px",
                ...sizes.base,
                width: "600px",
                margin: rem(1.4),
                border: rem(0.2) + " solid",
                borderRadius: rem(0.3),
                padding: rem(0.4),
                font: "inherit",
                color: colors.light.gray02,
                borderColor: colors.light.gray06,
            } as React.CSSProperties,
            prioIncident: {
                flex: "0 1 320px",
                height: "80px",
                ...sizes.base,
                width: "600px",
                margin: rem(1.4),
                border: rem(0.2) + " solid",
                borderRadius: rem(0.3),
                padding: rem(0.4),
                font: "inherit",
                color: colors.light.gray02,
                borderColor: colors.light.brand00Dark,
                backgroundColor: colors.light.brand14
            } as React.CSSProperties,
            input: {

            }

        };
        return (
            <Panel>
                <div style={this.props.priority === "High" ? styles.prioIncident : styles.incident}>
                    <div style={styles.title}>{this.props.title}</div>
                    Case number: <a href={`https://${process.env.TENANT_NAME}.crm4.dynamics.com/main.aspx?appid=222612b7-414e-e911-a96d-000d3a45ddd6&pagetype=entityrecord&etn=incident&id=` + this.props.id}>{this.props.caseno}</a>
                    <div style={styles.facts}>
                        Status: {this.props.status}<br />
                        Created: {this.props.created}<br />
                        Created by: {this.props.createdBy}<br />
                        Owner: {this.props.owner}<br />
                        Case origin: {this.props.origin}<br />
                    </div>
                    Status: {this.props.status}<br />
                    Type: {this.props.type}<br />
                    Priority: {this.props.priority}<br />
                    <Checkbox checked={this.props.contacted} label="Customer contacted" style={styles.input} />
                    <PrimaryButton title={"Call " + this.props.contact} onClick={() => alert("Calling...")} style={styles.input}>Call {this.props.contact}</PrimaryButton>
                    <SecondaryButton style={styles.input}>Assign to...</SecondaryButton>
                </div>
            </Panel>
        );
    }
}

export const IncidentCard = connectTeamsComponent(incidentCard);
