import * as React from "react";
import {
    PrimaryButton,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    Checkbox,
    TeamsThemeContext,
    getContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State for the CreateIncidentMessageExtensionConfig React component
 */
export interface ICreateIncidentMessageExtensionConfigState extends ITeamsBaseComponentState {
}

/**
 * Properties for the CreateIncidentMessageExtensionConfig React component
 */
export interface ICreateIncidentMessageExtensionConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Create Incident configuration page
 */
export class CreateIncidentMessageExtensionConfig extends TeamsBaseComponent<{ ICreateIncidentMessageExtensionConfigProps }, ICreateIncidentMessageExtensionConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
        if (this.getQueryVariable("Success")) {
            microsoftTeams.authentication.notifySuccess();
        } else {
            const err = this.getQueryVariable("Failed");
            if (err) {
                microsoftTeams.authentication.notifyFailure(err);
            } else {
                microsoftTeams.authentication.notifyFailure("Unknown error");
            }
        }
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
            footer: { ...sizes.xsmall }
        };
        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        <PanelHeader>
                            <div style={styles.header}>Create Incident configuration</div>
                        </PanelHeader>
                        <PanelBody>
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
