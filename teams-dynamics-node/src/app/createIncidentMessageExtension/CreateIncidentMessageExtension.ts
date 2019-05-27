import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { ITaskModuleResult, IMessagingExtensionActionRequest } from "botbuilder-teams-messagingextensions";
import * as AuthenticationContext from "adal-node";

import JsonDB = require("node-json-db");
import CasesMessageExtension from "../casesMessageExtension/CasesMessageExtension";
import Incidents from "../incidents";
import { NodeRequestFactory } from "../NodeRequestFactory";
import { Incident } from "../DynamicsDefinitions";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/createIncidentMessageExtension/config.html")
@PreventIframe("/createIncidentMessageExtension/action.html")
export default class CreateIncidentMessageExtension implements IMessagingExtensionMiddlewareProcessor {



    public async onFetchTask(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult | ITaskModuleResult> {

        const tokens = new JsonDB("tokens", true, false);

        let token: any;
        try {
            token = tokens.getData(`/tokens/${value.messagePayload.from.user.id}`);
        } catch (error) {
            token = undefined;
        }
        if (!token) {
            return Promise.resolve<MessagingExtensionResult>({
                type: "auth", // use "config" or "auth" here
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: `https://${process.env.HOSTNAME}/api/auth/auth?notifyUrl=createIncidentMessageExtension/config.html`,
                            title: "Setup"
                        }
                    ]
                }
            });
        }
        const settings = new JsonDB("settings", true, false);
        let setting;
        // take care of the setting returned from the dialog, with the value stored in state
        if (context.activity.value.state) {
            setting = JSON.parse(context.activity.value.state);
            settings.push(`/${context.activity.channelData.tenant.id}/${context.activity.channelData.team.id}/${context.activity.channelData.channel.id}`, setting, false);
            log(`New setting: ${setting}`);

        } else {
            try {
                setting = settings.getData(`/${context.activity.channelData.tenant.id}/${context.activity.channelData.team.id}/${context.activity.channelData.channel.id}`);
            } catch (err) {
                setting = undefined;
            }
        }
        if (!setting) {
            return Promise.resolve<MessagingExtensionResult>({
                type: "config", // use "config" or "auth" here
                suggestedActions: {
                    actions: [
                        {
                            type: "openUrl",
                            value: `https://${process.env.HOSTNAME}//casesMessageExtension/config.html`,
                            title: "Configuration is required"
                        }
                    ]
                }
            });
        }

        return new Promise<ITaskModuleResult>((resolve, reject) => {
            const authenticationContext = new AuthenticationContext.AuthenticationContext(`https://login.windows.net/${process.env.TENANT_NAME}.onmicrosoft.com`);
            let message: string;
            authenticationContext.acquireTokenWithRefreshToken(
                token.refreshToken,
                process.env.CLIENT_APP_ID as string,
                process.env.CLIENT_APP_PASSWORD as string,
                `https://${process.env.TENANT_NAME}.crm4.dynamics.com`,
                (refreshErr, refreshResponse) => {
                    if (refreshErr) {
                        message += "refreshError: " + refreshErr.message + "\n";
                    }
                    message += "refreshResponse: " + JSON.stringify(refreshResponse);
                    resolve({
                        type: "continue",
                        value: {
                            title: "Input form",
                            card: CardFactory.adaptiveCard({
                                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                                type: "AdaptiveCard",
                                version: "1.0",
                                body: [
                                    {
                                        type: "TextBlock",
                                        text: "New incident"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: setting.selectedAccount
                                    },
                                    {
                                        type: "Input.Text",
                                        id: "title",
                                        placeholder: "Enter a title",
                                        required: true
                                    },
                                    {
                                        type: "Input.Text",
                                        id: "description",
                                        value: value.messagePayload.body.content
                                    },
                                ],
                                actions: [
                                    {
                                        type: "Action.Submit",
                                        title: "OK",
                                        data: { accountId: setting.selectedId }
                                    }
                                ]
                            })
                        }
                    });
                });
        });
    }


    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: IMessagingExtensionActionRequest): Promise<MessagingExtensionResult> {

        return CasesMessageExtension.verifySignedIn(context, (token) => {
            return new Promise<MessagingExtensionResult>((resolve, reject) => {
                const incidents = new Incidents<NodeRequestFactory<Incident>>(token, new NodeRequestFactory<Incident>());
                incidents.add({
                    "description": value.data.description,
                    "customerid_account@odata.bind": `/accounts(${value.data.accountId})`,
                    "title": value.data.title
                }).then(newIncident => {
                    resolve({
                        type: "result",
                        attachmentLayout: "list",
                        attachments: [Incidents.createAdaptiveCard(newIncident)]
                    } as MessagingExtensionResult);
                }).catch(err => {
                    reject(err);
                })

            });
        });
    }


}
