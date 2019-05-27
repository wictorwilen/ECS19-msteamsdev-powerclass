import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, TokenResponse } from "botbuilder";
import { MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder-teams";
import { IMessagingExtensionMiddlewareProcessor, IAppBasedLinkQuery } from "botbuilder-teams-messagingextensions";
import nodeAuthHelper from "../nodeAuthHelper";
import { Incident } from "../DynamicsDefinitions";
import { NodeRequestFactory } from "../NodeRequestFactory";
import Incidents from "../incidents";
import * as request from "request";
import JsonDB = require("node-json-db");
import * as AuthenticationContext from "adal-node";
import { rejects } from "assert";


// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/casesMessageExtension/config.html")
export default class CasesMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public static verifySignedIn(context: TurnContext, success: (accessToken: string) => Promise<MessagingExtensionResult>): Promise<MessagingExtensionResult> {
        const tokens = new JsonDB("tokens", true, false);

        let token: any;
        try {
            token = tokens.getData(`/tokens/${context.activity.from.aadObjectId}`);
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
        } else {
            return new Promise<MessagingExtensionResult>(async (resolve, reject) => {
                const authenticationContext = new AuthenticationContext.AuthenticationContext(`https://login.windows.net/${process.env.TENANT_NAME}.onmicrosoft.com`);
                authenticationContext.acquireTokenWithRefreshToken(
                    token.refreshToken,
                    process.env.CLIENT_APP_ID as string,
                    process.env.CLIENT_APP_PASSWORD as string,
                    `https://${process.env.TENANT_NAME}.crm4.dynamics.com`,
                    async (refreshErr, refreshResponse) => {
                        if (refreshErr) {
                            reject(refreshErr.message);
                        }
                        resolve(await success((refreshResponse as any).accessToken));
                    });
            });
        }
    }

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        return CasesMessageExtension.verifySignedIn(context, async (accessToken) => {
            let accountId;

            // get settings
            const settings = new JsonDB("settings", true, false);
            try {
                accountId = settings.getData(`/${context.activity.channelData.tenant.id}/${context.activity.channelData.team.id}/${context.activity.channelData.channel.id}/selectedId`);
            } catch (err) {
                accountId = undefined;
            }

            if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
                // initial run
                try {
                    const incidents = new Incidents<NodeRequestFactory<Incident>>(accessToken, new NodeRequestFactory<Incident>());
                    const all = accountId ? await incidents.getByFilter(`_accountid_value eq '${accountId}'`) : await incidents.getAll();
                    return Promise.resolve({
                        type: "result",
                        attachmentLayout: "list",
                        attachments:
                            all.map(i => {
                                return Incidents.createAdaptiveCard(i);
                            })
                    } as MessagingExtensionResult);
                } catch (err) {
                    log(err);
                    return Promise.reject(err);
                }
            } else {
                // the rest
                const incidents = new Incidents<NodeRequestFactory<Incident>>(accessToken, new NodeRequestFactory<Incident>());
                if (query.parameters && query.parameters[0] && query.parameters[0].value) {
                    try {
                        const all = accountId ?
                            await incidents.getByFilter(`_accountid_value eq '${accountId}' and contains(title,'${query.parameters[0].value}')`) :
                            await incidents.getByFilter(`contains(title,'${query.parameters[0].value}')`);
                        return Promise.resolve({
                            type: "result",
                            attachmentLayout: "list",
                            attachments:
                                all.map(i => {
                                    return Incidents.createAdaptiveCard(i);
                                })
                        } as MessagingExtensionResult);
                    } catch (err) {
                        log(err);
                        return Promise.resolve({
                            type: "result",
                            attachmentLayout: "list",
                            attachments: [
                            ]
                        } as MessagingExtensionResult);
                    }


                } else {
                    return Promise.resolve({
                        type: "result",
                        attachmentLayout: "list",
                        attachments: [
                        ]
                    } as MessagingExtensionResult);
                }

            }
        });
    }

    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "Cases Configuration",
            value: `https://${process.env.HOSTNAME}/casesMessageExtension/config.html`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = JSON.parse(context.activity.value.state);
        log(`New setting: ${setting}`);
        const settings = new JsonDB("settings", true, false);
        settings.push(`/${context.activity.channelData.tenant.id}/${context.activity.channelData.team.id}/${context.activity.channelData.channel.id}`, setting, false);
        return Promise.resolve();
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        return new Promise<void>(async (resolve, reject) => {
            const tokens = new JsonDB("tokens", true, false);

            let token: any;
            try {
                token = tokens.getData(`/tokens/${context.activity.from.aadObjectId}`);
            } catch (error) {
                token = undefined;
            }
            if (!token) {
                return reject("User is not signed in");
            }
            // Handle the Action.Submit action on the adaptive card
            if (value.action === "moreDetails") {
                log(`I got this ${value.id}`);
            }
            if (value.incidentid) {
                const authenticationContext = new AuthenticationContext.AuthenticationContext(`https://login.windows.net/${process.env.TENANT_NAME}.onmicrosoft.com`);
                authenticationContext.acquireTokenWithRefreshToken(
                    token.refreshToken,
                    process.env.CLIENT_APP_ID as string,
                    process.env.CLIENT_APP_PASSWORD as string,
                    `https://${process.env.TENANT_NAME}.crm4.dynamics.com`,
                    async (refreshErr, refreshResponse) => {
                        if (refreshErr) {
                            reject(refreshErr.message);
                        }
                        const incidents = new Incidents<NodeRequestFactory<Incident>>((refreshResponse as any).accessToken, new NodeRequestFactory<Incident>());
                        await incidents.resolve({
                            IncidentId: {
                                "incidentid": value.incidentid,
                                "@odata.type": "Microsoft.Dynamics.CRM.incident"
                            },
                            Status: 5,
                            BillableTime: value.billableTime,
                            Resolution: value.resolution,
                            Remarks: "Resolved through adaptive card"
                        })
                        resolve();
                    });


            } else {
                resolve();
            }
        });
    }

    public async onQueryLink(context: TurnContext, value: IAppBasedLinkQuery): Promise<MessagingExtensionResult> {
        if (this.getQueryVariable(value.url, "pagetype") === "entityrecord" &&
            this.getQueryVariable(value.url, "etn") === "incident" &&
            this.getQueryVariable(value.url, "id")) {

            const id = this.getQueryVariable(value.url, "id");
            return CasesMessageExtension.verifySignedIn(context, async (accessToken) => {
                const incidents = new Incidents<NodeRequestFactory<Incident>>(accessToken, new NodeRequestFactory<Incident>());
                try {
                    const incident = await incidents.getById(id!);
                    const card = CardFactory.adaptiveCard({
                        type: "AdaptiveCard",
                        body: [
                            {
                                type: "TextBlock",
                                size: "Large",
                                text: incident.title
                            },
                            {
                                type: "TextBlock",
                                size: "Medium",
                                text: incident["_customerid_value@OData.Community.Display.V1.FormattedValue"]
                            }
                        ],
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        version: "1.0",
                    });
                    const preview = {
                        contentType: "application/vnd.microsoft.card.thumbnail",
                        content: {
                            title: incident.title,
                            text: incident["_customerid_value@OData.Community.Display.V1.FormattedValue"],
                            images: [
                                {
                                    url: `https://${process.env.HOSTNAME}/assets/icon.png`
                                }
                            ]
                        }
                    };
                    return Promise.resolve<MessagingExtensionResult>({
                        type: "result",
                        attachmentLayout: "list", // required
                        attachments: [{ ...card, preview }]
                    });
                } catch (err) {
                    return Promise.reject(`Unable to find the incident: ${err}`);
                }
            });
        } else {
            return new Promise<MessagingExtensionResult>((resolve, reject) => {
                resolve({});
            });
        }
    }

    protected getQueryVariable = (url: string, variable: string): string | undefined => {
        const arr = url.split("?");
        if (arr.length === 2) {
            const vars = arr[1].split("&");
            for (const varPairs of vars) {
                const pair = varPairs.split("=");
                if (decodeURIComponent(pair[0]) === variable) {
                    return decodeURIComponent(pair[1]);
                }
            }
        }
        return undefined;
    }

}
