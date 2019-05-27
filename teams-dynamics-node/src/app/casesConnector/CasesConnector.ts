import * as request from "request";
import { Request } from "express";
import { ConnectorDeclaration, IConnector, PreventIframe } from "express-msteams-host";
import { CardFactory } from "botbuilder-core";
import JsonDB = require("node-json-db");
import * as debug from "debug";
import nodeAuthHelper from "../nodeAuthHelper";
import Accounts from "../accounts";
import { NodeRequestFactory } from "../NodeRequestFactory";
import { IncidentAndAccountResult, Account, Incident } from "../DynamicsDefinitions";
import Incidents from "../incidents";

const log = debug("msteams");

/**
 * The connector data interface
 */
interface ICasesConnectorData {
    webhookUrl: string;
    user: string;
    appType: string;
    groupName: string;
    accountId: string;
    existing: boolean;
}

/**
 * Implementation of the "CasesConnectorConnector" Office 365 Connector
 */
@ConnectorDeclaration(
    "/api/connector/connect",
    "/api/connector/ping"
)
@PreventIframe("/casesConnector/config.html")
export class CasesConnector implements IConnector {
    private connectors: any;

    public constructor() {
        // Instantiate the node-json-db database (connectors.json)
        this.connectors = new JsonDB("connectors", true, false);
    }

    public Connect(req: Request) {
        if (req.body.state === "myAppsState") {
            this.connectors.push("/connectors[]", {
                appType: req.body.appType,
                accountId: req.body.accountid,
                existing: true,
                groupName: req.body.groupName,
                user: req.body.user,
                webhookUrl: req.body.webhookUrl
            });
        }
    }

    public Ping(req: Request): Array<Promise<void>> {
        // clean up connectors marked to be deleted
        try {
            this.connectors.push("/connectors",
                (this.connectors.getData("/connectors") as ICasesConnectorData[]).filter(((c) => {
                    return c.existing;
                })));
        } catch (error) {
            if (error.name && error.name === "DataError") {
                // there"s no registered connectors
                return [];
            }
            throw error;
        }
        const data = req.body;

        // send pings to all subscribers
        return (this.connectors.getData("/connectors") as ICasesConnectorData[]).map((connector, index) => {
            return new Promise<void>(async (resolve, reject) => {
                const auth = new nodeAuthHelper();
                const token = await auth.getTokenWithUsernamePassword();
                const accounts = new Incidents<NodeRequestFactory<Incident>>(token, new NodeRequestFactory<Incident>());
                log(`Processing connector for ${data.PrimaryEntityId}`);
                let result: IncidentAndAccountResult;
                try {
                    result = await accounts.getGenericById(data.PrimaryEntityId, "customerid_account");
                } catch (err) {
                    log(err);
                    reject(err);
                    return;
                }
                if (connector.accountId === result._customerid_value) {
                    // only if subscribed to this one
                    // log(result);
                    // TODO: implement adaptive cards when supported
                    const card = {
                        title: "New case",
                        text: result.title,
                        sections: [
                            {
                                activityTitle: result["_customerid_value@OData.Community.Display.V1.FormattedValue"],
                                activityText: result.ticketnumber,
                                facts: [
                                    {
                                        name: "Product",
                                        value: result["_productid_value@OData.Community.Display.V1.FormattedValue"]
                                    },
                                    {
                                        name: "Created by",
                                        value: result["_createdby_value@OData.Community.Display.V1.FormattedValue"]
                                    }
                                    ,
                                    {
                                        name: "Priority",
                                        value: result["prioritycode@OData.Community.Display.V1.FormattedValue"]
                                    }
                                ]
                            }
                        ],
                        potentialAction: [{
                            "@context": "http://schema.org",
                            "@type": "ViewAction",
                            "name": "Show details",
                            "target": [`https://${process.env.TENANT_NAME}.crm4.dynamics.com/main.aspx?appid=222612b7-414e-e911-a96d-000d3a45ddd6&pagetype=entityrecord&etn=incident&id=` + result.incidentid]
                        }],
                    };

                    log(`Sending card to ${connector.webhookUrl}`);

                    request({
                        method: "POST",
                        uri: decodeURI(connector.webhookUrl),
                        headers: {
                            "content-type": "application/json",
                        },
                        body: JSON.stringify(card)
                    }, (error: any, response: any, body: any) => {
                        log(`Response from Connector endpoint is: ${response.statusCode}`);
                        if (error) {
                            reject(error);
                        } else {
                            // 410 - the user has removed the connector
                            if (response.statusCode === 410) {
                                this.connectors.push(`/connectors[${index}]/existing`, false);
                            }
                            resolve();
                        }
                    });
                } else {
                    log("Skipping this one...");
                    resolve();
                }

            });
        });
    }
}

