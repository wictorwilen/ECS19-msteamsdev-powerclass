import DynamicsEntity from "./DynamicsEntity";
import { RequestFactory } from "./RequestFactory";
import { Incident } from "./DynamicsDefinitions";
import { CardFactory } from "botbuilder";

export default class Incidents<S extends RequestFactory<Incident>> extends DynamicsEntity<Incident, S> {

    public static createAdaptiveCard(incident: Incident) {
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: incident.title
                    },
                    {
                        type: "TextBlock",
                        text: incident.ticketnumber
                    },
                    {
                        type: "FactSet",
                        facts: [
                            {
                                title: "Customer:",
                                value: incident["_customerid_value@OData.Community.Display.V1.FormattedValue"]
                            },
                            {
                                title: "Type:",
                                value: incident["casetypecode@OData.Community.Display.V1.FormattedValue"]
                            },
                            {
                                title: "Assigned to:",
                                value: incident["_ownerid_value@OData.Community.Display.V1.FormattedValue"]
                            },
                            {
                                title: "Priority:",
                                value: incident["prioritycode@OData.Community.Display.V1.FormattedValue"]
                            }
                        ]
                    }
                ],
                actions: [
                    {
                        type: "Action.OpenUrl",
                        title: "More details",
                        url: `https://${process.env.TENANT_NAME}.crm4.dynamics.com/main.aspx?appid=222612b7-414e-e911-a96d-000d3a45ddd6&pagetype=entityrecord&etn=incident&id=${incident.incidentid}`
                    },
                    {
                        type: "Action.ShowCard",
                        title: "Resolve incident...",
                        card: {
                            type: "AdaptiveCard",
                            body: [
                                {
                                    type: "Input.Text",
                                    id: "resolution",
                                    isMultiline: true,
                                    placeholder: "Add a resolution"
                                },
                                {
                                    type: "TextBlock",
                                    text: "Number of billable minutes"
                                },
                                {
                                    type: "Input.Number",
                                    id: "billableTime",
                                    min: 0,
                                    value: 30
                                }
                            ],
                            actions: [
                                {
                                    type: "Action.Submit",
                                    title: "Resolve",
                                    data: {
                                        ticketnumber: incident.ticketnumber,
                                        incidentid: incident.incidentid
                                    }
                                }
                            ]
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.0"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: incident.title,
                text: incident.ticketnumber,
                images: [
                    {
                        url: `https://${process.env.HOSTNAME}/assets/icon.png`
                    }
                ]
            }
        };
        return { ...card, preview };
    }
    public entityName = "incidents";
    public entityIdName = "incidentid";

    constructor(token: string, factory: S) {
        super(token, factory);
    }


}
