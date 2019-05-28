import * as microsoftTeams from "@microsoft/teams-js";
/**
 * Implementation of the teams tab1 AdminConsent page
 */
export class AdminConsent {
  /**
   * Constructor for Tab that initializes the Microsoft Teams script and themes management
   */
  constructor() {
    microsoftTeams.initialize();
  }

  public requestConsent(tenantId: string) {
    const redirectUri = "https://" + window.location.host + "/adminconsent.html";
    const clientId = "204e6baa-1955-4261-be06-a1e173fd93d6";
    const state = "officedev-trainingconent"; // any unique value

    const consentEndpoint = "https://login.microsoftonline.com/common/adminconsent?" +
      "client_id=" + clientId +
      "&state=" + state +
      "&redirect_uri=" + redirectUri;

    window.location.replace(consentEndpoint);
  }

  public processResponse(response: boolean, error: string) {
    if (response) {
      microsoftTeams.authentication.notifySuccess();
    } else {
      microsoftTeams.authentication.notifyFailure(error);
    }
  }
}
