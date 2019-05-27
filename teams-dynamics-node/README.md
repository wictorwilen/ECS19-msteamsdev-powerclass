# Field Service demo applications

Demo application by Wictor Wilen

## Setup help

### Create a `.env` file

You need to create a configuration for all your settings and add your settings

``` .env
# The domain name of where you host your application
HOSTNAME=<URL to location where you host your site, withouth https://>

# App Id and App Password ofr the Bot Framework bot
MICROSOFT_APP_ID=<App Id for your bot>
MICROSOFT_APP_PASSWORD=<Client secret for your bot>

CLIENT_APP_ID=<App Id of your client app>
CLIENT_APP_PASSWORD=<Client secret for your client app>
TENANT_NAME=<Name of your tenant>

DYNAMICS_USER=<User account with access to dynamics>
DYNAMICS_PASSWORD=<User account password for the above>

# Port for local debugging
PORT=3007

# ID of the Outlook Connector
CONNECTOR_ID=<ID of your connector>

# Application Insights instrumentation key
APPINSIGHTS_INSTRUMENTATIONKEY=<instrumentation key for app insights>

# NGROK configuration for development
# NGROK authentication token (leave empty for anonymous)
NGROK_AUTH=
# NGROK sub domain. ex "myapp" or  (leave empty for random)
NGROK_SUBDOMAIN=
# NGROK region. (us, eu, au, ap - default is us)
NGROK_REGION=us

# Debug settings
# Default is "msteams".
# Add "express:*" for more Express logging
DEBUG=msteams
```

### Create a Bot

Register a Bot in the Azure portal and enable it for Microsoft Teams. 

Add the following permissions to the bot

- Dynamics CRM - user_impersonation
- Microsoft Graph - User.Read, openid, profile

### Create a client app

Create a new Azure AD App

Add the following permissions to the bot

- Dynamics CRM - user_impersonation
- Microsoft Graph - User.Read

Add the following Redirect URIs

- https://**your host name**/silentEnd.html
- https://**your host name**/api/auth/getAToken



## Getting started with Microsoft Teams Apps development

Head on over to [official documentation](https://msdn.microsoft.com/en-us/microsoft-teams/tabs) to learn how to build Microsoft Teams Tabs.
