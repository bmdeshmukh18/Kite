# Kite
Custom UI for Kite Zerodha by modifying browser CSS and maniputlating data using JS.

To inject custom CSS and trigger custom JS, I am using a chrome extension "User JS and CSS".
Below is the link for the extension.

https://tenrabbits.github.io/user-js-css-docs/

In the JS Code, I am fetching some data like SL, Target from an Excel file.
This Excel file is stored on Personal One Drive.
Data from excel file is fetched using microsoft APIs using Microsoft Graph API endpoint (https://graph.microsoft.com)
And the connection is made by providing the client ID of an app created on Azure.
The connection is authorized by OAuth2.0 authorization endpoint(v2) ()

App can be registered in a Personal Azure account using the AppRegistration(which is free to use feature. Refer below URL)

https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade

How to Use:
1. Install the extension.
2. Create a rule for the kite website.
3. Provide the Website name in this case "https://kite.zerodha.com/"
4. Add the JS and CSS Code from Repository files.
