# tms-ninja-dr
Tool for enabling and disabling Direct Routing users.

![TMS-NINJA LOGO](https://i.postimg.cc/fLw4zyTz/Frame-5.png)

# Description
The application will connect using OAuth to Graph and Teams using the provided admin and client credentials, and allow the admin to enable users for Direct Routing, configure voice routing, and disable users. When the application is closed, it will also disconnect from Graph and Teams.

``The application does not support MFA. Security considerations in regards to admin and client credentials are left to the user's own diligence.``

# Requirements
- PowerShell
- Admin credentials
- [AAD App Registration](https://docs.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory)
- [Install-Module MgGraph](https://docs.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0)
- [Install-Module MicrosoftTeams](https://docs.microsoft.com/en-us/MicrosoftTeams/teams-powershell-install)

# Functionalities
## Sign in with admin and client credentials
![TMS-NINJA-DR Sign in](https://i.postimg.cc/28CPnV4n/signin.png)

## Enable user
![Enable user](https://i.postimg.cc/B6VXHdzK/enableuser.png)

## Configure voice routing
![Configure voice routing](https://i.postimg.cc/25H499Jz/configurevoicerouting.png)

## Disable user
![Disable user](https://i.postimg.cc/Kjh40nZG/disableuser.png)

# License
- MIT

# Author
Mihai Filip (voidthevillain)