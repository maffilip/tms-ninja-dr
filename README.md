# tms-ninja-dr
Tool for enabling and disabling Direct Routing users.

![TMS-NINJA LOGO](https://i.postimg.cc/fLw4zyTz/Frame-5.png)

# Description
The application will connect using OAuth to Graph and Teams using the provided admin and client credentials, and allow the admin to enable users for Direct Routing, configure voice routing, and disable users. When the application is closed, it will also disconnect from Graph and Teams.

> The application does not support MFA. Security considerations in regards to admin and client credentials are left to the user's own diligence.

# Disclaimer
``THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.``

# Requirements
- PowerShell
- Admin credentials
- [AAD App Registration](https://docs.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory)
- [Install-Module MgGraph](https://docs.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0)
- [Install-Module MicrosoftTeams](https://docs.microsoft.com/en-us/MicrosoftTeams/teams-powershell-install)

# Instructions
## Register the app in Azure AD
1. Go to [Azure Active Directory](https://aad.portal.azure.com/)
2. Sign in
3. Navigate to Azure Active Directory > App registrations
4. Select **+ New registration**
5. Name your app, select the scope **Single tenant**, and select **Register**
6. From the **Overview** blade of your newly created app registration, copy the **Application (client) ID**
7. Navigate to the **Certificates & secrets** blade, select the **Client secrets** tab, and **+ New client secret**
8. Provide a description, expiration (recommended a short custom time that you plan to use the app for) and select **Add**
9. Copy the value of the newly created secret to a safe location

> Consider deleting the client secret after you're done using the app.

# Functionalities
## Sign in with admin and client credentials
![TMS-NINJA-DR Sign in](https://i.postimg.cc/28CPnV4n/signin.png)
1. Insert **Username**
2. Insert **Password**
3. Insert **Client ID**
4. Insert **Client Secret**
5. Select **Connect**

> It will take approximately 10-20s to connect to both modules. If it takes more, kill the process from Task Manager and retry.

## Enable user
![Enable user](https://i.postimg.cc/B6VXHdzK/enableuser.png)
1. Insert **UPN**
2. Insert **Phone number**
3. Select **Enable**

> Await 5-10s for the user to be enabled and a confirmation will be shown in the text area.

## Configure voice routing
![Configure voice routing](https://i.postimg.cc/25H499Jz/configurevoicerouting.png)
1. Insert **UPN**
2. Select a **Voice routing policy** from a list of available policies in your tenant
3. Select **Assign**

> Await 5-10s for the user to be assigned the voice routing policy and a confirmation will be shown in the text area.

## Disable user
![Disable user](https://i.postimg.cc/Kjh40nZG/disableuser.png)
1. Insert **UPN**
2. (Optional) Select if you want to **Remove voice routing policy**
3. Select **Disable**

> Await 5-10s for the user to be disabled and a confirmation will be shown in the text area.

# License
[MIT](https://raw.githubusercontent.com/voidthevillain/tms-ninja-dr/main/LICENSE)

# Author
Mihai Filip (voidthevillain)
