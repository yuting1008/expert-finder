# Expert Finder message extension sample with SSO

This sample demonstrates a Teams message extension integrated with Microsoft Copilot, enabling users to efficiently search experts with Microsoft Graph API based on their skills, location, and availability. The extension features Single Sign-On (SSO) for enhanced user experience and authentication. 
> Refer to the Traditional Chinese version: [README(zh-tw).md](README(zh-tw).md)

### Table of Contents

- [Prerequisites](#prerequisites)
- [Setup and run the sample locally](#setup-and-run-the-sample-locally)
- [Deploy the app to Azure](#deploy-the-app-to-azure)
- [Use the app in Teams and Copilot for Microsoft 365](#use-the-app-in-teams-and-copilot-for-microsoft-365)
- [Further reading](#further-reading)

## Prerequisites

- [Node.js 18.x](https://nodejs.org/download/release/v18.18.2/)
- [Visual Studio Code](https://code.visualstudio.com/)
- [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
- An Azure account. The required resources will later built in this environment.
- A Microsoft 365 work or school account with **Copilot for Microsoft 365 license** and the [**permissions to upload custom Teams applications**](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading). If the permission is disabled, you can either use [Microsoft 365 development account](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program) or contact your tenant administrator to turn on the upload custom app permission for your organization. Below are the steps to enable permission as a Teams admin:
    - When the custom app upload is disabled, the following error is displayed: \
        <img src="images/custom-app-disabled.png" alt="custom-app-disabled" >
    - Go to the [**Teams admin center**](https://admin.teams.microsoft.com/).
    - Navigate to **Teams apps** > **Permission policies**.
    - Click on **Global (Org-wide app default)**. \
        <img src="images/teams-app-upload-permission-1.png" alt="teams-app-upload-permission" height="400">
    - Enable **Upload custom app**. \
        <img src="images/teams-app-upload-permission-2.png" alt="teams-app-upload-permission" height="400">
    - Go to Teams apps > Manage apps > Actions > Org-wide app settings. \
        <img src="images/manage-app.png" alt="teams-app-upload-permission">
    - Turn on the Upload custom apps for personal use toggle. \
        <img src="images/org-wide-app-settings.png" alt="teams-app-upload-permission" height="400">


## Setup and run the sample locally

1. Ensure you have installed [Visual Studio Code](https://code.visualstudio.com/docs/setup/setup-overview) and [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) extension.
1. Clone the repository

    ```bash
    git clone https://github.com/yuting1008/expert-finder.git
    ```
1. Navigate to the `expert-finder` folder and open it with Visual Studio Code.
1. Select **File > Open Folder** in VS Code and choose this samples directory from the repo.
1. Using the extension, sign in with your Azure account and Microsoft 365 account where you have permissions to upload custom apps. \
    <img src="images/account-login-1.png" alt="account-login" height="400"> ⮕
    <img src="images/account-login-2.png" alt="account-login" height="400">
1. Select **Debug > Start Debugging** to run the app in a Teams web client. Once the debugging process starts, your default browser will open and navigate to the Teams web client, allowing you to test the application. \
    <img src="images/debug-in-Teams.png" alt="debug-in-teams" height="400">

## Deploy the app to Azure

1. Ensure the application is running locally to identify and resolve any potential bugs before proceeding.
1. Open the Teams Toolkit and select **Provision** under the Lifecycle section. This action will create the required resources in your Azure environment. \
    <img src="images/teams-toolkit-lifecycle.png" alt="teams-toolkit-lifecycle" height="400">
1. Select **Deploy** under the Lifecycle section. This action will deploy the base folder to Azure.
1. Select **Publish** under the Lifecycle section. This action will publish the application to Teams admin center.
1. Navigate to the [**Teams admin center**](https://admin.teams.microsoft.com/) and approve the app.
1. Open your Teams app store and install the app. And now you can refer the [next section](#use-the-app-in-teams-and-copilot-for-microsoft-365) to use the app.\
    <img src="images/install-app.png" alt="install-app" width="500">
1. If you encounter any errors when using the app, you can check the error log on Azure App Services. Navigate to the Web App you provisioned in Azure. Click **Monitoring** > **App Service Logs**.
1. Enable **Application logging (Filesystem)** and click **Save**. 
1. You will then be able to see the log of the web app in **log stream**. \
   <img src="images/enable-error-log.png" alt="enable-error-log" height="400">
> If you make any changes to the source code, click **Deploy** again to apply the changes to the deployed app.

## Use the app in Teams and Copilot for Microsoft 365

Navigate to the Microsoft Copilot for Microsoft 365 chat. Check the upper right of the chat user interface. You should see the Expert Finder as and agent. Click it and start using Expert-finder plugin.

#### SSO auth with consent popup when sign in button is clicked

On first use, a window requesting for Sign In pops-up and once SSO is completed, the user flow continues as usual.

#### Search based on skills and country in Copilot

<img src="images/m365-copilot-demo.gif" alt="Plugin" height="600">


Here are some sample prompts you can try : 
1) `Find experts with skill in Azure.`
2) `Find experts with skill in React and who are from Taipei.`
3) `Find experts with skill in Azure and available for interview.`

#### Try out message extension in chat

<img src="images/teams-message-extension-demo.gif" alt="Plugin" height="600">

## Further reading

- [Message extensions for Microsoft Copilot for Microsoft 365](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-message-extension-bot)
- [Get started with Microsoft Graph](https://developer.microsoft.com/en-us/graph)


<img src="https://pnptelemetry.azurewebsites.net/microsoft-teams-samples/samples/msgext-expert-finder-js" />
