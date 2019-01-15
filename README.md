
# Moodle Assistant Bot for Microsoft Teams

The Moodle Assistant Bot for Microsoft Teams helps teachers and students answer questions about their courses, assignments, grades and other information in Moodle. The bot also sends Moodle notifications to students and teachers right within Teams!

## How to deploy the Moodle Assistant Bot to Azure

### Step 1: Register

* Go to the [Microsoft Application Registration Portal](https://aka.ms/MoodleBotRegistration) to regsiter a new app.

* Once you enter the name of the app (Eg. MoodleBot), simply click on the `Generate New Password` buttn and copy the one-time generated password. 

* Once the password is generated, copy the `Application ID` of the app along with the generated password from above. 

* Scroll down and click on Save.

### Step 2: Deploy

Click on the Deploy to Azure button and fill in the following details in the form:

* **Bot Application ID** - The Application ID from Step 1
* **Bot Application Password** - Application Password from Step 1
* **moodleURL** - The URL of your Moodle server
* **Azure Ad Application ID** - The Application ID saved in the *Setup* page of your Office 365 Moodle Plugin 
* **Azure Ad Application Key** - The Application Key saved in the *Setup* page of your Office 365 Moodle Plugin
* **Azure Ad Tenant** - The tenant name (xyz.onmicrosoft.com) of your Azure AD tenant

[![Deploy to Azure](http://azuredeploy.net/deploybutton.png)](https://aka.ms/DeployMoodleTeamsBot)

### Step 3: Configure

* Once the bot is deployed, go to the [Azure Portal](https://portal.azure.com), navigate to your bot's Resource Group and select the "Web App Bot".

* Copy the Messaging Endpoint highlighted in the `Overview` section (Ex: https://*provisioned-bot-name*.azurewebsites.net/api/messages), rename `messages` to `webhook` (Ex: https://*provisioned-bot-name*.azurewebsites.net/api/webook)

* Paste this endpoint to the `Bot Endpoint` field in the *Teams Settings* page of your Office 365 Moodle Plugin.
  
## Feedback

Thoughts? Questions? Ideas? Share them with us in our [Moodle+Teams Uservoice](https://microsoftteams.uservoice.com/forums/916759-moodle) channel!

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

&copy; Microsoft, Inc.  Code for this script is licensed under the MIT license.

Any Microsoft trademarks and logos included in these plugins are property of Microsoft and should not be reused, redistributed, modified, repurposed, or otherwise altered or used outside of this plugin.
