# SharePoint Webhooks Azure Functions Deployment #

### Summary ###
This SharePoint Azure functions project is part of [Collaboration Foundry's](https://www.collaboration-foundry.com) SharePoint webhook demonstration solution.  It deploys two .NET (C#) Azure functions to an existing app service:
* A *WebHookHandler* function to be attached to a SharePoint list.  The function receives webhook notifications received from SharePoint and places them into an Azure storage queue.
* A *QueueHandler* function that is attached to an Azure storage queue.  It is responsible for performing the required processing for the webhook notifications.

### Applies to ###
- Azure.  Project has been tested with Azure Commercial and Azure Government.

### Prerequisites ###
- An Azure account over which you have administrative authority.
- All Azure components included in the two projects in this solution (storage account, app service and related components) have been successfully deployed.

----------

## Deployment #
*  Within Visual Studio, right-click on this project (*AzureFunctions* and click on *Publish...*
![Azure Function Visual Studio Publish Menu Option](https://www.collaboration-foundry.com/CFGitImages/AzureFunctionsVisualStudioPublishMenu.png)
* If you have not published before, the *Publish* dialog will appear as
![Azure Function Visual Studio Publish Dialog](https://www.collaboration-foundry.com/CFGitImages/AzureFunctionsVisualStudioPublishWindow.png)
* Within the *Publish* dialog, click on *Azure Function App* and then click on the *Select Existing* radio button.  (We are publishing to an existing Azure app service—the one we just deployed in the previous step.)  Then, click on the *Publish* button.
The *App Service* dialog will appear
* Within the *App Service* dialog, use the *Subscription* pull-down to select the appropriate Azure subscription (if you have multiple subscriptions).   Then, within the resource group hierarchy display, select the name of the app service to which you want to deploy the function (the one included in and deployed from this Visual Studio solution).
![Visual Studio App Service Dialog](https://www.collaboration-foundry.com/CFGitImages/AzureFunctionsVisualStudioPublishSelectAppService.png)
* Click the *OK* button on the *App Service* dialog.
At this point, Visual Studio will publish the functions, and the *Publish* dialog will contain your newly created publish profile:
![Visual Studio Azure Publish Profile](https://www.collaboration-foundry.com/CFGitImages/AzureFunctionsVisualStudioPublishProfile.png)
 
