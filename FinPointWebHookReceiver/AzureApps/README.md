# SharePoint Webhooks Azure App Service Deployment #

### Summary ###
This SharePoint Azure Resource Template project is part of [Collaboration Foundry's](https://www.collaboration-foundry.com) SharePoint webhook demonstration solution.  It deploys an Azure app service and related components.

### Applies to ###
- Azure.  Project has been tested with Azure Commercial and Azure Government.

### Prerequisites ###
- An Azure account over which you have administrative authority.
- The storage project (part of this solution) has been deployed to Azure.

----------

## Deployment #
### Step 1: Configure Azure App Service Parameters
*  Within this project in Visual Studio, open the appropriate parameters file (Azure Commercial or Azure Government).
There are several values to be supplied (discussed below)  ![Set Azure Storage Account Name](https://www.collaboration-foundry.com/CFGitImages/AzureAppServiceParameters.png)
   * **appServicePlanName**:  The app service plan essentially defines the available resources and costs of the computing environment in which the webhook will run.   The plan type has been hard-coded in the deployment templates, but this parameter allows you to define the plan’s name.
   * **appServiceName**:  This is your pre-determined Azure app service name.
   * **appServiceUrlPrefix**:  This is your pre-determined Azure app service name.
   * **hostUrl**:  This is your pre-determined Azure app service name, suffixed with *azurewebsites.net* (Azure Commercial) or *azurewebsites.us* (Azure Government).
   * **UrlSuffix**:  *Azurewebsites.net*, if deploying to Azure Commercial.  *Azurewebsites.us*, if deploying to Azure Government.
   * **storageAccountName**:  This is the storage account name that you selected when you deployed the *AzureAppStorage* project.
   * **storageEndPointSuffix**:  *Core.windows.net*, if deploying to Azure Commercial.  *Core.usgovcloudapi.net*, if deploying to Azure Government.
   * **AppClientSecret**:  This is the client secret that you generated for your SharePoint add-in when you registered it as part of deployment for the *EventsHandlerAddIn* project.
   * **AppClientId**:  This is the client id that you generated for your SharePoint add-in when you registered it as part of deployment for the *EventsHandlerAddIn* project.
   * **O365TenantName**:  This is the SharePoint tenant name to which you are deploying.
### Step 2: Deploy Azure App Service
*  Within Visual Studio, right-click on the *AzureApps* project and click on *New...*.
![Create New Deployment for App Service](https://www.collaboration-foundry.com/CFGitImages/AzureAppServiceNewDeployment.png)

   The *Deploy to Resource Group* dialog will appear.
*  Select the existing resource group into which you previously deployed your storage account.  In addition, select the appropriate *Deployment template* and *Templare parameters file* (commerical or government).
![Select App Service Template and Parameters](https://www.collaboration-foundry.com/CFGitImages/AzureAppServiceSelectTemplateAndParameters.png)

*  Click the *Deploy* button.  
   You will see the deployment process initiate within Visual Studio (likely in a separate window).  It will run for a few minutes (longer than the previous storage deployment).  When it is complete, you should see an output window within Visual Studio.
![App Service Successfully Deployed](https://www.collaboration-foundry.com/CFGitImages/AzureAppServiceDeploymentSuccess.png)

Azure Notifications app service deployment is now complete.  

## Notes #
  1.  The consumption plan for app services is available in Azure Commercial but not Azure 
      Government.  
	  The sku and kind for the service (Microsoft.Web/serverfarms) for consumption is:
```
     "sku": {
        "name": "Y1",
        "tier": "Dynamic",
        "size": "Y1",
        "family": "Y",
        "capacity": 0
      },
      "kind": "functionapp",
```
   The sku for the service (Microsoft.Web/serverfarms) for the app service plan is:
```
      "sku": {
        "name": "S1",
        "tier": "Standard",
        "size": "S1",
        "family": "S",
        "capacity": 1
       },
      "kind": "app",
```

2.  The "WEBSITE_CONTENTAZUREFILECONNECTIONSTRING" and "WEBSITE_CONTENTSHARE" AppSettings are not included in the government 
     template because these parameters are only required for the commercial consumption plan.  The deployment will throw an 
	 error if the parameters are supplied in the government template.
