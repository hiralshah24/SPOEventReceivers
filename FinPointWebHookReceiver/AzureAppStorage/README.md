# SharePoint Webhooks Azure Storage Deployment #

### Summary ###
This SharePoint Azure Resource Template project is part of [Collaboration Foundry's](https://www.collaboration-foundry.com) SharePoint webhook demonstration solution.  It deploys the Azure storage account required to support the webhook's Azure app service.

### Applies to ###
- Azure.  Project has been tested with Azure Commercial and Azure Government.

### Prerequisites ###
- An Azure account over which you have administrative authority.

----------

## Deployment #
### Step 1: Configure Azure Storage Account Name
*  Within this project in Visual Studio, open the appropriate parameters file (Azure Commercial or Azure Government).  Edit the token's value, supplying a unique name for your Azure app service storage account: ![Set Azure Storage Account Name](https://www.collaboration-foundry.com/CFGitImages/AzureStorageAccountParameter.png)
Note: Your storage account name must be unique through all of Azure, so be creative.
*  Save the file.

### Step 2: Create Azure Resource Group, Deploy Azure Storage Account
* Within Visual Studio, right-click on this project and click on *New...*:  ![Creating New Azure Deployment](https://www.collaboration-foundry.com/CFGitImages/AzureCreateNewResourceGroup.png)
Note:  We recommend always creating a new deployment even if one you have previously used is displayed in the menu.
* To create a new resource group select *&lt;Create New...>* with the *Deploy to Resource Group* dialog.  Otherwise, you can select an existing resource group as needed. ![Select Azure Resource Group](https://www.collaboration-foundry.com/CFGitImages/AzureSelectResourceGroup.png)
* If you are creating a new resource group, provide its name in the *Resource Group Name* box: ![Select Azure Resource Group](https://www.collaboration-foundry.com/CFGitImages/AzureCreateNewResourceGroup.png)
* Click the *Create* button within the *Create Resource Group* dislog.  Visual Studio will create the resource group within Azure in real-time.
The *Deploy to Resource Group* dialog will now be available for you to select your resource group.
* The *appstoragedeploy.json* file should already be selected within the *Deployment template* box.  Within the *Template parameters file* box, select the parameters file that you previously edited:  ![Selecting Azure Template and Parameters](https://www.collaboration-foundry.com/CFGitImages/AzureSelectTemplateAndParameters.png)
* Click the *Deploy* button.  
You will see the deployment process start within Visual Studio (likely in a seperate window).  It will run for a minute or two.  When it is complete, you should see an output window in Visual Studio like this:  ![Successful Azure Storage Deployment](https://www.collaboration-foundry.com/CFGitImages/AzureStorageDeploymentSuccess.png)