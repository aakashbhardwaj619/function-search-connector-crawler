# Microsoft Search Connector Crawler using Azure Function

This timer triggered Azure function shows how crawling of external data for Microsoft Graph connector can be implemented at scheduled intervals. For this sample, the external data being indexed in Microsoft Search is coming from a SharePoint list but the approach can be modified to fetch data from any other custom service as well.

## Prerequisites

### Setup the custom connector

- Register an Azure AD App with `ExternalItem.ReadWrite.All` application permission on Microsoft Graph. Also create a Client Secret for the app and make a note of it along with App Id and Tenant Id.
- Create a connection
- Register a schema for the type of external data (For example, `Appliances` data as shown in below mentioned GitHub sample)
- Create a vertical
- Create a result type

> Refer to the sample [https://github.com/microsoftgraph/msgraph-search-connector-sample](https://github.com/microsoftgraph/msgraph-search-connector-sample) for setting up the connector and registering schema

### Create a SharePoint List 

Create a SharePoint list with the below columns (in accordance with the schema registered) that will serve as the external data source.

- `Title` of type `Single line of Text`
- `Description` of type `Single line of Text`
- `Appliances` of type `Single line of Text`
- `Inventory` of type `Number`
- `Price` of type `Number`

### Update configuration

- For local development create a `local.settings.json` file with the below configuration

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "<STORAGE_ACCOUNT_CONNECTION_STRING>",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet",
    "schedule": "<CRON_SCHEDULE>",
    "appId": "<APP_ID>",
    "tenantId": "<TENANT_ID>",
    "secret": "<APP_SECRET>",
    "connectionId": "<CONNECTION_ID>",
    "siteId": "<SITE_ID>",
    "listId": "<LIST_ID>"
  }
}
```

- For deployment to Azure Function App add the above properties as application settings in the function app
