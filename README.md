# Microsoft Search Connector Crawler using Azure Function

This timer triggered Azure function shows how Full Crawling of custom data for Microsoft search connector can be implemented at scheduled intervals.

## Prerequisites

### Setup the custom connector

- Create a connection
- Register a schema

> Refer to the sample [https://github.com/microsoftgraph/msgraph-search-connector-sample](https://github.com/microsoftgraph/msgraph-search-connector-sample) for setting up the connector and registering schema

### Create a SharePoint List 

In this step a SharePoint List will be created from which external data would be indexed to our search connector. Create a SharePoint list with the below columns (in accordance with the schema registered).

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
