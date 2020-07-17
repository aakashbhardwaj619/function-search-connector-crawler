using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Collections.Generic;

namespace SearchConnector.CustomSearchCrawler
{
  public static class CustomSearchCrawler
  {
    private static GraphHelper _graphHelper;
    private static ILogger _log;

    [FunctionName("CustomSearchCrawler")]
    public static async void Run([TimerTrigger("%schedule%")] TimerInfo myTimer, ExecutionContext context, ILogger log)
    {
      _log = log;

      _log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

      var appConfig = LoadAppSettings(context);
      if (appConfig == null)
      {
        _log.LogInformation("The configuration values are missing.");
        return;
      }

      var authProvider = new ClientCredentialAuthProvider(
        appConfig["appId"],
        appConfig["tenantId"],
        appConfig["secret"]
      );

      _graphHelper = new GraphHelper(authProvider);

      string lastModifiedTime = myTimer.ScheduleStatus.Last.ToUniversalTime().ToString("s") + "Z";
      _log.LogInformation("Last Trigger time: {0}", lastModifiedTime);

      //Get data to be indexed
      
      //Incremental Crawl
      var listItems = await _graphHelper.GetIncrementalListData(appConfig["siteId"], appConfig["listId"], lastModifiedTime);
      //Full Crawl
      //var listItems = await _graphHelper.GetFullListData(appConfig["siteId"], appConfig["listId"]);

      if (listItems.Count == 0)
      {
        _log.LogInformation("Search index is up to date...");
      }
      else
      {
        await UpdateSearchIndex(listItems, appConfig["connectionId"], appConfig["tenantId"]);
      }
    }

    //
    // Summary:
    //     Indexes data for custom search conenctor
    //
    // Parameters:
    //     listItems: Data that needs to be updated
    //     connectionId: Connection Id for which crawl needs to be done
    //     tenantId: Tenant Id
    private static async Task UpdateSearchIndex(dynamic listItems, string connectionId, string tenantId)
    {
      _log.LogInformation("Updating custom search connector index...");
      foreach (var listItem in listItems)
      {
        int lid = Convert.ToInt32(listItem.Id.ToString());
        _log.LogInformation("Updating SharePoint list item: {0}", lid);

        string[] parts = listItem.Fields.AdditionalData["Appliances"].ToString().Split(',');
        List<string> partsList = new List<string>(parts);

        AppliancePart part = new AppliancePart()
        {
          PartNumber = lid,
          Description = listItem.Fields.AdditionalData["Description"].ToString(),
          Name = listItem.Fields.AdditionalData["Title"].ToString(),
          Price = Convert.ToDouble(listItem.Fields.AdditionalData["Price"].ToString()),
          Inventory = Convert.ToInt32(listItem.Fields.AdditionalData["Inventory"].ToString()),
          Appliances = partsList
        };

        var newItem = new ExternalItem
        {
          Id = part.PartNumber.ToString(),
          Content = new ExternalItemContent
          {
            // Need to set to null, service returns 400
            // if @odata.type property is sent
            ODataType = null,
            Type = ExternalItemContentType.Text,
            Value = part.Description
          },
          Acl = new List<Acl>
          {
            new Acl {
              AccessType = AccessType.Grant,
              Type = AclType.Everyone,
              Value = tenantId,
              IdentitySource = "Azure Active Directory"
            }
          },
          Properties = part.AsExternalItemProperties()
        };

        await _graphHelper.AddOrUpdateItem(connectionId, newItem);

        _log.LogInformation("Item updated!");
      }
      _log.LogInformation("Custom search connector index update completed...");
    }

    private static IConfigurationRoot LoadAppSettings(ExecutionContext context)
    {
      var appConfig = new ConfigurationBuilder()
      .SetBasePath(context.FunctionAppDirectory)
      .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
      .AddEnvironmentVariables()
      .Build();

      if (string.IsNullOrEmpty(appConfig["appId"]) ||
          string.IsNullOrEmpty(appConfig["tenantId"]) ||
          string.IsNullOrEmpty(appConfig["secret"]) ||
          string.IsNullOrEmpty(appConfig["connectionId"]) ||
          string.IsNullOrEmpty(appConfig["siteId"]) ||
          string.IsNullOrEmpty(appConfig["listId"]))
      {
        return null;
      }

      return appConfig;
    }
  }
}
