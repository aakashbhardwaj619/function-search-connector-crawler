using System;
using Microsoft.Graph;
using System.Threading.Tasks;
using System.Net.Http;
using System.Collections.Generic;

namespace SearchConnector.CustomSearchCrawler
{
  public class GraphHelper
  {
    private GraphServiceClient _graphClient;

    public GraphHelper(IAuthenticationProvider authProvider)
    {
      _graphClient = new GraphServiceClient(authProvider);
    }

    public async Task AddOrUpdateItem(string connectionId, ExternalItem item)
    {
      try
      {
        // The SDK's auto-generated request builder uses POST here,
        // which isn't correct. For now, get the HTTP request and change it
        // to PUT manually.
        var putItemRequest = _graphClient.External.Connections[connectionId]
            .Items[item.Id].Request().GetHttpRequestMessage();

        putItemRequest.Method = HttpMethod.Put;
        putItemRequest.Content = _graphClient.HttpProvider.Serializer.SerializeAsJsonContent(item);

        var response = await _graphClient.HttpProvider.SendAsync(putItemRequest);
        if (!response.IsSuccessStatusCode)
        {
          throw new ServiceException(
              new Error
              {
                Code = response.StatusCode.ToString(),
                Message = "Error indexing item."
              }
          );
        }
      }
      catch (Exception exception)
      {
        throw new Exception("Error: ", exception);
      }
    }

    public async Task<IListItemsCollectionPage> GetFullListData(string siteId, string listId)
    {
      var queryOptions = new List<QueryOption>()
      {
        new QueryOption("expand", "fields")
      };
      var items = await _graphClient.Sites[siteId].Lists[listId].Items
        .Request(queryOptions)
        .GetAsync();

      return items;
    }

    public async Task<IListItemsCollectionPage> GetIncrementalListData(string siteId, string listId, string lastModified)
    {
      string filterQuery = $"fields/Modified ge '{lastModified}'";

      var queryOptions = new List<QueryOption>()
      {
        new QueryOption("expand", "fields"),
        new QueryOption("filter", filterQuery)
      };
      var items = await _graphClient.Sites[siteId].Lists[listId].Items
        .Request(queryOptions)
        .GetAsync();

      return items;
    }
  }
}