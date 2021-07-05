using CopySharepointList.Configurations;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CopySharepointList.Services
{
    class SiteService : ISiteService
    {
        private readonly GraphServiceClient graphClient;
        private readonly IReaderFields readerFields;
        private readonly ListConfigurations listConfigutation;

        public SiteService(GraphServiceClient graphClient,
            IOptions<ListConfigurations> listConfigutation,
            IReaderFields readerFields)
        {
            this.graphClient = graphClient;
            this.readerFields = readerFields;
            this.listConfigutation = listConfigutation?.Value;
        }

        public async Task<bool> CheckListExist(string siteId, string listId)
        {
            var site = await graphClient.Sites[siteId].Request().GetAsync();
            if (site == null)
                throw new System.ArgumentOutOfRangeException($"Sites {siteId} not exists");

            var list = graphClient.Sites[siteId].Lists[listId].Request().GetAsync();

            return list == null;
        }

        public async Task CreateListAsync(string siteId, string listName, string[] fields)
        {
            var listToCreate = new List
            {
                DisplayName = listName,
                Columns = new ListColumnsCollectionPage(),
                ListInfo = new ListInfo { Template = "genericList" }
            };

            for (int i = 0; i < fields.Length; i++)
            {
                listToCreate.Columns.Add(new ColumnDefinition
                {
                    Name = fields[i],
                    Text = new TextColumn()
                });
            }

            await graphClient.Sites[siteId].Lists
                .Request()
                .AddAsync(listToCreate);
        }

        public async Task ReadListsToCopyAsync()
        {
            var listsMaster = await graphClient.Sites[listConfigutation.ListMasterId]
                .Lists
                .Request()
                .GetAsync();

            var listsToCopy = readerFields.GetListFields().Lists.Select(x => x.ListName);

            do
            {
                foreach (var l in listsMaster.CurrentPage)
                {
                    if (listsToCopy.Contains(l.Name))
                        readerFields.SetListId(l.Name, l.Id);
                }

                if (listsMaster.NextPageRequest != null)
                    await listsMaster.NextPageRequest.GetAsync();
            }
            while (listsMaster.NextPageRequest != null);
        }

        public async IAsyncEnumerable<Dictionary<string, object>> CopyFromMaster(
            string siteMasterId,
            string siteMasterListId,
            string[] fields)
        {
            var queryOptions = new List<QueryOption>(1)
            {
                new QueryOption("expand", "fields")
            };

            var listItems = await graphClient
                           .Sites[siteMasterId]
                           .Lists[siteMasterListId]
                           .Items
                           .Request(queryOptions)
                           .GetAsync();
            do
            {
                foreach (var row in listItems.CurrentPage)
                {
                    var rowReader = new Dictionary<string, object>();
                    foreach (var f in fields)
                    {
                        rowReader.Add(f, row.Fields.AdditionalData[f]);
                    }
                    yield return rowReader;
                }
                if (listItems.NextPageRequest != null)
                    await listItems.NextPageRequest.GetAsync();
            }
            while (listItems.NextPageRequest != null);
        }

        public async Task AddRowToSlave(
            string siteSlaveId,
            string siteSlaveListId,
            Dictionary<string, object> row)
        {
            var listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = row
                }
            };

            await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListId]
                .Items
                .Request()
                .AddAsync(listItem);
        }

        public async Task UpdateRowToSlave(
           string siteSlaveId,
           string siteSlaveListId,
           string itemId,
           Dictionary<string, object> row)
        {
            var listItemUpdate = new FieldValueSet
            {
                AdditionalData = row
            };

            await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListId]
                .Items[itemId]
                .Fields
                .Request()
                .UpdateAsync(listItemUpdate);
        }

        public async Task DeleteRowToSlave(
           string siteSlaveId,
           string siteSlaveListId,
           string itemId)
        {
            await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListId]
                .Items[itemId]
                .Request()
                .DeleteAsync();
        }

        public async Task GetListItemsMasterVersion(
          string siteSlaveId,
          string siteSlaveListId)
        {
            var version = await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListId]
                .Items[""].Versions
                .Request().GetAsync();
        }
    }
}
