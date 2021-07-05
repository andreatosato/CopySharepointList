using CopySharepointList.Configurations;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using System;
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

        private async Task<bool> CheckListExist(string siteId, string listId)
        {
            var site = await graphClient.Sites[siteId].Request().GetAsync();
            if (site == null)
                throw new System.ArgumentOutOfRangeException($"Sites {siteId} not exists");

            var list = graphClient.Sites[siteId].Lists[listId].Request().GetAsync();

            return list == null;
        }

        private async Task CreateListAsync(string siteId, string listName, string[] fields)
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

        private async Task ReadListsToCopyAsync()
        {
            var listsMaster = await graphClient.Sites[listConfigutation.SiteMasterId]
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

        private async IAsyncEnumerable<Dictionary<string, object>> CopyFromMaster(
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
                        if (row.Fields.AdditionalData.TryGetValue(f, out object currentValue))
                            rowReader.Add(f, currentValue);
                        else
                            Console.WriteLine($"Il field {f} is not present");
                    }
                    yield return rowReader;
                }
                if (listItems.NextPageRequest != null)
                    await listItems.NextPageRequest.GetAsync();
            }
            while (listItems.NextPageRequest != null);
        }

        private async Task AddRowToSlave(
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

        private async Task UpdateRowToSlave(
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

        /// <summary>
        /// Non utilizzato in questo momento la cancellazione
        /// </summary>
        /// <param name="siteSlaveId"></param>
        /// <param name="siteSlaveListId"></param>
        /// <param name="itemId"></param>
        /// <returns></returns>
        private async Task DeleteRowToSlave(
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

        /// <summary>
        /// Check row from first field of a row
        /// </summary>
        /// <param name="siteSlaveId"></param>
        /// <param name="siteSlaveListId"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private async Task<ListItem> FindRowToSlave(
            string siteSlaveId,
            string siteSlaveListId,
            Dictionary<string, object> row)
        {
            var field = row.FirstOrDefault();
            var queryOptions = new List<QueryOption>(1)
            {
                new QueryOption("filter", $"{field.Key} eq {field.Value}")
            };

            var oldRow = await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListId]
                .Items
                .Request(queryOptions)
                .GetAsync();

            return oldRow.FirstOrDefault();
        }


        public async Task ExecuteAsync()
        {
            await ReadListsToCopyAsync();
            var listFieldsNames = readerFields.GetListFields();
            var sitesToCopy = listConfigutation.SitesToCopy.Split(";");
            foreach (var site in sitesToCopy)
            {
                foreach (var currentListFields in listFieldsNames.Lists)
                {
                    if (!await CheckListExist(site, currentListFields.ListId))
                    {
                        await CreateListAsync(site, currentListFields.ListName, currentListFields.Fields);
                    }

                    await foreach (var row in CopyFromMaster(
                        listConfigutation.SiteMasterId,
                        currentListFields.ListId,
                        currentListFields.Fields))
                    {
                        var oldRow = await FindRowToSlave(site, currentListFields.ListId, row);
                        if (oldRow != null)
                            await UpdateRowToSlave(site, currentListFields.ListId, oldRow.Id, row);
                        else
                            await AddRowToSlave(site, currentListFields.ListId, row);
                    }

                }
            }
        }
    }
}
