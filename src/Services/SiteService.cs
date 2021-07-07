using CopySharepointList.Configurations;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace CopySharepointList.Services
{
    class SiteService : ISiteService
    {
        private readonly GraphServiceClient graphClient;
        private readonly IReaderFields readerFields;
        private readonly ILogger<SiteService> logger;
        private readonly ListConfigurations listConfigutation;

        public SiteService(GraphServiceClient graphClient,
            IOptions<ListConfigurations> listConfigutation,
            IReaderFields readerFields,
            ILogger<SiteService> logger)
        {
            this.graphClient = graphClient;
            this.readerFields = readerFields;
            this.logger = logger;
            this.listConfigutation = listConfigutation?.Value;
        }

        private async Task<bool> CheckListExist(string siteId, string listName)
        {
            var site = await graphClient.Sites[siteId].Request().GetAsync();
            if (site == null)
                throw new System.ArgumentOutOfRangeException($"Sites {siteId} not exists");

            try
            {
                await graphClient.Sites[siteId].Lists[listName].Request().GetAsync();
                logger.LogWarning("site-id: {0} and list-name {1} exist", siteId, listName);
                return true;
            }
            catch (Exception)
            {
                logger.LogWarning("site-id: {0} and list-name {1} NOT exist", siteId, listName);
                return false;
            }
        }

        private async Task CreateListAsync(string siteId, string listName, string[] fields, string[] displayName)
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
                    DisplayName = displayName[i],
                    Text = new TextColumn()
                });
            }

            await graphClient.Sites[siteId].Lists
                .Request()
                .AddAsync(listToCreate);

            logger.LogInformation("list created in: site {0}, listname: {1}, with fields {2} and description {3}", siteId, listName, JsonSerializer.Serialize(fields), JsonSerializer.Serialize(displayName));
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
                    logger.LogInformation("list to copy: {0}", l.Name);
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
            string siteMasterListName,
            string[] fields)
        {
            var queryOptions = new List<QueryOption>(1)
            {
                new QueryOption("expand", "fields")
            };

            var listItems = await graphClient
                           .Sites[siteMasterId]
                           .Lists[siteMasterListName]
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
                        logger.LogInformation("Read {0}", JsonSerializer.Serialize(row.Fields.AdditionalData));
                        if (row.Fields.AdditionalData.TryGetValue(f, out object currentValue))
                            rowReader.Add(f, currentValue);
                        else
                            logger.LogError($"Il field {f} is not present");
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

            logger.LogInformation("Add record {0} to list {1} in site {2}", JsonSerializer.Serialize(row), siteSlaveListId, siteSlaveId);

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

            logger.LogInformation("Update record {0} to list {1} in site {2}", JsonSerializer.Serialize(row), siteSlaveListId, siteSlaveId);

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
        /// <param name="siteSlaveListName"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private async Task<ListItem> FindRowToSlave(
            string siteSlaveId,
            string siteSlaveListName,
            Dictionary<string, object> row)
        {
            // https://stackoverflow.com/questions/49172556/microsoft-graph-filtering-in-sdk-c-sharp/49172694
            var field = row.FirstOrDefault();
            logger.LogInformation($"Finding record for {field.Key} eq '{field.Value}");

            var oldRow = await graphClient
                .Sites[siteSlaveId]
                .Lists[siteSlaveListName]
                .Items
                .Request()
                .Filter($"fields/{field.Key} eq '{field.Value}'")
                .Header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
                .GetAsync();

            var find = oldRow.FirstOrDefault();
            if (find == null)
                return null;

            logger.LogInformation($"Finded record for {field.Key} eq '{field.Value}");
            return find;
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
                    if (!await CheckListExist(site, currentListFields.ListName))
                    {
                        await CreateListAsync(site, currentListFields.ListName, currentListFields.Fields, currentListFields.DisplayName);
                    }

                    await foreach (var row in CopyFromMaster(
                        listConfigutation.SiteMasterId,
                        currentListFields.ListName,
                        currentListFields.Fields))
                    {
                        var oldRow = await FindRowToSlave(site, currentListFields.ListName, row);
                        if (oldRow != null)
                            await UpdateRowToSlave(site, currentListFields.ListName, oldRow.Id, row);
                        else
                            await AddRowToSlave(site, currentListFields.ListName, row);
                    }

                }
            }
        }
    }
}
