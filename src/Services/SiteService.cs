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
        private readonly GraphServiceClient graphServiceClient;
        private readonly IReaderFields readerFields;
        private readonly ListConfigurations listConfigutation;

        public SiteService(GraphServiceClient graphServiceClient,
            IOptions<ListConfigurations> listConfigutation,
            IReaderFields readerFields)
        {
            this.graphServiceClient = graphServiceClient;
            this.readerFields = readerFields;
            this.listConfigutation = listConfigutation?.Value;
        }

        public async Task<bool> CheckListExist(string siteId, string listId)
        {
            var site = await graphServiceClient.Sites[siteId].Request().GetAsync();
            if (site == null)
                throw new System.ArgumentOutOfRangeException($"Sites {siteId} not exists");

            var list = graphServiceClient.Sites[siteId].Lists[listId].Request().GetAsync();

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

            await graphServiceClient.Sites[siteId].Lists
                .Request()
                .AddAsync(listToCreate);
        }

        public async Task ReadListsToCopyAsync()
        {
            var listsMaster = await graphServiceClient.Sites[listConfigutation.ListMasterId]
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


        // TODO: draft
        public async Task CopyFromMaster(string siteMasterId,
            string siteMasterListId,
            string siteSlaveId,
            string siteSlaveListId,
            string[] fields)
        {
            var queryOptions = new List<QueryOption>(1)
            {
                new QueryOption("expand", "fields")
            };

            var listItems = await graphServiceClient
                           .Sites[siteMasterId]
                           .Lists[siteMasterListId]
                           .Items
                           .Request(queryOptions).GetAsync();
            do
            {
                foreach (var row in listItems.CurrentPage)
                {
                    var rowValues = new Dictionary<string, object>();
                    foreach (var f in fields)
                    {
                        rowValues.Add(f, row.Fields.AdditionalData[f]);
                    }
                }
                if (listItems.NextPageRequest != null)
                    await listItems.NextPageRequest.GetAsync();
            }
            while (listItems.NextPageRequest != null);
        }
    }
}
