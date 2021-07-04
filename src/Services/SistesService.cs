using CopySharepointList.Configurations;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using System.Linq;
using System.Threading.Tasks;

namespace CopySharepointList.Services
{
    class SistesService : ISiteService
    {
        private readonly GraphServiceClient graphServiceClient;
        private readonly IReaderFields readerFields;
        private readonly ListConfigurations listConfigutation;

        public SistesService(GraphServiceClient graphServiceClient,
            IOptions<ListConfigurations> listConfigutation,
            IReaderFields readerFields)
        {
            this.graphServiceClient = graphServiceClient;
            this.readerFields = readerFields;
            this.listConfigutation = listConfigutation?.Value;
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
    }
}
