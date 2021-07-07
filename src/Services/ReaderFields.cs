using CopySharepointList.Configurations;
using CopySharepointList.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Linq;
using System.Text.Json;

namespace CopySharepointList.Services
{
    public class ReaderFields : IReaderFields
    {
        private readonly ListConfigurations listConfiguration;
        private readonly ILogger<ReaderFields> logger;
        private ListFields listFields;

        public ReaderFields(IOptions<ListConfigurations> listConfiguration, ILogger<ReaderFields> logger)
        {
            this.listConfiguration = listConfiguration?.Value;
            this.logger = logger;
        }

        public ListFields GetListFields()
        {
            if (string.IsNullOrEmpty(listConfiguration.ListsToCopy) || string.IsNullOrEmpty(listConfiguration.FieldToCopy))
                return null;

            if (listFields != null)
                return listFields;

            listFields = new ListFields();
            var lists = listConfiguration.ListsToCopy.Split(";");
            for (int i = 0; i < lists.Length; i++)
            {
                listFields.Lists.Add(new ListModel()
                {
                    ListName = lists[i],
                    Fields = listConfiguration.FieldToCopy.Split(";")[i]
                                        .Split(",")
                                        .ToArray(),
                    DisplayName = listConfiguration.DisplayNameToCopy.Split(";")[i]
                                        .Split(",")
                                        .ToArray(),
                });
            }
            logger.LogInformation("list and fields {0}", JsonSerializer.Serialize(listFields));
            return listFields;
        }

        public void SetListId(string listName, string listId)
        {
            if (listFields != null)
            {
                var l = listFields.Lists.FirstOrDefault(t => t.ListName == listName);
                l.ListId = listId;
                logger.LogInformation("list is setted: {0}", l.ListId);
            }
        }
    }
}
