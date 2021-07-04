using System.Collections.Generic;

namespace CopySharepointList.Models
{
    public class ListModel
    {
        public string ListName { get; set; }
        public string ListId { get; set; }
        public string[] Fields { get; set; }
    }

    public class ListFields
    {
        public List<ListModel> Lists { get; set; }
    }
}
