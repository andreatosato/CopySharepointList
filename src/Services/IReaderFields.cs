using CopySharepointList.Models;

namespace CopySharepointList.Services
{
    public interface IReaderFields
    {
        ListFields GetListFields();
        void SetListId(string listName, string listId);
    }
}
