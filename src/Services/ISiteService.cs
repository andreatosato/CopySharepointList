using System.Threading.Tasks;

namespace CopySharepointList.Services
{
    public interface ISiteService
    {
        Task ReadListsToCopyAsync();
        Task<bool> CheckListExist(string siteId, string listId);
        Task CreateListAsync(string siteId, string listName, string[] fields);
    }
}
