using MVCDay1.Models;
using MVCDay1.NewFolder;

namespace MVCDay1
{
    public interface IPersonService
    {
        PaginatedPersonList GetAll(int? page = null, int? recordPerPage = null);
        void Create(Person person);
        void Edit(Person person);
        void Delete(Person person);
        Person? Find(Guid id);
    }
}
