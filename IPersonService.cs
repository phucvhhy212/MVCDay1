using MVCDay1.Models;

namespace MVCDay1
{
    public interface IPersonService
    {
        IEnumerable<Person> GetAll();
        void Create(Person person);
        void Edit(Person person);
        void Delete(Person person);
        Person? Find(Guid id);
    }
}
