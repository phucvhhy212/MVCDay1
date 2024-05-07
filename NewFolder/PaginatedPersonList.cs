using MVCDay1.Models;

namespace MVCDay1.NewFolder
{
    public class PaginatedPersonList
    {
        public IEnumerable<Person> Persons { get; set; }
        public int CountTotal { get; set; }
    }
}
