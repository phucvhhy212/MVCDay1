using MVCDay1.Models;
using MVCDay1.NewFolder;

namespace MVCDay1
{
    public class PersonService:IPersonService
    {
        private static List<Person> data = GetAllPersons();
        
        public PaginatedPersonList GetAll(int? page,int? recordsPerPage)
        {
            return new PaginatedPersonList
            {
                CountTotal = data.Count,
                Persons = page == null && recordsPerPage == null ? data : data.Skip((page.Value-1) * recordsPerPage.Value).Take(recordsPerPage.Value)
            };
        }

        public void Create(Person person)
        {
            data.Add(person);
        }

        public void Edit(Person person)
        {
            data.Remove(data.FirstOrDefault(x => x.Id == person.Id));
            data.Add(person);
        }

        public void Delete(Person person)
        {
            data.Remove(person);
        }

        public Person? Find(Guid id)
        {
            return data.FirstOrDefault(x=>x.Id == id);
        }

        private static List<Person> GetAllPersons()
        {
            return new List<Person>
            {
                new Person
                {
                    Id = Guid.NewGuid(),
                    FirstName = "Huy1",
                    LastName = "Phuc1",
                    BirthPlace = "Ha Noi",
                    DateOfBirth = DateTime.Today,
                    Gender = "Female",
                    IsGraduated = false,
                    PhoneNumber = "7329074222"
                },
                new Person
                {
                    Id = Guid.NewGuid(),
                    FirstName = "Huy2",
                    LastName = "Phuc2",
                    BirthPlace = "HN",
                    DateOfBirth = DateTime.Parse("2000-11-02"),
                    Gender = "Male",
                    IsGraduated = false,
                    PhoneNumber = "7329074222"
                },
                new Person
                {
                    Id = Guid.NewGuid(),
                    FirstName = "Huy3",
                    LastName = "Phuc3",
                    BirthPlace = "Ha Noi",
                    DateOfBirth = DateTime.Parse("2000-12-02"),
                    Gender = "Male",
                    IsGraduated = false,
                    PhoneNumber = "7329074222"
                },
                new Person
                {
                    Id = Guid.NewGuid(),
                    FirstName = "Huy4",
                    LastName = "Phuc4",
                    BirthPlace = "Ha Noi",
                    DateOfBirth = DateTime.Parse("1999-12-02"),
                    Gender = "Male",
                    IsGraduated = false,
                    PhoneNumber = "7329074222"
                }
            };
        }
    }
}
