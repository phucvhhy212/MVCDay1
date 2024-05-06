using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Person = MVCDay1.Models.Person;

namespace MVCDay1.Areas.NashTech.Controllers
{
    [Area("NashTech")]
    public class RookiesController : Controller
    {
        private static List<Person> GetAllPersons()
        {
            return new List<Person>
            {
                new Person
                {
                    Id = 1,
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
                    Id = 2,
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
                    Id = 3,
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
                    Id = 4,
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

        public IActionResult Index()
        {
            return View(GetAllPersons());
        }

        public IActionResult Create()
        {
            return View();
        }

        public IActionResult Edit(int id)
        {
            var persons = GetAllPersons();
            return View(persons.FirstOrDefault(p => p.Id == id));
        }

        public IActionResult Details(int id)
        {
            var persons = GetAllPersons();
            return View(persons.FirstOrDefault(p => p.Id == id));
        }

        public IActionResult OldestMember()
        {
            var members = GetAllPersons();
            return View("Index",members.Where(m => DateTime.Now.Year - m.DateOfBirth.Year == members.Max(x => DateTime.Now.Year - x.DateOfBirth.Year)));
        }

        public IActionResult MaleMembers()
        {
            return View("Index",GetAllPersons().Where(x => x.Gender == "Male"));
        }

        public IActionResult FullName()
        {
            return View(GetAllPersons().Select(x => x.FullName));
        }

        public IActionResult BirthYear(string option)
        {
            switch (option)
            {
                case "older":
                    return RedirectToAction("Older");
                case "younger":
                    return RedirectToAction("Younger");
                default:
                    return RedirectToAction("Equal");
            }
        }

        public IActionResult Delete(int id)
        {
            var persons = GetAllPersons();
            return View(persons.FirstOrDefault(p => p.Id == id));
        }

        public IActionResult Older()
        {
            var members = GetAllPersons();
            return View("Index",members.Where(x => x.DateOfBirth.Year > 2000));
        }

        public IActionResult Younger()
        {
            var members = GetAllPersons();
            return View("Index",members.Where(x => x.DateOfBirth.Year < 2000));
        }

        public IActionResult Equal()
        {
            var members = GetAllPersons();
            return View("Index",members.Where(x => x.DateOfBirth.Year == 2000));
        }

        public IActionResult Export()
        {
            var stream = ExportData(GetAllPersons());
            stream.Position = 0;
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PersonRecords.xlsx");
        }

        public MemoryStream ExportData(IEnumerable<Person> persons)
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage(stream);
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;
            //Header of table  

            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;
            workSheet.Cells[1, 1].Value = "First Name";
            workSheet.Cells[1, 2].Value = "Last Name";
            workSheet.Cells[1, 3].Value = "Gender";
            workSheet.Cells[1, 4].Value = "Date Of Birth";
            workSheet.Cells[1, 5].Value = "Birth Place";
            workSheet.Cells[1, 6].Value = "Is Graduated";
            //Body of table  
            //  
            int recordIndex = 2;
            foreach (var person in persons)
            {
                workSheet.Cells[recordIndex, 1].Value = person.FirstName;
                workSheet.Cells[recordIndex, 2].Value = person.LastName;
                workSheet.Cells[recordIndex, 3].Value = person.Gender;
                workSheet.Cells[recordIndex, 4].Value = person.DateOfBirth.ToString("MM/dd/yyyy");
                workSheet.Cells[recordIndex, 5].Value = person.BirthPlace;
                workSheet.Cells[recordIndex, 6].Value = person.IsGraduated;
                recordIndex++;
            }

            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();
            workSheet.Column(4).AutoFit();
            workSheet.Column(5).AutoFit();
            workSheet.Column(6).AutoFit();
            excel.Save();
            return stream;
        }

    }
}
