using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Person = MVCDay1.Models.Person;

namespace MVCDay1.Areas.NashTech.Controllers
{
    [Area("NashTech")]
    public class RookiesController : Controller
    {
        //public static List<Person> data = GetAllPersons();

        private readonly IPersonService _personService;

        public RookiesController(IPersonService personService)
        {
            _personService = personService;
        }
        public IActionResult Index()
        {
            return View(_personService.GetAll());
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(Person p)
        {
            if (ModelState.IsValid)
            {
                p.Id = Guid.NewGuid();
                _personService.Create(p);
            }

            return RedirectToAction("Index");
        }

        public IActionResult Edit(Guid id)
        {
            var data = _personService.GetAll();
            return View(data.FirstOrDefault(p => p.Id == id));
        }

        [HttpPost]
        public IActionResult Edit(Person p)
        {
            if (ModelState.IsValid)
            {
                _personService.Delete(_personService.Find(p.Id));
                _personService.Create(p);
            }
            return RedirectToAction("Index");
        }
        public IActionResult Details(Guid id)
        {
            return View(_personService.Find(id));
        }

        public IActionResult OldestMember()
        {
            var data = _personService.GetAll();
            return View("Index",data.Where(m => DateTime.Now.Year - m.DateOfBirth.Year == data.Max(x => DateTime.Now.Year - x.DateOfBirth.Year)));
        }

        public IActionResult MaleMembers()
        {
            var data = _personService.GetAll();
            return View("Index",data.Where(x => x.Gender == "Male"));
        }

        public IActionResult FullName()
        {
            var data = _personService.GetAll();
            return View(data.Select(x => x.FullName));
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

        public IActionResult Delete(Guid id)
        {
            var data = _personService.GetAll();
            return View(data.FirstOrDefault(p => p.Id == id));
        }

        [HttpPost]
        public IActionResult DeletePerson(Guid id)
        {
            var personToRemove = _personService.Find(id);
            _personService.Delete(personToRemove);
            TempData["personName"] = personToRemove.FullName;
            return RedirectToAction("ConfirmDelete");
        }

        public IActionResult ConfirmDelete()
        {
            return View();
        }
        public IActionResult Older()
        {
            var data = _personService.GetAll();
            return View("Index",data.Where(x => x.DateOfBirth.Year > 2000));
        }

        public IActionResult Younger()
        {
            var data = _personService.GetAll();
            return View("Index",data.Where(x => x.DateOfBirth.Year < 2000));
        }

        public IActionResult Equal()
        {
            var data = _personService.GetAll();
            return View("Index",data.Where(x => x.DateOfBirth.Year == 2000));
        }

        public IActionResult Export()
        {
            var data = _personService.GetAll();
            var stream = ExportData(data);
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
