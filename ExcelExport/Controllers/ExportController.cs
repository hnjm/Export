using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelExport.Abstact;
using ExcelExport.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;

namespace ExcelExport.Controllers
{
    public class ExportController : Controller
    {
        private readonly IEmployeeService _employeeService;
        private readonly IMemoryCache _memoryCache;
        public ExportController(IEmployeeService employeeService, IMemoryCache memoryCache)
        {
            _employeeService = employeeService;
            _memoryCache = memoryCache;
        }

        [HttpGet]
        public ActionResult Index() => View();

        const string key = "allEmployeeData";
        [HttpPost]
        public async Task<ActionResult> ExportToCvs()
        {
            if (_memoryCache.Get(key) == null)
            {
                var employees = await _employeeService.GetEmployeeList();
                _memoryCache.Set(key, employees);
            }
            var employeeData = _memoryCache.Get<List<Employees>>(key);
            var builder = new StringBuilder();
            builder.AppendLine("EmployeeID,FirstName,LastName,Title,Address,City,PostalCode,TitleOfCourtesy,HomePhone,BirthDate");
            foreach (var employee in employeeData)
            {
                builder.AppendLine($"{employee.EmployeeID},{employee.FirstName},{employee.LastName}," +
                                   $"{employee.Title},{employee.Address}," +
                                   $"{employee.City},{employee.PostalCode},{employee.TitleOfCourtesy},{employee.HomePhone},{employee.BirthDate.ToLongDateString()}");
            }
            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "EmployeeList");
        }

        [HttpPost]
        public async Task<ActionResult> ExportToExcel()
        {
            if (_memoryCache.Get(key) == null)
            {
                var employees = await _employeeService.GetEmployeeList();
                _memoryCache.Set(key, employees);
            }
            var employeeData = _memoryCache.Get<List<Employees>>(key);
            using (var workbook=new XLWorkbook())
            {
                var workSheet = workbook.Worksheets.Add("Employees");
                var currentRow = 1;
                workSheet.Cell(currentRow, 1).Value = "EmployeeID";
                workSheet.Cell(currentRow, 2).Value = "First Name";
                workSheet.Cell(currentRow, 3).Value = "Last Name";
                workSheet.Cell(currentRow, 4).Value = "Title";
                workSheet.Cell(currentRow, 5).Value = "City";
                workSheet.Cell(currentRow, 6).Value = "Postal Code";
                workSheet.Cell(currentRow, 7).Value = "Title Of Courtesy";
                workSheet.Cell(currentRow, 8).Value = "Home Phone";
                workSheet.Cell(currentRow, 9).Value = "Birth Date";
                foreach (var employee in employeeData)
                {
                    currentRow++;
                    workSheet.Cell(currentRow, 1).Value = employee.EmployeeID;
                    workSheet.Cell(currentRow, 2).Value =employee.FirstName ;
                    workSheet.Cell(currentRow, 3).Value = employee.LastName;
                    workSheet.Cell(currentRow, 4).Value = employee.Title;
                    workSheet.Cell(currentRow, 5).Value = employee.City;
                    workSheet.Cell(currentRow, 6).Value = employee.PostalCode;
                    workSheet.Cell(currentRow, 7).Value = employee.TitleOfCourtesy;
                    workSheet.Cell(currentRow, 8).Value = employee.HomePhone;
                    workSheet.Cell(currentRow, 9).Value = employee.BirthDate;
                }
              
                using (var stream=new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmployeeList.xlsx");
                }
            }
        }

    }
}