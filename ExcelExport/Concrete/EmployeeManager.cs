using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelExport.Abstact;
using ExcelExport.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelExport.Concrete
{
    public class EmployeeManager : IEmployeeService
    {
        private readonly ExportContext _exportDbContext;
        public EmployeeManager(ExportContext exportDbContext)
        {
            _exportDbContext = exportDbContext;
        }
        public async Task<List<Employees>> GetEmployeeList()
        {
            try
            {
                var employees = await _exportDbContext.Employees.ToListAsync();
                return employees;
            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message);
            }
        }
    }
}
