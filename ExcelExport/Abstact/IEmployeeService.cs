using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelExport.Models;

namespace ExcelExport.Abstact
{
    public interface IEmployeeService
    {
         Task<List<Employees>> GetEmployeeList();
    }
}
