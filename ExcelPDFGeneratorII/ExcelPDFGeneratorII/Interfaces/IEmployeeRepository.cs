using ExcelPDFGeneratorII.Models;

namespace ExcelPDFGeneratorII.Interfaces
{
    public interface IEmployeeRepository
    {
        Task<IEnumerable<Employee>> GetAllEmployeesAsync();
    }
}
