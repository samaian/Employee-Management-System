namespace EmployeeManagement;

public interface IEmployee
{
    void AddEmployee(Employee employee);
    bool DeleteEmployee(int id);
    bool UpdateEmployee(int id, Dictionary<string, object> updates);
    List<Employee> GetAllEmployees();
    Employee GetEmployeeById(int id);
    bool AddEmployeesFromExcel(string filePath);
}