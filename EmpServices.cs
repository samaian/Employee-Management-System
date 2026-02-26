namespace EmployeeManagement;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

public class EmpServices : IEmployee
{
   private List<Employee> _employees;
    private int _index = 1;

    public EmpServices()
    {
        _employees = new List<Employee>();
    }


    #region ADD METHOD

    public void AddEmployee(Employee employee)
    {
        if (employee != null)
        {
            employee.Id = _index++;
            _employees.Add(employee);
            Console.WriteLine("Employee added");
        }
        else
        {
            Console.WriteLine("Employee not added");
        }
    }

    #endregion


    #region DELETE METHOD
    public bool DeleteEmployee(int id)
    {
        var employee = GetEmployeeById(id);

        if (employee != null)
        {
            _employees.Remove(employee);
            return true;
        }
        else
        {
            Console.WriteLine("Employee not found");
            return false;
        }
    }


    #endregion

    #region UPDATE METHOD

    public bool UpdateEmployee(int id, Dictionary<string, object> updates)
    {
        var existingEmployee = GetEmployeeById(id);

        if (existingEmployee != null)
        {
            int updateCount = 0;

            foreach (var update in updates)
            {
                try
                {
                    switch (update.Key.ToLower())
                    {
                        case "name":
                            if (update.Value != null)
                            {
                                existingEmployee.Name = update.Value.ToString();
                                updateCount++;
                                Console.WriteLine($"Updated Name");
                            }
                            break;

                        case "department":
                            if (update.Value != null)
                            {
                                existingEmployee.Department = update.Value.ToString();
                                updateCount++;
                                Console.WriteLine($"Updated Department");
                            }

                            break;

                        case "salary":
                            if (decimal.TryParse(update.Value?.ToString(), out decimal salary))
                            {
                                existingEmployee.Salary = salary;
                                updateCount++;
                                Console.WriteLine($"Updated Salary");
                            }

                            break;

                        case "email":
                            if (update.Value != null && update.Value.ToString().Contains("@"))
                            {
                                existingEmployee.Email = update.Value.ToString();
                                updateCount++;
                                Console.WriteLine($"Updated Email");
                            }

                            break;

                        case "phone":
                            if (update.Value != null)
                            {
                                existingEmployee.Phone = update.Value.ToString();
                                updateCount++;
                                Console.WriteLine($"Updated Phone");
                            }

                            break;

                        default:
                            Console.WriteLine($"Property '{update.Key}' not found");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error updating {update.Key}: {ex.Message}");
                }
            }

            if (updateCount > 0)
            {
                Console.WriteLine($"Employee with ID {id} updated successfully. ({updateCount} fields updated)");
                return true;
            }

            Console.WriteLine($"ℹNo fields were updated for employee with ID {id}.");
            return false;
        }

        Console.WriteLine($"Employee with ID {id} not found.");
        return false;
    }

    #endregion


    #region GET ALL EMPLOYEES

    public List<Employee> GetAllEmployees()
    {
        return _employees.ToList();
    }

    #endregion


    #region Get Employee By Id

    public Employee GetEmployeeById(int id)
    {
        return _employees.Find(e => e.Id == id);
    }

    #endregion

    #region Add Employees From Excel
   public bool AddEmployeesFromExcel(string filePath)
{
    try
    {
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found.");
            return false;
        }

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null || worksheet.Dimension == null)
            {
                Console.WriteLine("Excel file is empty or invalid.");
                return false;
            }

            var excelHeaders = new List<string>();
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                var header = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(header))
                    excelHeaders.Add(header);
            }
            
            var employeeProps = typeof(Employee).GetProperties().Select(p => p.Name).Where(p => p != "Id").ToList();
            bool headersMatch = excelHeaders.SequenceEqual(employeeProps, StringComparer.OrdinalIgnoreCase);
            
            if (!headersMatch)
            {
                Console.WriteLine("Invalid Excel format");
                return false;
            }

            int successCount = 0;
            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                var employee = new Employee();
                
                employee.Name = worksheet.Cells[row, 1].Value?.ToString() ?? "";
                employee.Department = worksheet.Cells[row, 2].Value?.ToString() ?? "";
                
                decimal.TryParse(worksheet.Cells[row, 3].Value?.ToString(), out decimal salary);
                employee.Salary = salary;
                
               
                employee.Email = worksheet.Cells[row, 4].Value?.ToString() ?? "";
                employee.Phone = worksheet.Cells[row, 5].Value?.ToString() ?? "";

                if (!string.IsNullOrEmpty(employee.Name))
                {
                    AddEmployee(employee);
                    successCount++;
                }
            }

            Console.WriteLine($"Added {successCount} employees");
            return successCount > 0;
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error reading Excel file: {ex.Message}");
        return false;
    }
}
    #endregion
}