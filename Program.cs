namespace EmployeeManagement
{
    class Program
    {
        private static EmpServices _employeeService;

        static void Main(string[] args)
        {
            _employeeService = new EmpServices();
            bool exit = false;
            
            while (!exit)
            {
                DisplayMenu();
                string choice = Console.ReadLine();
                Console.WriteLine("");
                Console.WriteLine("------------------------------------------------------------------------------------");

                switch (choice)
                {
                    case "1":
                        AddEmployee();
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    case "2":
                        DeleteEmployee();
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    case "3":
                        EditEmployee();
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    case "4":
                        DisplayAllEmployees();
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    case "5":
                        AddEmployeesFromExcel();
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    case "6":
                        exit = true;
                        Console.WriteLine("Exiting application...");
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                    default:
                        Console.WriteLine("Invalid option. Please try again.");
                        Console.WriteLine("");
                        Console.WriteLine("------------------------------------------------------------------------------------");
                        break;
                }

                if (!exit)
                {
                    Console.WriteLine("\nPress any key to continue...");
                    Console.ReadKey();
                    Console.Clear();
                }
            }
        }

        static void DisplayMenu()
        {
            Console.WriteLine("=========== EMPLOYEE MANAGEMENT SYSTEM ===========");
            Console.WriteLine("==           1- add an employee                 ==");
            Console.WriteLine("==           2- delete an employee              ==");
            Console.WriteLine("==           3- edit an employee                ==");
            Console.WriteLine("==           4- display all employees           ==");
            Console.WriteLine("==     5 - add employees from an Excel sheet    ==");
            Console.WriteLine("==                   6 - exit                   ==");
            Console.WriteLine("==================================================");
            
            Console.Write("\nEnter your choice: ");
        }

        static void AddEmployee()
        {
            Console.WriteLine("\n=== ADD EMPLOYEE ===");
            
            try
            {
                var employee = new Employee();

                Console.Write("Enter Name: ");
                employee.Name = Console.ReadLine();

                Console.Write("Enter Department: ");
                employee.Department = Console.ReadLine();

                Console.Write("Enter Salary: ");
                if (decimal.TryParse(Console.ReadLine(), out decimal salary))
                {
                    employee.Salary = salary;
                }

                Console.Write("Enter Email: ");
                employee.Email = Console.ReadLine();

                Console.Write("Enter Phone: ");
                employee.Phone = Console.ReadLine();

                _employeeService.AddEmployee(employee);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding employee: {ex.Message}");
            }
        }

        static void DeleteEmployee()
        {
            Console.WriteLine("\n=== DELETE EMPLOYEE ===");
            Console.Write("Enter Employee ID to delete: ");
            
            if (int.TryParse(Console.ReadLine(), out int id))
            {
                _employeeService.DeleteEmployee(id);
            }
            else
            {
                Console.WriteLine("Invalid ID format.");
            }
        }

        static void EditEmployee()
        {
            Console.WriteLine("\n=== EDIT EMPLOYEE ===");
            Console.Write("Enter Employee ID to edit: ");
            
            if (int.TryParse(Console.ReadLine(), out int id))
            {
                var existingEmployee = _employeeService.GetEmployeeById(id);
                if (existingEmployee != null)
                {
                    var updates = new Dictionary<string, object>();

                    Console.WriteLine($"Current Name: {existingEmployee.Name}");
                    Console.Write("Enter New Name (or press Enter to keep current): ");
                    string name = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        updates["Name"] = name;
                    }

                    Console.WriteLine($"Current Department: {existingEmployee.Department}");
                    Console.Write("Enter New Department (or press Enter to keep current): ");
                    string dept = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(dept))
                    {
                        updates["Department"] = dept;
                    }

                    Console.WriteLine($"Current Salary: {existingEmployee.Salary}");
                    Console.Write("Enter New Salary (or press Enter to keep current): ");
                    string salaryStr = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(salaryStr) && decimal.TryParse(salaryStr, out decimal salary))
                    {
                        updates["Salary"] = salary;
                    }
                    
                    Console.WriteLine($"Current Email: {existingEmployee.Email}");
                    Console.Write("Enter New Email (or press Enter to keep current): ");
                    string email = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        updates["Email"] = email;
                    }

                    Console.WriteLine($"Current Phone: {existingEmployee.Phone}");
                    Console.Write("Enter New Phone (or press Enter to keep current): ");
                    string phone = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(phone))
                    {
                        updates["Phone"] = phone;
                    }

                    _employeeService.UpdateEmployee(id, updates);
                }
                else
                {
                    Console.WriteLine($"Employee with ID {id} not found.");
                }
            }
            else
            {
                Console.WriteLine("Invalid ID format.");
            }
        }

        static void DisplayAllEmployees()
        {
            Console.WriteLine("\n=== ALL EMPLOYEES ===");
            var employees = _employeeService.GetAllEmployees();

            if (employees.Count == 0)
            {
                Console.WriteLine("No employees found.");
                return;
            }
    
            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine("| ID   | Name                | Department    | Salary      | Email                          | Phone          |");
            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------------------------------");
    
            foreach (var emp in employees)
            {
                Console.WriteLine($"| {emp.Id,-4} | {emp.Name,-18} | {emp.Department,-12} | {emp.Salary,8} EGP   | {emp.Email,-28} | {emp.Phone,-13} |");
            }
    
            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------------------------------");
            Console.WriteLine($"Total Employees: {employees.Count}");
        }
        static void AddEmployeesFromExcel()
        {
            Console.WriteLine("\n=== ADD EMPLOYEES FROM EXCEL ===");
            Console.Write("Enter Excel file path: ");
            string filePath = Console.ReadLine();
            filePath = filePath.Trim('"');
            _employeeService.AddEmployeesFromExcel(filePath);
        }
    }
}