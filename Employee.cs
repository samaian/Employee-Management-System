namespace EmployeeManagement;

public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Department { get; set; }
    public decimal Salary { get; set; }
    public string Email { get; set; }
    public string Phone { get; set; }

    public override string ToString()
    {
        return $"ID: {Id}, Name: {Name}, Department: {Department}, " +
               $"Salary: {Salary}," +
               $"Email: {Email}, Phone: {Phone}";
    }
}