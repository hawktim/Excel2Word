using System.Collections.Generic;

namespace Excel2Word.Entities
{
    public class Department
    {
        public string DepartmentName { get; set; }
        public int AllProblem { get; set; }
        public List<Staff> Staffs { get; set; }
    }
}