using Excel2Word.Entities;
using LinqToExcel;
using System.Collections.Generic;
using System.Linq;

namespace Excel2Word.Activity
{
    public class ReadXls
    {
        internal List<Department> ReadFile(string filePath)
        {
            using (var excelQueryFactory = new ExcelQueryFactory(filePath))
            {
                var staffs = (from staff in excelQueryFactory.Worksheet("Сотрудники")
                             where staff["Табельный номер"] != null
                             select new
                             {
                                 Number = staff["Табельный номер"].ToString(),
                                 LastName = staff["Фамилия"].ToString(),
                                 FirstName = staff["Имя "].ToString(),
                                 Middlename = staff["Отчество"].ToString(),
                                 IdDepartment = staff["Отдел"].ToString(),
                             })
                             .ToList();

                var departments = (from department in excelQueryFactory.Worksheet("Отделы")
                                  where department["ИД отдела"] != null
                                  select new
                                  {
                                      IdDepartment = department["ИД отдела"].ToString(),
                                      DepartmentName = department["Наименование отдела"].ToString()
                                  })
                                  .ToList();

                var problems = (from problem in excelQueryFactory.Worksheet("Задачи")
                               where problem["ИД задачи"] != null
                               select new
                               {
                                   IdProblem = problem["ИД задачи"].ToString(),
                                   Number = problem["Табельный номер"].ToString()
                               })
                               .ToList();

                var staff_problems = (from s in staffs
                                     join p in problems on s.Number equals p.Number into pppp
                                     from ppp in pppp.DefaultIfEmpty()
                                     group ppp?.Number by s into pp
                                     select new { Staff = pp.Key, CountProblem = pp.Count(x => x != null) })
                                     .ToList();
                                     
                                     


                return staff_problems
                    .GroupJoin(departments,
                               c => c.Staff.IdDepartment,
                               s => s.IdDepartment,
                               (c, s) => new
                               {
                                   c.Staff,
                                   c.CountProblem,
                                   Department = s.FirstOrDefault()?.DepartmentName ?? "Не определен"
                               })
                    .Union(departments
                        .GroupJoin(staff_problems,
                                   s => s.IdDepartment,
                                   c => c.Staff.IdDepartment,
                                   (s, c) => new
                                   {
                                       c.FirstOrDefault()?.Staff,
                                       CountProblem = c.FirstOrDefault()?.CountProblem ?? 0,
                                       Department = s.DepartmentName
                                   }))
                    .GroupBy(k => k.Department)
                    .Select(g => new Department
                    {
                        DepartmentName = g.Key,
                        AllProblem = g.Sum(x => x.CountProblem),
                        Staffs = g
                                .Where(x => x.Staff != null)
                                .OrderByDescending(x => x.CountProblem)
                                .Select(x => new Staff
                                {
                                    LastName = x?.Staff.LastName ?? "",
                                    FirstName = x?.Staff.FirstName ?? "",
                                    Middlename = x?.Staff.Middlename ?? "",
                                    CountProblem = x?.CountProblem ?? 0
                                })
                                .ToList()
                    })
                    .OrderByDescending(x => x.AllProblem)
                    .ToList();
            }
        }
    }
}