using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Fast_Report_Web_Api_First.Model
{
    public class ApplicationDbContextClass : DbContext
    {
        public ApplicationDbContextClass(DbContextOptions<ApplicationDbContextClass> option):base(option)
        {

        }
        public DbSet<Person> Person { get; set; }
    }
}
