using Microsoft.EntityFrameworkCore;

namespace ExportExcel.Models
{
    public class AppDbContext:DbContext   
    {
        //apsettingdeki connectionstringi dbcontextoptionsile alıyoruz
        public AppDbContext(DbContextOptions options) : base(options) { 
        

        }

        public DbSet<Comments> Comments { get; set; }
    }
}
