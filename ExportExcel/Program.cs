using ExportExcel.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);


ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Ticari olmayan bir proje için


// Add services to the container.
builder.Services.AddControllersWithViews();



builder.Services.AddDbContext<AppDbContext>(options =>
{
    options.UseSqlServer(builder.Configuration.GetConnectionString("SqlCon"));
}
);

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
