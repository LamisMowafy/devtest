var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllersWithViews();  // For MVC and views
builder.Services.AddControllers();          // For Web API controllers

var app = builder.Build();

// Middleware to handle exceptions, static files, routing, etc.
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

// MVC routing for views
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Upload}/{id?}");

// API routing for API controllers
app.MapControllers();  // This will map API controllers (i.e., controllers that are decorated with [ApiController])

app.Run();

