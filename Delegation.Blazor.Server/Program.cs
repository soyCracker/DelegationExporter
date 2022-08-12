using BlazorDownloadFile;
using Delegation.Blazor.Server.Data;
using Delegation.Service.Services;
using MudBlazor.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddMudServices();
builder.Services.AddSingleton<WeatherForecastService>();
builder.Services.AddScoped<RecordService>();
builder.Services.AddScoped<ExportService>();
builder.Services.AddScoped(provider=>new PDFService(""));
builder.Services.AddBlazorDownloadFile();
builder.Services.AddScoped<ZipService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();

app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();
