using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Globalization;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using HoursApp;

class Program
{
    static async Task Main(string[] args)
    {
        var config = new ConfigurationBuilder()
            .AddUserSecrets<Program>()
            .Build();

        using var host = Host.CreateDefaultBuilder(args)
            .ConfigureServices((context, services) =>
            {
                services.AddLogging(); // Add logging service
                services.AddTransient<App>(); // App is your custom service
                services.AddSingleton<MailService>();
            })
            .Build();

        var app = host.Services.GetRequiredService<App>();
        await app.RunAsync();
    }    
}

public class App
{
    private readonly ILogger<App> _logger;
    private readonly MailService _mailService;

    public App(ILogger<App> logger, MailService mailService)
    {
        _logger = logger;
        _mailService = mailService;
    }

    public Task RunAsync()
    {
        _logger.LogInformation("App started!");

        #region Download approval
        // This is a manual process, unfortunately I need to register an app and asign it permissions to access Toyota's corporate email
        #endregion

        var today = DateTime.Today;
        int diff = today.DayOfWeek - DayOfWeek.Monday;
        if (diff < 0) diff += 7; // Ensure we go backwards to previous Monday if needed
        var monday = today.AddDays(-diff);

        #region Send Last week's approval
        // Then get previous week's Monday
        var previousMonday = monday.AddDays(-7);
        string lastWeekMondayFormattedDate = previousMonday.ToString("yyyy MM dd");
        string lastWeekApprovalfileName = $"Re Hours Week {lastWeekMondayFormattedDate} Monday.msg";
        string lastWeekApprovalFullPath = $"C:\\Users\\Bbronson\\Documents\\Time Tracking\\Approvals\\{lastWeekApprovalfileName}";

        if (File.Exists(lastWeekApprovalFullPath))
        {
            _logger.LogInformation($"Approval for week {lastWeekMondayFormattedDate} found.");

            // Send the Approval email to Employer
            _mailService.SendEmail("bmondragonbrozon@outlook.com", $"Toyota's Hours Approval", $"Hi,\nHere's Toyota's hours approval for last week and the week before last week.\nBest regards,\nBernardo", null, lastWeekApprovalFullPath);

            _logger.LogInformation($"Approval for week {lastWeekMondayFormattedDate} found.");

            // Rename approval file to mark it as sent
            string originalPath = lastWeekApprovalFullPath;
            string sentOnDate = DateTime.Today.ToString("yyyy-MM-dd");
            string folder = Path.GetDirectoryName(originalPath);

            string newFileName = $"Re Hours Week {lastWeekMondayFormattedDate} Monday (Sent {sentOnDate}).msg";
            string newPath = Path.Combine(folder, newFileName);

            File.Move(originalPath, newPath);
            _logger.LogInformation($"File renamed to: {newFileName}");
        }
        #endregion

        #region Generate this weeks hours file
        // Generate filename
        string formattedDate = monday.ToString("yyyy MM dd");
        string fileName = $"{formattedDate} Monday.xlsx";
        string outputFolder = "C:\\Users\\Bbronson\\Documents\\Hours";
        string fullPath = Path.Combine(outputFolder, fileName);

        // Check if file already exists
        if (!File.Exists(fullPath))
        {
            // Load configuration
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            // Set EPPlus license context
            var licenseContext = config["EPPlus:ExcelPackage:LicenseContext"];
            if (licenseContext == "NonCommercial")
                ExcelPackage.License.SetNonCommercialPersonal("Bernardo Mondragon Brozon");

            string templatePath = "C:\\Users\\Bbronson\\Documents\\Hours\\Template.xlsx";
            Directory.CreateDirectory(outputFolder);

            using var package = new ExcelPackage(new FileInfo(templatePath));
            var worksheet = package.Workbook.Worksheets[0]; // Access first sheet

            // Modify cell D9 with the week
            worksheet.Cells["D9"].Value = monday.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            // Modify cell E9
            worksheet.Cells["D12"].Value = 8;
            worksheet.Cells["D14"].Value = 8;
            worksheet.Cells["D16"].Value = 8;
            worksheet.Cells["D18"].Value = 8;
            worksheet.Cells["D20"].Value = 8;

            package.SaveAs(new FileInfo(fullPath));
            _logger.LogInformation($"Generate hours file: {fileName}.");
        }
        #endregion

        #region Send hours to supervisor
        // This is a manual process, unfortunately I need to register an app and asign it permissions to access Toyota's corporate email
        #endregion

        #region Send reporting email to Employer
        // Send the reporting email to Employer
        _mailService.SendEmail("bmondragonbrozon@outlook.com", $"Toyota's Hours", $"Hi,\nHere are this week's hours for Toyota's project.\nI'll send the approval when available.\nBest regards,\nBernardo", null, fullPath);

        _logger.LogInformation($"Hours for week {formattedDate} sent to employer.");
        #endregion

        return Task.CompletedTask;
    }    
}