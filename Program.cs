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
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .AddUserSecrets<Program>()
            .Build();

        using var host = Host.CreateDefaultBuilder(args)
            .ConfigureServices((context, services) =>
            {
                services.AddLogging(); // Add logging service
                services.AddSingleton<IConfiguration>(config);
                services.AddTransient<App>(); // App is your custom service
                services.AddSingleton<MailService>();
            })
            .Build();

        var app = host.Services.GetRequiredService<App>();
        if (args.Length > 0)
        {
            // Validate Input
            if (args.Length % 2 != 0)
            {
                Console.WriteLine("Invalid arguments. Please provide pairs of date and hours.");
                return;
            }
            // Process DayException arguments
            Console.WriteLine("Processing DayException arguments...");
            List<DayException> dayExceptions = new List<DayException>();
            for (int i = 0; i < args.Length; i += 2)
            {
                if (!DateOnly.TryParse(args[i], out DateOnly date))
                {
                    Console.WriteLine($"Invalid date format: {args[i]}");
                    return;
                }
                if (!int.TryParse(args[i + 1], out int hours) || hours < 0)
                {
                    Console.WriteLine($"Invalid hours value: {args[i + 1]}");
                    return;
                }
                // Create a DayException object
                var dayException = new DayException
                {
                    Date = date,
                    Hours = hours
                };
                dayExceptions.Add(dayException);
            }
            // Run the app with DayException list   
            try
            {
                await app.RunAsync(dayExceptions);
            }
            catch (Exception ex)
            {
                // Log the exception
                var logger = host.Services.GetRequiredService<ILogger<App>>();
                logger.LogError(ex, "An error occurred while running the app with DayException arguments.");
            }
            finally
            {
                Console.WriteLine("App execution completed with DayException arguments.");
            }
        }
        else
        {
            // Normal run
            await app.RunAsync();
        }
    }    
}

public class App
{
    private readonly ILogger<App> _logger;
    private readonly MailService _mailService;
    private readonly IConfiguration _config;

    public App(ILogger<App> logger, MailService mailService, IConfiguration config)
    {
        _logger = logger;
        _mailService = mailService;
        _config = config;
    }

    internal Task RunAsync(List<DayException>? dayExceptions = null)
    {
        _logger.LogInformation("App started!");

        DateTime today = DateTime.Today;
        int diff = today.DayOfWeek - DayOfWeek.Monday;
        if (diff < 0) diff += 7; // Ensure we go backwards to previous Monday if needed
        DateTime monday = today.AddDays(-diff);
        // Then get previous week's Monday
        var previousMonday = monday.AddDays(-7);
        string lastWeekMondayFormattedDate = previousMonday.ToString("yyyy MM dd");

        #region Download approval
        GmailAttachmentDownloader.DownloadLatestHoursWeekAttachment(
            gmailUser: _config["PersonalFromAddress"],
            appPassword: _config["GoogleAppPassword"],
            downloadFolder: _config["ApprovalsLocation"],
            fileName: $"Re_ Hours Week {lastWeekMondayFormattedDate} Monday.eml"
        );
        #endregion

        #region Send Last week's approval to employer
        string extension = _config["ApprovalsExtension"];
        string lastWeekApprovalfileName = $"Re_ Hours Week {lastWeekMondayFormattedDate} Monday.{extension}";
        string lastWeekApprovalFullPath = $"{_config["ApprovalsLocation"]}\\{lastWeekApprovalfileName}";

        if (File.Exists(lastWeekApprovalFullPath))
        {
            _logger.LogInformation($"Approval for week {lastWeekMondayFormattedDate} found.");

            // Send the Approval email to Employer
            _mailService.SendEmail(_config["PersonalFromAddress"], _config["PersonalFromDisplay"], _config["EmployerEmail"], null, $"Toyota's Hours Approval", $"Hi,\n\nHere's Toyota's hours approval for last week and the week before last week.\n\nBest regards,\n\nBernardo", null, lastWeekApprovalFullPath);

            // Rename approval file to mark it as sent
            string originalPath = lastWeekApprovalFullPath;
            string sentOnDate = DateTime.Today.ToString("yyyy-MM-dd");
            string folder = Path.GetDirectoryName(originalPath);

            string newFileName = $"Re_ Hours Week {lastWeekMondayFormattedDate} Monday (Sent {sentOnDate}).{extension}";
            string newPath = Path.Combine(folder, newFileName);

            // If the destination file exists, delete it before renaming
            if (File.Exists(newPath))
            {
                File.Delete(newPath);
            }

            File.Move(originalPath, newPath);
            _logger.LogInformation($"File renamed to: {newFileName}");
        }
        else
        {
            _logger.LogInformation($"Approval not found");
        }
        #endregion

        #region Generate this weeks hours file
        // Generate filename
        string formattedDate = monday.ToString("yyyy MM dd");
        string fileName = $"{formattedDate} Monday.xlsx";
        string outputFolder = _config["HoursLocation"];
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

            string templatePath = $"{_config["TemplateLocation"]}\\Template.xlsx";
            Directory.CreateDirectory(outputFolder);

            using var package = new ExcelPackage(new FileInfo(templatePath));
            var worksheet = package.Workbook.Worksheets[0]; // Access first sheet

            // Modify cell D9 with the week
            var dateCellValue = monday.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            worksheet.Cells["D9"].Value = dateCellValue;

            // Modify cell E9
            worksheet.Cells["D12"].Value = 8;
            worksheet.Cells["D14"].Value = 8;
            worksheet.Cells["D16"].Value = 8;
            worksheet.Cells["D18"].Value = 8;
            worksheet.Cells["D20"].Value = 8;

            if (dayExceptions != null && dayExceptions.Count > 0)
            {
                foreach (var dayException in dayExceptions)
                {
                    // Find the row for the specific date
                    int row = 12 + (dayException.Date.DayOfWeek - DayOfWeek.Monday) * 2; // Assuming Monday starts at row 12
                    worksheet.Cells[row, 4].Value = dayException.Hours; // Column D is index 4
                }
            }

            package.SaveAs(new FileInfo(fullPath));
            _logger.LogInformation($"Generate hours file: {fileName}.");
        }
        else
        {
            _logger.LogInformation($"Hours file already exists. No need to generate new file.");
        }
        #endregion

        #region Send reporting email to Employer
        _mailService.SendEmail(_config["PersonalFromAddress"], _config["PersonalFromDisplay"], _config["EmployerEmail"], null, $"Toyota's Hours", $"Hi,\n\nHere are this week's hours for Toyota's project.\nI'll send the approval when available.\n\nBest regards,\n\nBernardo", null, fullPath);

        _logger.LogInformation($"Hours for week {formattedDate} sent to employer.");
        #endregion

        #region Send hours to supervisor
        _mailService.SendEmail(_config["PersonalFromAddress"], _config["PersonalFromDisplay"], _config["CorporateAddress"], null, $"Hours Week {formattedDate} Monday", $"Hi Chris,\n\nHere are my hours for last week.\n\nBest regards,\n\nBernardo", null, fullPath);

        _logger.LogInformation($"Hours for week {formattedDate} sent out to {_config["CorporateAddress"]} to trigger PowerAutomate flow.");
        #endregion

        return Task.CompletedTask;
    }    
}