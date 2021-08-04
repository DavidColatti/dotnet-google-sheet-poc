using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Google.Sheet.Web.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;

namespace Google.Sheet.Web.Controllers
{
    public class HomeController : Controller
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1FsgNIbugVQBZvwgWNx9Xn58RHopkkCtf0xsinGe23pQ";
        private const string GoogleCredentialsFileName = "google-credentials.json";
        private const string ReadRange = "Sheet1!A:B";
        private const string WriteRange = "A5:B5";

        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            var serviceValues = GetSheetsService().Spreadsheets.Values;
            await WriteAsync(serviceValues);

            var results = await ReadAsync(serviceValues);
            return View(results);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private static SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };
                return new SheetsService(serviceInitializer);
            }
        }

        private static async Task<List<IList<object>>> ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values.ToList();

            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return new();
            }

            return values;
        }

        private static async Task WriteAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { "stan", 18 } } }; // TODO: Hardcoded for POC/Testing Purposes

            var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            await update.ExecuteAsync();
        }
    }
}
