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
        private readonly ILogger<HomeController> _logger;
        private readonly IGoogleSheetController _googleSheet;

        public HomeController(ILogger<HomeController> logger, IGoogleSheetController googleSheet)
        {
            _logger = logger;
            _googleSheet = googleSheet;
        }

        public async Task<IActionResult> Index()
        {
            var writeRange = "A344:P344";
            var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> { "Done", "DColatti", "8/4/2021", "Jimmy Colatti / W2/ FL", "Sales Manager" } } };
            await _googleSheet.WriteAsync(writeRange, valueRange);

            var results = await _googleSheet.ReadAsync();
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
    }
}
