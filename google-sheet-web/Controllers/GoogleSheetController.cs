using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Sheet.Web.Models;

namespace Google.Sheet.Web.Controllers
{
    public class GoogleSheetController : IGoogleSheetController
    {
        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        //private const string SpreadsheetId = "1FsgNIbugVQBZvwgWNx9Xn58RHopkkCtf0xsinGe23pQ"; // MY SHEET
        private const string SpreadsheetId = "1Mte4J-XcbuNT5Ln3dKLM7Mwy-RElJVgfOB3ACQG3Za0"; // TEST Onboarding
        private const string GoogleCredentialsFileName = "google-credentials.json";
        private const string ReadRange = "'2021 Enrollments'!A:P";

        private SpreadsheetsResource.ValuesResource ServiceValues { get; set; }

        public GoogleSheetController()
        {
            ServiceValues = GetSheetsService().Spreadsheets.Values;
        }

        public SheetsService GetSheetsService()
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

        public async Task<List<IList<object>>> ReadAsync()
        {
            var response = await ServiceValues.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values.ToList();

            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return new();
            }

            return values;
        }

        public async Task WriteAsync(string writeRange, ValueRange valueRange)
        {
            var update = ServiceValues.Update(valueRange, SpreadsheetId, writeRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            await update.ExecuteAsync();
        }
    }
}
