using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace Google.Sheet.Web.Models
{
    public interface IGoogleSheetController
    {
        SheetsService GetSheetsService();
        Task<List<IList<object>>> ReadAsync();
        Task WriteAsync(string writeRange, ValueRange valueRange);
    }
}
