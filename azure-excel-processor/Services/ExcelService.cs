using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelProcessor.Models;

namespace ExcelProcessor.Services
{
    public class ExcelService
    {
        public List<XrfShot> ReadShotsFromExcel(byte[] content)
        {
            var shots = new List<XrfShot>();
            using (var stream = new MemoryStream(content))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet.RowsUsed().Skip(1); // Skip header

                // Map columns by header name (case-insensitive)
                var headerRow = worksheet.Row(1);
                var colMap = new Dictionary<string, int>();
                for (int i = 1; i <= headerRow.LastCellUsed().Address.ColumnNumber; i++)
                {
                    colMap[headerRow.Cell(i).Value.ToString().ToLower().Trim()] = i;
                }

                foreach (var row in rows)
                {
                    try
                    {
                        var shot = new XrfShot
                        {
                            Reading = GetInt(row, colMap, new[] { "reading", "shot #", "shot" }),
                            Component = GetString(row, colMap, new[] { "component" }),
                            Side = GetString(row, colMap, new[] { "side" }),
                            Color = GetString(row, colMap, new[] { "color" }),
                            Substrate = GetString(row, colMap, new[] { "substrate", "subtrate" }), // Handle typo in sample
                            Condition = GetString(row, colMap, new[] { "condition" }),
                            RoomNumber = GetString(row, colMap, new[] { "room number" }),
                            RoomType = GetString(row, colMap, new[] { "room type" }),
                            Floor = GetString(row, colMap, new[] { "floor" }),
                            Result = GetString(row, colMap, new[] { "result" }),
                            Pbc = GetDouble(row, colMap, new[] { "pbc", "pb", "lead", "pb mg/cm2" })
                        };
                        shots.Add(shot);
                    }
                    catch (Exception)
                    {
                        // Skip invalid rows
                    }
                }
            }
            return shots;
        }

        public byte[] CreateSummaryExcel(List<ComponentSummary> summaries, string title)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Results");
                worksheet.Cell(1, 1).Value = title.ToUpper();
                worksheet.Range(1, 1, 1, 5).Merge().Style.Font.SetBold();

                var headers = new[] { "Component", "Count", "Negative%", "Positive%", "Lead Content" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(2, i + 1).Value = headers[i];
                    worksheet.Cell(2, i + 1).Style.Font.SetBold();
                }

                for (int i = 0; i < summaries.Count; i++)
                {
                    var s = summaries[i];
                    worksheet.Cell(i + 3, 1).Value = s.Component;
                    worksheet.Cell(i + 3, 2).Value = s.Count;
                    worksheet.Cell(i + 3, 3).Value = s.NegativePercentage;
                    worksheet.Cell(i + 3, 4).Value = s.PositivePercentage;
                    worksheet.Cell(i + 3, 5).Value = s.LeadContent;
                }

                worksheet.Columns().AdjustToContents();

                using (var ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }

        public byte[] CreateConflictingExcel(List<XrfShot> shots, string title)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Conflicting Results");
                worksheet.Cell(1, 1).Value = title.ToUpper();
                worksheet.Range(1, 1, 1, 14).Merge().Style.Font.SetBold();

                var headers = new[] { "Reading", "Component", "Side", "Color", "Substrate", "Condition", "Room Number", "Room Type", "Floor", "Result", "Pbc" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(2, i + 1).Value = headers[i];
                    worksheet.Cell(2, i + 1).Style.Font.SetBold();
                }

                for (int i = 0; i < shots.Count; i++)
                {
                    var s = shots[i];
                    worksheet.Cell(i + 3, 1).Value = s.Reading;
                    worksheet.Cell(i + 3, 2).Value = s.Component;
                    worksheet.Cell(i + 3, 3).Value = s.Side;
                    worksheet.Cell(i + 3, 4).Value = s.Color;
                    worksheet.Cell(i + 3, 5).Value = s.Substrate;
                    worksheet.Cell(i + 3, 6).Value = s.Condition;
                    worksheet.Cell(i + 3, 7).Value = s.RoomNumber;
                    worksheet.Cell(i + 3, 8).Value = s.RoomType;
                    worksheet.Cell(i + 3, 9).Value = s.Floor;
                    worksheet.Cell(i + 3, 10).Value = s.Result;
                    worksheet.Cell(i + 3, 11).Value = s.Pbc;
                }

                worksheet.Columns().AdjustToContents();

                using (var ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    return ms.ToArray();
                }
            }
        }

        private string GetString(IXLRow row, Dictionary<string, int> map, string[] keys)
        {
            foreach (var key in keys)
            {
                if (map.TryGetValue(key.ToLower(), out int colIdx))
                    return row.Cell(colIdx).Value.ToString();
            }
            return string.Empty;
        }

        private int GetInt(IXLRow row, Dictionary<string, int> map, string[] keys)
        {
            var val = GetString(row, map, keys);
            return int.TryParse(val, out int result) ? result : 0;
        }

        private double GetDouble(IXLRow row, Dictionary<string, int> map, string[] keys)
        {
            var val = GetString(row, map, keys);
            return double.TryParse(val, out double result) ? result : 0;
        }
    }
}

