using OfficeOpenXml;
using Result.Models;
using System.Collections.Generic;
using System.IO;

public class ExcelService
{
    public List<StudentResult> ReadExcelFile(string filePath)
    {
        var results = new List<StudentResult>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                results.Add(new StudentResult
                {
                    Serial = worksheet.Cells[row, 1].Text,
                    TermNo = worksheet.Cells[row, 2].Text,
                    Date = DateTime.Parse(worksheet.Cells[row, 3].Text),
                    Student = worksheet.Cells[row, 4].Text,
                    Father = worksheet.Cells[row, 5].Text,
                    Teacher = worksheet.Cells[row, 6].Text,
                    Class = worksheet.Cells[row, 7].Text,
                    Month1 = decimal.TryParse(worksheet.Cells[row, 8].Text, out var month1) ? month1 : 0,
                    Month2 = decimal.TryParse(worksheet.Cells[row, 9].Text, out var month2) ? month2 : 0,
                    Written = worksheet.Cells[row, 10].Text,
                    Wordlist = worksheet.Cells[row, 11].Text,
                    Viva = worksheet.Cells[row, 12].Text,
                    PresentationConversation = worksheet.Cells[row, 13].Text,
                    AttendanceBookReview = worksheet.Cells[row, 14].Text,
                    AssignmentFacilitators = worksheet.Cells[row, 15].Text,
                    Total = decimal.TryParse(worksheet.Cells[row, 16].Text, out var total) ? total : 0,
                    Obtained = decimal.TryParse(worksheet.Cells[row, 17].Text, out var obtained) ? obtained : 0,
                    Percentage = worksheet.Cells[row, 18].Text,
                    Result = worksheet.Cells[row, 19].Text,
                    PassingPercentage = worksheet.Cells[row, 20].Text,
                    Grade = worksheet.Cells[row, 21].Text
                });
            }
        }

        return results;
    }
}
