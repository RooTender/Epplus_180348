using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

static void WriteAllFilesAndDirectoriesUnderPathToWorksheet(string path, ref ExcelWorksheet worksheet, int column = 1)
{
    var rootDirectory = new DirectoryInfo(path);

    var iterator = 1;
    foreach (DirectoryInfo currentDirectory in rootDirectory.GetDirectories())
    {
        worksheet.Cells[iterator++, column].Value = currentDirectory.FullName;
        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            worksheet.Cells[iterator++, column].Value = file.FullName;
        }
    }
}

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var excelSaveLocation = Path.Combine(Directory.GetCurrentDirectory(), "lab1_180348.xlsx");
Console.WriteLine($"File will be saved under location: {excelSaveLocation}");

if (File.Exists(excelSaveLocation))
{
    File.Delete(excelSaveLocation);
}

var excelPackage = new ExcelPackage(new FileInfo(excelSaveLocation));
var worksheetWithFiles = excelPackage.Workbook.Worksheets.Add("Struktura katalogu");

WriteAllFilesAndDirectoriesUnderPathToWorksheet(Directory.GetCurrentDirectory() + "\\..\\..\\..", ref worksheetWithFiles);

excelPackage.Save();
excelPackage.Dispose();