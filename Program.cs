using OfficeOpenXml;
using System;
using System.IO;
using File = System.IO.File;

static string FileSizeToString(long fileSizeInBytes)
{
    string[] suffix = { "B", "KB", "MB", "GB", "TB", "PB" };
    int suffixIndex = 0;

    while (fileSizeInBytes / 1024 > 0)
    {
        fileSizeInBytes /= 1024;
        suffixIndex++;
    }

    return string.Format("{0} {1}", fileSizeInBytes, suffix[suffixIndex]);
}

static void WriteAllFilesAndDirectoriesUnderPathToWorksheet(string path, ref ExcelWorksheet worksheet, int column = 1)
{
    var rootDirectory = new DirectoryInfo(path);

    var iterator = 1;
    foreach (DirectoryInfo currentDirectory in rootDirectory.GetDirectories())
    {
        worksheet.Cells[iterator++, column].Value = currentDirectory.FullName;
        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            worksheet.Cells[iterator, column].Value = file.FullName;
            worksheet.Cells[iterator, column + 1].Value = file.Extension;
            worksheet.Cells[iterator, column + 2].Value = FileSizeToString(file.Length);
            worksheet.Cells[iterator, column + 3].Value = file.Attributes.ToString();

            iterator++;
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