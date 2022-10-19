using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
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
        worksheet.Cells[iterator, column].Value = currentDirectory.FullName;
        worksheet.Row(iterator).OutlineLevel = 1;

        iterator++;

        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            worksheet.Cells[iterator, column].Value = file.FullName;
            worksheet.Cells[iterator, column + 1].Value = file.Extension;
            worksheet.Cells[iterator, column + 2].Value = FileSizeToString(file.Length);
            worksheet.Cells[iterator, column + 3].Value = file.Attributes.ToString();

            worksheet.Row(iterator).OutlineLevel = 2;
            worksheet.Row(iterator).Collapsed = true;

            iterator++;
        }
    }
}

static List<KeyValuePair<string, long>> GetTopLargestFiles(string path)
{
    var result = new List<KeyValuePair<string, long>>();

    var rootDirectory = new DirectoryInfo(path);
    foreach (DirectoryInfo currentDirectory in rootDirectory.GetDirectories())
    {
        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            result.Add(new( file.FullName, file.Length ));
        }
    }

    return result.OrderByDescending(x => x.Value).ToList();
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
var rootPath = Directory.GetCurrentDirectory() + "\\..\\..\\..";

WriteAllFilesAndDirectoriesUnderPathToWorksheet(rootPath, ref worksheetWithFiles);

var worksheetWithStats = excelPackage.Workbook.Worksheets.Add("Statystyki");
var topLargestFiles = GetTopLargestFiles(rootPath).GetRange(0, 10).ToList();

for(int i = 0; i < topLargestFiles.Count; i++)
{
    worksheetWithStats.Cells[i + 1, 1].Value = topLargestFiles[i].Key;
}

excelPackage.Save();
excelPackage.Dispose();