
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
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

static void WriteAllFilesAndDirectoriesUnderPathToWorksheet(ref ExcelWorksheet worksheet, ref int iterator, string path, int depth = 1, int directoryLevelDepth = 1)
{
    var rootDirectory = new DirectoryInfo(path);

    foreach (DirectoryInfo currentDirectory in rootDirectory.GetDirectories())
    {
        worksheet.Cells[iterator, 1].Value = currentDirectory.FullName;
        worksheet.Row(iterator).OutlineLevel = directoryLevelDepth;

        iterator++;

        if (depth > 1) WriteAllFilesAndDirectoriesUnderPathToWorksheet(ref worksheet, ref iterator, Path.Combine(path, currentDirectory.Name), depth - 1, directoryLevelDepth + 1);

        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            worksheet.Cells[iterator, 1].Value = file.FullName;
            worksheet.Cells[iterator, 2].Value = file.Extension;
            worksheet.Cells[iterator, 3].Value = FileSizeToString(file.Length);
            worksheet.Cells[iterator, 4].Value = file.Attributes.ToString();

            worksheet.Row(iterator).OutlineLevel = directoryLevelDepth + 1;
            worksheet.Row(iterator).Collapsed = true;

            iterator++;
        }
    }
}

static List<Tuple<string, long, string>> GetTopLargestFiles(string path, int depth = 1)
{
    var result = new List<Tuple<string, long, string>>();

    var rootDirectory = new DirectoryInfo(path);
    foreach (DirectoryInfo currentDirectory in rootDirectory.GetDirectories())
    {
        if (depth > 1) GetTopLargestFiles(Path.Combine(path, currentDirectory.Name), depth - 1);
        foreach (FileInfo file in currentDirectory.GetFiles())
        {
            result.Add(new(file.FullName, file.Length, file.Extension ));
        }
    }

    return result.OrderByDescending(x => x.Item2).ToList();
}

static List<(string, int, long)> GetTopLargestFilesStats(List<Tuple<string, long, string>> topLargestFiles)
{
    var uniqueExtensions = topLargestFiles.Select(x => x.Item3).Distinct().ToList();

    var result = new List<(string, int, long)>();
    foreach (var extension in uniqueExtensions)
    {
        var extensionCountInList = topLargestFiles.Count(x => x.Item3 == extension);
        var totalSizeOfFilesWithExtension = topLargestFiles
            .Where(x => x.Item3 == extension)
            .Sum(x => x.Item2);

        result.Add((extension, extensionCountInList, totalSizeOfFilesWithExtension));
    }

    return result;
}



ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Console.Write("Name of the directory to scan: ");
var path = Console.ReadLine();

Console.Write("Depth scan level (default is 1): ");
var depthLevelAsText = Console.ReadLine();

if (path == null || depthLevelAsText == null)
{
    Console.WriteLine("Something is wrong. The given parameters were registered as NULLs. Quitting...");
    return;
}

if (!int.TryParse(depthLevelAsText, out int depthLevel))
{
    Console.WriteLine("The given depth level has incorrect format (not a number)!");
    return;
}

if (depthLevel < 1)
{
    Console.WriteLine("The given depth level cannot be smaller than 1!");
    return;
}

if (!Directory.Exists(path))
{
    Console.WriteLine("The given directory does not exist!");
    return;
}

var excelSaveLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "lab1_180348.xlsx");
Console.WriteLine($"File will be saved under location: {excelSaveLocation}");

if (File.Exists(excelSaveLocation))
{
    try
    {
        File.Delete(excelSaveLocation);
    }
    catch(Exception e)
    {
        if (e.GetType() == typeof(UnauthorizedAccessException))
        {
            Console.WriteLine($"This application is not authorized to remove the old excel file.");
        }

        Console.WriteLine($"Cannot remove stale excel file ({excelSaveLocation}). Please do it manually.");
        return;
    }
}


var excelPackage = new ExcelPackage(new FileInfo(excelSaveLocation));
var worksheetWithFiles = excelPackage.Workbook.Worksheets.Add("Struktura katalogu");

var iterator = 1;
var depth = 2;
WriteAllFilesAndDirectoriesUnderPathToWorksheet(ref worksheetWithFiles, ref iterator, path, depth);

var worksheetWithStats = excelPackage.Workbook.Worksheets.Add("Statystyki");
var topLargestFiles = GetTopLargestFiles(path, depth);
topLargestFiles = topLargestFiles.GetRange(0, topLargestFiles.Count < 10 ? topLargestFiles.Count : 10).ToList();



for (int i = 0; i < topLargestFiles.Count; i++)
{
    worksheetWithStats.Cells[i + 1, 1].Value = topLargestFiles[i].Item1;
    worksheetWithStats.Cells[i + 1, 2].Value = topLargestFiles[i].Item3;
    worksheetWithStats.Cells[i + 1, 3].Value = topLargestFiles[i].Item2;
}

var topLargestFilesStats = GetTopLargestFilesStats(topLargestFiles);
for (int i = 0; i < topLargestFilesStats.Count; i++)
{
    worksheetWithStats.Cells[i + 1, 4].Value = topLargestFilesStats[i].Item1;
    worksheetWithStats.Cells[i + 1, 5].Value = topLargestFilesStats[i].Item2;
    worksheetWithStats.Cells[i + 1, 6].Value = topLargestFilesStats[i].Item3;
}



var chartWithExtensionAmount = (worksheetWithStats.Drawings.AddChart("Extensions Amount", eChartType.Pie3D) as ExcelPieChart);
if (chartWithExtensionAmount != null)
{
    chartWithExtensionAmount.Title.Text = "Procent rozszerzeń ilościowo";
    chartWithExtensionAmount.SetPosition(12, 1, 1, 1);
    chartWithExtensionAmount.SetSize(600, 600);

    ExcelPieChartSerie chartSerie = chartWithExtensionAmount.Series.Add($"E1:E{topLargestFilesStats.Count}", $"D1:D{topLargestFilesStats.Count}") as ExcelPieChartSerie; ;
    
    chartWithExtensionAmount.DataLabel.ShowCategory = true;
    chartWithExtensionAmount.DataLabel.ShowPercent = true;
}



var chartWithFileSizeByExtensions = (worksheetWithStats.Drawings.AddChart("File Size By Extensions", eChartType.Pie3D) as ExcelPieChart);
if (chartWithFileSizeByExtensions != null)
{
    chartWithFileSizeByExtensions.Title.Text = "Procent rozszerzeń wg rozmiaru";
    chartWithFileSizeByExtensions.SetPosition(12, 10, 12, 10);
    chartWithFileSizeByExtensions.SetSize(600, 600);

    ExcelPieChartSerie chartSerie = chartWithFileSizeByExtensions.Series.Add($"F1:F{topLargestFilesStats.Count}", $"D1:D{topLargestFilesStats.Count}") as ExcelPieChartSerie; ;

    chartWithFileSizeByExtensions.DataLabel.ShowCategory = true;
    chartWithFileSizeByExtensions.DataLabel.ShowPercent = true;
}

excelPackage.Save();
excelPackage.Dispose();