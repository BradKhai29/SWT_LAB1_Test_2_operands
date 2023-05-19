namespace SWT_LAB1_DE170303;

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

public partial class ExcelReader
{
    public int StartRow { get; set; }
    public int EndRow  { get; set; }
    public int ColA { get; set; }
    public int ColB { get; set; }

    public ExcelReader(int startRow, int endRow, int colA, int colB)
    {
        StartRow = startRow;
        EndRow = endRow;
        ColA = colA;
        ColB = colB;
    }

    /// <summary>
    /// Read the Data from the Specific Excel file with given FilName
    /// </summary>
    /// <param name="fileName"></param>
    /// <returns>The Collection of ExcelDataItem contains Read Data</returns>
    public async Task<IList<ExcelDataItem>> ReadFromExcel(string fileName)
    {
        IList<ExcelDataItem> excelDataItems = new List<ExcelDataItem>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        try
        {
            SetDirectory();
            var fileInfo = new FileInfo(fileName: fileName);
            using var excelDocument = new ExcelPackage(fileInfo);
            await excelDocument.LoadAsync(fileInfo);
            var workSheet1 = excelDocument.Workbook.Worksheets[0];

            for (var currentRow = StartRow; currentRow <= EndRow; currentRow++)
            {
                var item = new ExcelDataItem
                (
                    Item1: workSheet1.Cells[currentRow, ColA].Value,
                    Item2: workSheet1.Cells[currentRow, ColB].Value
                );
                excelDataItems.Add(item);
            }
        }
        catch (Exception ex)
        {
            System.Console.WriteLine(ex);
        }

        return excelDataItems;
    }

    /// <summary>
    /// This method will Set the CurrentDirectory for this application.
    /// In situation the application is running in Visual Studio IDE or VSCode Editor
    /// </summary>
    private static void SetDirectory()
    {
        var currentDirectory = Environment.CurrentDirectory;
        var separator = Path.DirectorySeparatorChar;
        var directoryList = currentDirectory.Split(separator);

        //Checking if the currentDirectory is in Bin Directory or not
        bool inBinDirectory = false;
        foreach (var directory in directoryList)
        {
            if (directory.Equals("bin")) inBinDirectory = true;
        }
        if (inBinDirectory)
        {
            string newCurrentDirectory = Path.Combine(currentDirectory, "..", "..", "..");
            Environment.CurrentDirectory = newCurrentDirectory;
            Console.WriteLine(Environment.CurrentDirectory);
        }
    }
}