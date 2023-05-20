using System.Diagnostics;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SWT_LAB1_DE170303;

public static class Program
{
    public static async Task Main()
    {
        ExcelReader excelReader = new
        (
            startRow: 2,
            endRow: 11,
            colA: 1,
            colB: 2
        );
        IList<ExcelDataItem> listOfItems = await excelReader.ReadFromExcel(fileName: "LAB1_DATA.xlsx");
        try
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            var results = await GetResultListAsync(listOfItems);
            stopwatch.Stop();
            foreach (var item in results) { await Console.Out.WriteLineAsync(item); }
            System.Console.WriteLine(stopwatch.ElapsedMilliseconds);
        }
        catch (Exception ex)
        {
            System.Console.WriteLine(ex);
        }
    }

    static async Task<string[]> GetResultListAsync(IList<ExcelDataItem> listOfItems)
    {
        IList<Task<string>> tasks = new List<Task<string>>();
        Parallel.ForEach
        (
            source: listOfItems,
            body: dataItem => tasks.Add(ResolveDataItemAsync(excelDataItem: dataItem))
        );
        return await Task.WhenAll(tasks);
    }

    static async Task<string> ResolveDataItemAsync(ExcelDataItem excelDataItem)
    {
        StringBuilder stringBuilder = new();
        await Task.Delay(1);
        try
        {
            int c = 0;
            int a = int.MinValue;
            int b = int.MinValue;
            bool testA = int.TryParse(excelDataItem.Item1.ToString(), out a);
            bool testB = int.TryParse(excelDataItem.Item2.ToString(), out b);
            if(!(testA && testB) || a < 0 || b < 0) throw new Exception("Receive Inputs are not Positive Integers"); 
            for (short i = 0; i < 10_000; i++)
            {
                for (short j = 0; j < 10_000; j++)
                {
                   c = a + b + c;
                }
            }
            if(c < 0) throw new OverflowException();
            excelDataItem.Result = c;
        }
        catch (Exception exception)
        {
            excelDataItem.Result = exception;
        }

        string resultString;
        if (excelDataItem.Result is int) resultString = stringBuilder.Append(excelDataItem.Result).Append(" : [Pass]").ToString();
        else resultString = stringBuilder.Append("Error : [").Append((excelDataItem.Result as Exception)!.Message).Append(']').ToString();
        return $"a = {excelDataItem.Item1,-12}, b = {excelDataItem.Item2,-12}, c = {resultString}";
    }
}
