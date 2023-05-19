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
        ExcelDataItem dataItem = new(0, 0);
        StringBuilder stringBuilder = new();
        await Task.Delay(1);
        for (short i = 0; i < 10_000; i++)
        {
            // var currentThread = Environment.CurrentManagedThreadId;
            // await Console.Out.WriteLineAsync($"Thread = {currentThread}");
            for (short j = 0; j < 10_000; j++)
            {
                //DataItem.Item1 = excelDataItem.Item1 + excelDataItem.Item2
                dataItem.Item1 = excelDataItem.GetResult();
                //DataItem.Item2 = excelDataItem.Item1 + excelDataItem.Item2 + DataItem.Item1
                dataItem.Item2 = dataItem.GetResult();
                if(dataItem.GetResult() is not int) break;
            }
            if(dataItem.GetResult() is not int) break;
        }
        var result = dataItem.GetResult();
        string resultString;

        if (result is int) resultString = stringBuilder.Append(result).Append(" : [Pass]").ToString();
        else resultString = stringBuilder.Append("Error : [").Append((result as Exception)!.Message).Append(']').ToString();

        return $"a = {excelDataItem.Item1,-12}, b = {excelDataItem.Item2,-12}, c = {resultString}";
    }
}
