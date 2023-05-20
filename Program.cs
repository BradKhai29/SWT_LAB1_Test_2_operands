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
        IList<ExcelDataItem> excelItems = await excelReader.ReadFromExcel(fileName: "LAB1_DATA.xlsx");
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();
        await GetResultListAsync(excelItems);
        stopwatch.Stop();
        System.Console.WriteLine("{0, -20} | {1, -20} | {2}", "a", "b", "c");
        foreach (var item in excelItems)
        {
            await Console.Out.WriteLineAsync($"{item.Item1, -20} | {item.Item2, -20} | {item.Result}");
        }
        System.Console.WriteLine(stopwatch.ElapsedMilliseconds);
    }

    static async Task GetResultListAsync(IList<ExcelDataItem> listOfItems)
    {
        IList<Task> tasks = new List<Task>();
        Parallel.ForEach
        (
            source: listOfItems,
            body: dataItem => tasks.Add(ResolveDataItemAsync(excelDataItem: dataItem))
        );
        await Task.WhenAll(tasks);
    }

    static async Task ResolveDataItemAsync(ExcelDataItem excelDataItem)
    {
        await Task.Delay(1);
        try
        {
            int c = 0;
            int a = int.MinValue;
            int b = int.MinValue;
            bool testA = int.TryParse(excelDataItem.Item1.ToString(), out a);
            bool testB = int.TryParse(excelDataItem.Item2.ToString(), out b);
            if (!(testA && testB) || a < 0 || b < 0)
            {
                excelDataItem.Result = "Receive Inputs are not Positive Integers";
                return;
            }
            for (short i = 0; i < 10_000; i++)
            {
                for (short j = 0; j < 10_000; j++)
                {
                    c = a + b + c;
                }
            }
            if (c < 0) excelDataItem.Result = "Overflow happened";
            else excelDataItem.Result = c;
        }
        catch (Exception)
        {
            excelDataItem.Result = "Invalid";
        }
    }
}
