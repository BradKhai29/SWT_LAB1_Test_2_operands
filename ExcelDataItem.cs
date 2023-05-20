using System;

namespace SWT_LAB1_DE170303;

public partial class ExcelDataItem
{
    public object Item1 { get; set; }
    public object Item2 { get; set; }
    public object Result { get; set; }

    public ExcelDataItem(object Item1, object Item2) {this.Item1 = Item1; this.Item2 = Item2; }
}