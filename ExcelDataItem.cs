using System;

namespace SWT_LAB1_DE170303;

public partial class ExcelDataItem
{
    public object Item1 {get; set;}
    public object Item2 {get; set;}

    public ExcelDataItem(object Item1, object Item2) {this.Item1 = Item1; this.Item2 = Item2; }

    /// <summary>
    /// Calculate the sum of Item1 and Item2 of this ExcelDataItem
    /// </summary>
    /// <returns>Object contains the result data</returns>
    public object GetResult()
    {
        object result = new Exception("Receive Inputs are not Positive Integers");
        int operand2 = -1;
        bool isValidOperands = int.TryParse(Item1?.ToString(), out int operand1)
                            && int.TryParse(Item2?.ToString(), out operand2);
        if (isValidOperands)
        {
            try
            {
                if(operand1 < 0 || operand2 < 0) return result;
                result = operand1 + operand2;
            }
            catch (ArgumentNullException ex)
            {
                result = ex;
            }
            catch (FormatException ex)
            {
                result = ex;
            }
            catch (OverflowException ex)
            {
                System.Console.WriteLine("Exception found");
                return new OverflowException("Overflow because result exceed the int.MaxValue");
            }
            catch (Exception ex)
            {
                result = ex;
            }
        }
        return result;
    }
}