using System;

namespace ExcelGenerator
{
    public class YearlyAmount
    {
        public double Amount { get; set; }
        public DateTime Date { get; set; }
        public double NPV { get; set; }
        public double Amortization { get; set; }
        public double Interest { get; set; }
    }
}