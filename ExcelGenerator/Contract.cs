using System;
using System.Collections.Generic;

namespace ExcelGenerator
{
    public class Contract
    {
        public string Code { get; internal set; }
        public DateTime StartDate { get; internal set; }
        public List<YearlyAmount> YearlyAmount { get; set; } = new List<YearlyAmount>();

    }
}