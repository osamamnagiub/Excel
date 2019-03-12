using System;
using System.Collections.Generic;

namespace _2019Generator
{
    public class Contract
    {
        public string Code { get; internal set; }
        public DateTime StartDate { get; internal set; }
        public List<YearlyAmount> YearlyAmount { get; set; } = new List<YearlyAmount>();
        public int Id { get; internal set; }
    }
}