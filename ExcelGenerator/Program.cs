using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numeric;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream stream = new FileStream("excel.xlsx", FileMode.Open))
            {
                using (ExcelPackage excel = new ExcelPackage(stream))

                {
                    var sheet = excel.Workbook.Worksheets["Site sharing"];
                    List<Contract> contracts = new List<Contract>();
                    var date = new DateTime(1900, 1, 1);

                    for (int i = 3; i <= 989; i++)

                    {
                        Contract contract = new Contract();
                        contract.Code = sheet.Cells[i, 3].Text;

                        contract.StartDate = FromExcelSerialDate(int.Parse(sheet.Cells[i, 5].Value.ToString()));
                        var currentDate = contract.StartDate;

                        for (int j = 17; j <= 33; j++)
                        {

                            YearlyAmount yearlyAmount = new YearlyAmount();
                            yearlyAmount.Date = currentDate;
                            var amount = sheet.Cells[i, j].Text;
                            if (!string.IsNullOrEmpty(amount))
                            {
                                double result = 0;
                                double.TryParse(amount.ToString(), out result);
                                yearlyAmount.Amount = result;



                                contract.YearlyAmount.Add(yearlyAmount);
                                currentDate = currentDate.AddYears(1);
                            }

                        }

                        contracts.Add(contract);
                    }

                    foreach (var item in contracts)
                    {
                        var count = item.YearlyAmount.Count;

                        for (int i = 0; i < count; i++)
                        {
                            var valuesToNpv = item.YearlyAmount
                                .GetRange(i, count - i)
                                .Select(s => s.Amount);

                            item.YearlyAmount[i].NPV = Financial.Npv(0.181520, valuesToNpv);
                            item.YearlyAmount[i].Amortization = item.YearlyAmount[0].NPV / count;
                            item.YearlyAmount[i].Interest = item.YearlyAmount[i].NPV * 0.181520;

                        }

                    }

                    using (FileStream st = new FileStream("excel2.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        using (ExcelPackage newExcel = new ExcelPackage(st))
                        {
                            newExcel.Workbook.Worksheets.Add("Worksheet1");

                            var data = contracts.SelectMany(r =>
                        r.YearlyAmount.Select(s => new object[]
                        {
                            r.Code,
                            s.Date.ToShortDateString(),
                            s.Amount,
                            s.NPV,
                            s.Amortization,
                            s.Interest
                        }));

                            var workSheet = newExcel.Workbook.Worksheets["Worksheet1"];

                            workSheet.Cells[1, 1].LoadFromArrays(data);


                            newExcel.SaveAs(st);
                        }

                    }
                }

            }
        }

        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }
    }
}
