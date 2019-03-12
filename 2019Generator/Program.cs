using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace _2019Generator
{
    class Program
    {
        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }

        static void Main(string[] args)
        {
            List<Contract> contracts = new List<Contract>();

            using (FileStream stream = new FileStream("v.xlsx", FileMode.Open))
            {

                using (ExcelPackage excel = new ExcelPackage(stream))
                {
                    var sheet = excel.Workbook.Worksheets["Cell Sites"];
                    for (int i = 4; i < 7283; i++)
                    {
                        Contract contract = new Contract();
                        contract.Id = i - 3;
                        contract.Code = sheet.Cells[i, 1].Text;

                        contract.StartDate = FromExcelSerialDate(
                            int.Parse(sheet.Cells[i, 2].Value.ToString()));

                        var currentDate = contract.StartDate;
                        var diff = DateTime.Now.Year - currentDate.Year;
                        currentDate = currentDate.AddYears(diff);

                        for (int j = 53; j < 68; j++)
                        {
                            var amount = sheet.Cells[i, j].Value?.ToString();
                            if (string.IsNullOrEmpty(amount) || amount =="0") continue;

                            var NPV = sheet.Cells[i, j + 17].Value?.ToString();
                            if (string.IsNullOrEmpty(NPV)) NPV = "0";

                            var interest = sheet.Cells[i, j + 33].Value?.ToString();
                            if (string.IsNullOrEmpty(interest)) interest = "0";

                            var amortization = sheet.Cells[i, 85].Value?.ToString();
                            if (string.IsNullOrEmpty(amortization)) amortization = "0";



                            YearlyAmount year = new YearlyAmount()
                            {
                                Amount = double.Parse(amount),
                                NPV = double.Parse(NPV),
                                Amortization = double.Parse(amortization),
                                Interest = double.Parse(interest),
                                Date = currentDate
                            };

                            currentDate = currentDate.AddYears(1);
                            contract.YearlyAmount.Add(year);
                        }

                        contracts.Add(contract);
                    }

                }
            }

            File.Delete("out.xlsx");
            using (FileStream st = new FileStream("out.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                using (ExcelPackage newExcel = new ExcelPackage(st))
                {
                    newExcel.Workbook.Worksheets.Add("Worksheet1");
                    var workSheet = newExcel.Workbook.Worksheets["Worksheet1"];

                    var header = new object[] { "Id" , "Contract name", "Date", "Amount", "NPV", "Amortization", "Interest" };

                    workSheet.Cells[1, 1].LoadFromArrays(new List<object[]> { header });


                    var data = contracts.SelectMany(r =>
                r.YearlyAmount.Select(s => new object[]
                {
                    r.Id,
                            r.Code,
                            s.Date.ToShortDateString(),
                            s.Amount,
                            s.NPV,
                            s.Amortization,
                            s.Interest
                }));


                    workSheet.Cells[2, 1].LoadFromArrays(data);


                    newExcel.SaveAs(st);
                }

            }

        }
    }
}
