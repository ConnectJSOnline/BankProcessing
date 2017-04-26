using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using FileHelpers;
using OfficeOpenXml;
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;

namespace BankProcessing
{
    class Program
    {
        static List<PaymentRecords> Payments = new List<PaymentRecords>();
        static void Main(string[] args)
        {
            ReadFiles();
            SaveExcel();
            Console.WriteLine(" -------  Processing Completed ----- Press Any Key to close -------");
            Console.ReadKey();
        }

        private static void SaveExcel()
        {
            FileInfo output = new FileInfo("output.xlsx");
            if (output.Exists) output.Delete();
            using (ExcelPackage package = new ExcelPackage(output))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add("output");
                ws.Cells[1, 1].Value = "File";
                ws.Cells[1, 2].Value = "Date";
                ws.Cells[1, 3].Value = "ReciptNo";
                ws.Cells[1, 4].Value = "Cheque";
                ws.Cells[1, 5].Value = "Amount";
                for (int i = 0; i < Payments.Count; i++)
                {
                    ws.Cells[i + 2, 1].Value = Path.GetFileName(Payments[i].FileName.FullName);
                    ws.Cells[i + 2, 2].Value = Payments[i].DateOfPayment.ToString("dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    ws.Cells[i + 2, 3].Value = Payments[i].ReciptNo;
                    ws.Cells[i + 2, 4].Value = Payments[i].ChequeNo;
                    ws.Cells[i + 2, 5].Value = Payments[i].Amount;
                }
                package.Save();
            }
        }

        public static string SanitizeValue(ExcelRange cellAddress)
        {
            return (cellAddress.Value ?? (object)"").ToString().Trim();
        }

        public static void ReadFiles()
        {
            List<FileInfo> Files = Directory.EnumerateFiles("Data", "*.xlsx", SearchOption.AllDirectories).Select(m => new FileInfo(m)).ToList();
            ExcelPackage package;
            foreach (FileInfo file in Files)
            {
                using(package = new ExcelPackage(file))
                {
                    Console.WriteLine("Processing file " + Path.GetFileName(file.FullName));
                    if (package.Workbook.Worksheets.Count == 0) {
                        Console.WriteLine("No Sheets.. Exiting.. ");
                        return;
                    };
                    
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                    ExcelRange range = worksheet.SelectedRange[worksheet.Dimension.Start.Address + ":" + worksheet.Dimension.End.Address];
                    int StartRow = 0;
                    int EndRow = 0;
                    bool isCash = false;
                    for (int i = 2; i < range.Rows; i++)
                    {
                        string[] row_ident = { "SL NO.", "USER ID", "COUNTER NO", "DIVISION NAME", "GROUP BOOK NO." };
                       
                        StartRow =     SanitizeValue(worksheet.Cells[i, 1]).Equals(row_ident[0]) &&
                                             SanitizeValue(worksheet.Cells[i, 2]).Equals(row_ident[1]) &&
                                             SanitizeValue(worksheet.Cells[i, 3]).Equals(row_ident[2]) &&
                                             SanitizeValue(worksheet.Cells[i, 4]).Equals(row_ident[3]) ? i+3 : StartRow;

                        EndRow = SanitizeValue(worksheet.Cells[i, 1]).Equals("SUB TOTAL :") ? i : EndRow;

                        if(StartRow != 0 && i >= StartRow && EndRow == 0)
                        {
                            
                            isCash = (worksheet.Cells[i, 8].Value ?? (object)"").ToString().Trim().Equals("");
                            Payments.Add(new PaymentRecords()
                            {
                                Amount = decimal.Parse(Regex.Replace((SanitizeValue(worksheet.Cells[i, 14]).Equals("") ? "0.0" : SanitizeValue(worksheet.Cells[i, 14])), "[^[^0-9.]", "")),
                                ChequeNo = !isCash ? SanitizeValue(worksheet.Cells[i, 8]) : null,
                                DateOfPayment = SanitizeValue(worksheet.Cells[i, 11]).Equals("") ? new DateTime() : DateTime.Parse(SanitizeValue(worksheet.Cells[i, 11]), System.Globalization.CultureInfo.InvariantCulture),
                                FileName = file,
                                ReciptNo = SanitizeValue(worksheet.Cells[i, 7])
                            });
                        }
                    }
                }
                
                PaymentRecords pm = Payments.FirstOrDefault(m => m.FileName == file && m.ChequeNo == null);
                if (pm != null)
                {
                    decimal SumOfCashTxns = Payments.Where(m => m.FileName == file && m.ChequeNo == null).Select(m => m.Amount).Sum();
                    pm.Amount = SumOfCashTxns;
                    Payments.RemoveAll(m => m.FileName == file && m.ChequeNo == null);
                    Payments.Add(pm);
                }
            }
        }
    }
}
