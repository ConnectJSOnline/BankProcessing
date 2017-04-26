using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankProcessing
{
    class PaymentRecords
    {
        public DateTime DateOfPayment { get; set; }
        public string ReciptNo { get; set; }
        public string ChequeNo { get; set; }
        public decimal Amount { get; set; }
        public FileInfo FileName { get; set; }
    }
}
