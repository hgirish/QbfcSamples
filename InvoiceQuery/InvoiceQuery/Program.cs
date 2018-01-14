using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceQuery
{
    class Program
    {

        static void Main(string[] args)
        {
            // InvoiceQueryHelper helper = new InvoiceQueryHelper();
             var helper = new InvoiceQueryHelperDetail();
            // var invoices = helper.GetInvoiceAndJob();
            //  var invoices = helper.GetInvoiceWithCustomer();
             var invoices = helper.GetInvoiceDetail();
            //var demo = new Demo();
            //demo.GetInvoiceDetail();
            

            Export(invoices);

        }
        public static void Export(IList<Invoice> list)
        {
            using (TextWriter sw = new StreamWriter("D:\\Invoice_JobNumbers.csv"))
            {
                sw.WriteLine("InvoiceNumber\tJobNumber\tCustomerName\tMemo\tCustomer_Name\tDescription");
                foreach (var item in list)
                {
                    sw.WriteLine($"{item.InvoiceNumber}\t{item.JobNumber}\t{item.CustomerName}\t{item.Memo}\t{item.Customer?.Name}\t{item.Description}");
                }
            }
        }
       
        
    }
}
