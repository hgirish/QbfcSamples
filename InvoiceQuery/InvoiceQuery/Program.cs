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
           var invoices = GetInvoiceAndJob();
            Export(invoices);

        }
        public static void Export(IList<Invoice> list)
        {
            using (TextWriter sw = new StreamWriter("D:\\Invoice_JobNumbers.csv"))
            {
                sw.WriteLine("InvoiceNumber\tJobNumber\tCustomerName\tMemo");
                foreach (var item in list)
                {
                    sw.WriteLine($"{item.InvoiceNumber}\t{item.JobNumber}\t{item.CustomerName}\t{item.Memo}");
                }
            }
        }
        public static IList<Invoice> GetInvoiceAndJob()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;
            var invoices = new List<Invoice>();
            sessionManager = new QBSessionManager();

            try
            {
                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "GenerateInvoicePDFs");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();
                invoiceQueryRq.IncludeLineItems.SetValue(true);

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                IInvoiceRetList invoiceRetList = (IInvoiceRetList)response.Detail;

                if (invoiceRetList != null)
                {
                    for (int i = 0; i < invoiceRetList.Count; i++)
                    {
                        IInvoiceRet invoiceRet = invoiceRetList.GetAt(i);

                        var invoice = new Invoice
                        {
                            QuickBooksID = invoiceRet.TxnID.GetValue(),
                            EditSequence = invoiceRet.EditSequence.GetValue(),
                            InvoiceNumber = invoiceRet.RefNumber?.GetValue(),
                            Memo = invoiceRet.Memo?.GetValue(),
                            JobNumber = invoiceRet.Other?.GetValue(),
                            CustomerName = invoiceRet.CustomerRef.FullName?.GetValue()

                        };
                        Console.WriteLine($"INv:{invoice.InvoiceNumber}, Job: {invoice.JobNumber}, Name:{invoice.CustomerName}");
                        requestMsgSet.ClearRequests();
                        invoices.Add(invoice);
                    }
                }
                if (requestMsgSet != null)
                {
                    Marshal.FinalReleaseComObject(requestMsgSet);
                }
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
            }
            catch (Exception ex )
            {

                Console.WriteLine(ex.Message.ToString() + "\nStack Trace: \n" + ex.StackTrace + "\nExiting the application");
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                    sessionManager.CloseConnection();
                }
            }
            return invoices;

        }
        
    }
}
