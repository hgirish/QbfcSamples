using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceQuery
{
   public class InvoiceQueryHelper
    {
        public  IList<Invoice> GetInvoiceWithCustomer()
        {
            bool sessionBegun = false;
            QBSessionManager sessionManager = null;
            var invoices = new List<Invoice>();
            sessionManager = new QBSessionManager();

            try
            {
                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "GenerateInvoicePDFs");
                //connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();
                // all invoices modified in the month of August 2016
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue(new DateTime(2017, 12, 1),true);
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue(new DateTime(2017, 12, 31),true);
                //invoiceQueryRq.IncludeLineItems.SetValue(true);

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
                        var customerListId = invoiceRet.CustomerRef?.ListID?.GetValue();
                        if (customerListId != null)
                        {
                            Console.WriteLine($"INv:{invoice.InvoiceNumber}, Job: {invoice.JobNumber}, Name:{invoice.CustomerName}");
                            requestMsgSet.ClearRequests();
                            ICustomerQuery customerQueryRq = requestMsgSet.AppendCustomerQueryRq();
                            customerQueryRq.ORCustomerListQuery.ListIDList.Add(customerListId);

                            //Send the request and get the response from QuickBooks
                            responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                            response = responseMsgSet.ResponseList.GetAt(0);

                            ICustomerRetList customerRetList = (ICustomerRetList)response.Detail;
                            ICustomerRet customerRet = customerRetList.GetAt(0);

                            invoice.Customer = new Customer
                            {
                                Name = customerRet.Name?.GetValue(),
                                QuickBooksID = customerRet.ListID?.GetValue(),
                                EditSequence = customerRet.EditSequence?.GetValue(),
                                FullName = customerRet.FullName?.GetValue(),
                                CompanyName = customerRet.CompanyName?.GetValue()
                            };
                            Console.WriteLine($"{i}\t{invoice.Customer.Name}\t{invoice.Customer.FullName}\t{invoice.Customer.CompanyName}");

                        }


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
            catch (Exception ex)
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
        public IList<Invoice> GetInvoiceAndJob()
        {
            bool sessionBegun = false;
            QBSessionManager sessionManager = null;
            var invoices = new List<Invoice>();
            sessionManager = new QBSessionManager();

            try
            {
                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "GenerateInvoicePDFs");
                //connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();
                // invoiceQueryRq.IncludeLineItems.SetValue(true);
                // all invoices modified in the month of August 2016
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue(new DateTime(2017, 8, 1), true);
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue(new DateTime(2017, 12, 31), true);

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
            catch (Exception ex)
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
