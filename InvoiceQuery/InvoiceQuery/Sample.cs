using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceQuery
{
    class Sample
    {
        public static void DoQbXml()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            sessionManager = new QBSessionManager();

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

            var invoices = new List<Invoice>();

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
                        JobNumber = invoiceRet.Other?.GetValue()

                    };
                    Console.WriteLine($"INv:{invoice.InvoiceNumber}, EditSequence: {invoice.JobNumber}");
                    requestMsgSet.ClearRequests();
                    ICustomerQuery customerQueryRq = requestMsgSet.AppendCustomerQueryRq();
                    customerQueryRq.ORCustomerListQuery.ListIDList.Add(invoiceRet.CustomerRef.ListID.GetValue());

                    //Send the request and get the response from QuickBooks
                    responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                    response = responseMsgSet.ResponseList.GetAt(0);
                    ICustomerRetList customerRetList = (ICustomerRetList)response.Detail;
                    ICustomerRet customerRet = customerRetList.GetAt(0);
                    if (i > 200)
                    {
                        return;
                    }
                    invoice.Customer = new Customer
                    {
                        Name = customerRet.Name.GetValue(),
                        QuickBooksID = customerRet.ListID.GetValue(),
                        EditSequence = customerRet.EditSequence.GetValue()
                    };
                    if (invoiceRet.ORInvoiceLineRetList != null)
                    {
                        for (int j = 0; j < invoiceRet.ORInvoiceLineRetList.Count; j++)
                        {
                            IORInvoiceLineRet ORInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(j);

                            try
                            {
                                var invoiceItem = new InvoiceItem
                                {
                                    Amount = ORInvoiceLineRet.InvoiceLineRet.Amount.GetValue(),
                                    QuickBooksID = ORInvoiceLineRet.InvoiceLineRet.TxnLineID.GetValue()
                                };


                                requestMsgSet.ClearRequests();
                                IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
                                itemQueryRq.ORListQuery.ListIDList.Add(ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID.GetValue());

                                //Send the request and get the response from QuickBooks
                                responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                                response = responseMsgSet.ResponseList.GetAt(0);
                                IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                IORItemRet itemRet = itemRetList.GetAt(0);

                                if (itemRet.ItemInventoryRet != null)
                                {
                                    IItemInventoryRet itemInventoryRet = itemRet.ItemInventoryRet;

                                    var item = new Item
                                    {
                                        Name = itemInventoryRet.Name.GetValue(),
                                        Description = itemInventoryRet.SalesDesc.GetValue(),
                                        Rate = itemInventoryRet.SalesPrice.GetValue(),
                                        //ItemType = ItemType.Inventory,
                                        QuickBooksID = itemInventoryRet.ListID.GetValue(),
                                        EditSequence = itemInventoryRet.EditSequence.GetValue()
                                    };

                                    invoiceItem.Item = item;
                                }
                                else if (itemRet.ItemServiceRet != null)
                                {
                                    IItemServiceRet itemServiceRet = itemRet.ItemServiceRet;

                                    var item = new Item
                                    {
                                        Name = itemServiceRet.Name.GetValue(),
                                        Description = itemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue(),
                                        Rate = itemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue(),
                                        //ItemType = ItemType.Service,
                                        QuickBooksID = itemServiceRet.ListID.GetValue(),
                                        EditSequence = itemServiceRet.EditSequence.GetValue()
                                    };

                                    invoiceItem.Item = item;
                                }

                                invoice.InvoiceItems.Add(invoiceItem);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(invoice.Customer.Name);
                                throw;
                            }
                        }
                    }

                    invoices.Add(invoice);
                }
            }
        }
        public static void Do()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "GenerateInvoicePDFs");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;



                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;




                string[] InvoiceList = { "30085", "30089" };
                foreach (string invNum in InvoiceList)
                {
                    IInvoiceQuery invQuery = requestMsgSet.AppendInvoiceQueryRq();
                    invQuery.ORInvoiceQuery.RefNumberList.Add(invNum);
                    Console.WriteLine($"invNum: {invNum}");
                }

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                // Check each response
                for (int index = 0; index < responseMsgSet.ResponseList.Count; index++)
                {
                    IResponse response = responseMsgSet.ResponseList.GetAt(index);
                    if (response.StatusCode != 0)
                    {
                        // Save the invoice number in the "not found" file
                        // and display the error
                        Console.WriteLine(". Error: " + response.StatusMessage);
                    }
                    else
                    {
                        Console.WriteLine(response.StatusCode);
                        var detail = (IInvoiceQuery)response.Detail;
                        Console.WriteLine(detail);
                    }
                }


                if (requestMsgSet != null)
                {
                    Marshal.FinalReleaseComObject(requestMsgSet);
                }






            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            sessionManager.EndSession();
            sessionManager.CloseConnection();

        }
    }
}
