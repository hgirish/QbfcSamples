using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceQuery
{
  public  class InvoiceQueryHelperDetail
    {
        public IList<Invoice> GetInvoiceDetail(DateTime fromDate, DateTime toDate)
        {
            bool sessionBegun = false;
            QBSessionManager sessionManager = null;
            var invoices = new List<Invoice>();
            sessionManager = new QBSessionManager();
            IMsgSetRequest requestMsgSet = null;
            

            try
            {
                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "GenerateInvoiceSummary");
                //connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Create the message set request object to hold our request
                requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IInvoiceQuery invoiceQueryRq = requestMsgSet.AppendInvoiceQueryRq();
                // all invoices modified in the month of August 2016
                // get all invoices for the month of august 2016
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(fromDate);
                invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(toDate);

               // invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.FromModifiedDate.SetValue(new DateTime(2017, 12, 1), true);
               // invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.ModifiedDateRangeFilter.ToModifiedDate.SetValue(new DateTime(2017, 12, 31), true);
                invoiceQueryRq.IncludeLineItems.SetValue(true);

               
                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                IInvoiceRetList invoiceRetList = (IInvoiceRetList)response.Detail;
                Console.WriteLine($"Invoices found: {invoiceRetList.Count}");
               
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
                            InvoiceDate = invoiceRet.TimeCreated?.GetValue(),
                            Memo = invoiceRet.Memo?.GetValue(),
                            JobNumber = invoiceRet.Other?.GetValue(),
                            CustomerName = invoiceRet.CustomerRef.FullName?.GetValue(),
                            Amount = invoiceRet.BalanceRemaining?.GetValue()

                        };
                        if (invoice.Amount == 0)
                        {
                            invoice.Amount = invoiceRet.Subtotal?.GetValue();
                        }
                        var customerListId = invoiceRet.CustomerRef?.ListID?.GetValue();
                        if (customerListId != null)
                        {
                          //  Console.WriteLine($"{i}\tInv:{invoice.InvoiceNumber}, Job: {invoice.JobNumber}, Name:{invoice.CustomerName}");
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
                                CompanyName = customerRet.CompanyName?.GetValue(),
                            };
                          //  Console.WriteLine($"{i}\t{invoice.Customer.Name}\t{invoice.Customer.FullName}\t{invoice.Customer.CompanyName}");
                           // Console.WriteLine($"{i}\t\tInvoice detail starts");
                            if (invoiceRet.ORInvoiceLineRetList != null)
                            {
                                invoice.Description = "";

                                if (invoiceRet.ORInvoiceLineRetList.Count > 0)
                                {
                                    try
                                    {
                                        IORInvoiceLineRet orInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(1);
                                    
                                        invoice.Description = orInvoiceLineRet?.InvoiceLineRet?.Desc?.GetValue();
                                        if (string.IsNullOrEmpty(invoice.Description))
                                        {
                                             orInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(0);

                                            invoice.Description = orInvoiceLineRet?.InvoiceLineRet?.Desc?.GetValue();
                                        }
                                    }
                                    catch (Exception ex )
                                    {

                                        Console.WriteLine(ex.Message);
                                        try
                                        {
                                            IORInvoiceLineRet orInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(0);

                                            invoice.Description = orInvoiceLineRet?.InvoiceLineRet?.Desc?.GetValue();
                                        }
                                        catch (Exception ex2)
                                        {

                                            Console.WriteLine(ex2.Message);
                                        }
                                    }
                                    
                                    Console.WriteLine($"{invoice.InvoiceNumber}\t{invoice.Amount} \t{invoice.JobNumber}\t{invoice.Description}");

                                }
                               // Console.WriteLine($"InvoiceList Count: {invoiceRet.ORInvoiceLineRetList.Count}");
                                /*
                                for (int j = 0; j < invoiceRet.ORInvoiceLineRetList.Count; j++)
                                {
                                    IORInvoiceLineRet orInvoiceLineRet = invoiceRet.ORInvoiceLineRetList.GetAt(j);


                                    if (orInvoiceLineRet != null && orInvoiceLineRet.InvoiceLineRet != null)
                                    {
                                        var invoiceItem = new InvoiceItem
                                        {
                                            Amount = orInvoiceLineRet.InvoiceLineRet.Amount?.GetValue(),
                                            QuickBooksID = orInvoiceLineRet.InvoiceLineRet.TxnLineID?.GetValue()
                                        };

                                        requestMsgSet.ClearRequests();
                                        IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
                                        itemQueryRq.ORListQuery.ListIDList.Add(orInvoiceLineRet.InvoiceLineRet.ItemRef?.ListID?.GetValue());
                                        
                                        //Send the request and get the response from QuickBooks
                                        responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                                        response = responseMsgSet.ResponseList.GetAt(0);
                                        IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                        // Console.WriteLine($"ItemRetList.Count: {itemRetList.Count}");
                                        
                                        IORItemRet itemRet = itemRetList.GetAt(0);
                                        var ortype = itemRet.ortype;
                                        if (itemRet.ItemInventoryRet != null)
                                        {
                                            IItemInventoryRet itemInventoryRet = itemRet.ItemInventoryRet;

                                            var item = new Item
                                            {
                                                Name = itemInventoryRet.Name?.GetValue(),
                                                Description = itemInventoryRet.SalesDesc?.GetValue(),
                                                Rate = itemInventoryRet.SalesPrice?.GetValue(),
                                                ItemType = ortype.ToString(),
                                                QuickBooksID = itemInventoryRet.ListID?.GetValue(),
                                                EditSequence = itemInventoryRet.EditSequence?.GetValue()
                                                 
                                            };
                                            if (string.IsNullOrEmpty(item.Name))
                                            {
                                                item.Name = itemInventoryRet.FullName?.GetValue();
                                            }

                                            invoiceItem.Item = item;
                                        }
                                        else if (itemRet.ItemServiceRet != null)
                                        {
                                            IItemServiceRet itemServiceRet = itemRet.ItemServiceRet;

                                            var item = new Item
                                            {
                                                Name = itemServiceRet.Name.GetValue(),
                                                Description = itemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc?.GetValue(),
                                                Rate = itemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price?.GetValue(),
                                                //ItemType = ItemType.Service,
                                                ItemType = ortype.ToString(),
                                                QuickBooksID = itemServiceRet.ListID?.GetValue(),
                                                EditSequence = itemServiceRet.EditSequence?.GetValue(),
                                               // FullName = itemServiceRet.ToString()
                                                };
                                            if (string.IsNullOrEmpty(item.Name))
                                            {
                                                item.Name = itemServiceRet.FullName?.GetValue();
                                            }
                                            invoiceItem.Item = item;
                                        }
                                        else if (itemRet.ItemOtherChargeRet != null)
                                        {
                                            IItemOtherChargeRet itemOtherChargeRet = itemRet.ItemOtherChargeRet;
                                            var item = new Item
                                            {
                                                Name = itemOtherChargeRet.Name?.GetValue(),
                                                Description = itemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.Desc?.GetValue(),
                                                Rate = itemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price?.GetValue(),
                                                ItemType = ortype.ToString()

                                            };
                                             if (string.IsNullOrEmpty(item.Name))
                                            {
                                                item.Name = itemOtherChargeRet.FullName?.GetValue();
                                            }
                                            invoiceItem.Item = item;
                                        }
                                        else if (itemRet.ItemNonInventoryRet != null)
                                        {
                                            IItemNonInventoryRet itemNonInventoryRet = itemRet.ItemNonInventoryRet;
                                            var item = new Item
                                            {
                                                Name = itemNonInventoryRet.Name?.GetValue(),
                                                Description = itemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc?.GetValue(),
                                                ItemType = ortype.ToString()

                                            };
                                            if (string.IsNullOrEmpty(item.Name))
                                            {
                                                item.Name = itemNonInventoryRet.FullName?.GetValue();
                                            }
                                            invoiceItem.Item = item;
                                        }
                                        Console.WriteLine($"{invoiceItem.Item.FullName}\t{invoice.InvoiceNumber}\t{invoiceItem.Amount}\t{invoiceItem.Item.Description}");
                                        invoice.InvoiceItems.Add(invoiceItem);

                                    }
                                }
                                */
                            }

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
                if (requestMsgSet != null)
                {
                    Marshal.FinalReleaseComObject(requestMsgSet);
                }
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
