//The following sample code is generated as an illustration of 
//Creating requests and parsing responses ONLY
//This code is NOT intended to show best practices or ideal code
//Use at your most careful discretion

using System;
using QBFC13Lib;

namespace InvoiceQuery
{
    public class Sample2
    {
        public void DoItemServiceQuery()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemServiceQueryRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemServiceQueryRs(responseMsgSet);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Error");
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }
        }


        void BuildItemServiceQueryRq(IMsgSetRequest requestMsgSet)
        {
            IItemServiceQuery ItemServiceQueryRq = requestMsgSet.AppendItemServiceQueryRq();
            //Set attributes
            //Set field value for metaData
            
            ItemServiceQueryRq.metaData.SetAsString("IQBENmetaDataType");
            //Set field value for iterator
            ItemServiceQueryRq.iterator.SetAsString("IQBENiteratorType");
            //Set field value for iteratorID
            ItemServiceQueryRq.iteratorID.SetValue("IQBUUIDType");
            string ORListQueryWithOwnerIDAndClassElementType42 = "ListIDList";
            if (ORListQueryWithOwnerIDAndClassElementType42 == "ListIDList")
            {
                //Set field value for ListIDList
                //May create more than one of these if needed
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListIDList.Add("200000-1011023419");
            }
            if (ORListQueryWithOwnerIDAndClassElementType42 == "FullNameList")
            {
                //Set field value for FullNameList
                //May create more than one of these if needed
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.FullNameList.Add("ab");
            }
            if (ORListQueryWithOwnerIDAndClassElementType42 == "ListWithClassFilter")
            {
                //Set field value for MaxReturned
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.SetValue(6);
                //Set field value for ActiveStatus
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly);
                //Set field value for FromModifiedDate
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.FromModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"), false);
                //Set field value for ToModifiedDate
                ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ToModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"), false);
                string ORNameFilterElementType43 = "NameFilter";
                if (ORNameFilterElementType43 == "NameFilter")
                {
                    //Set field value for MatchCriterion
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                    //Set field value for Name
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameFilter.Name.SetValue("ab");
                }
                if (ORNameFilterElementType43 == "NameRangeFilter")
                {
                    //Set field value for FromName
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.SetValue("ab");
                    //Set field value for ToName
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.ToName.SetValue("ab");
                }
                string ORClassFilterElementType44 = "ListIDList";
                if (ORClassFilterElementType44 == "ListIDList")
                {
                    //Set field value for ListIDList
                    //May create more than one of these if needed
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ClassFilter.ORClassFilter.ListIDList.Add("200000-1011023419");
                }
                if (ORClassFilterElementType44 == "FullNameList")
                {
                    //Set field value for FullNameList
                    //May create more than one of these if needed
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ClassFilter.ORClassFilter.FullNameList.Add("ab");
                }
                if (ORClassFilterElementType44 == "ListIDWithChildren")
                {
                    //Set field value for ListIDWithChildren
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ClassFilter.ORClassFilter.ListIDWithChildren.SetValue("200000-1011023419");
                }
                if (ORClassFilterElementType44 == "FullNameWithChildren")
                {
                    //Set field value for FullNameWithChildren
                    ItemServiceQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ClassFilter.ORClassFilter.FullNameWithChildren.SetValue("ab");
                }
            }
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            ItemServiceQueryRq.IncludeRetElementList.Add("ab");
            //Set field value for OwnerIDList
            //May create more than one of these if needed
            ItemServiceQueryRq.OwnerIDList.Add(Guid.NewGuid().ToString());
        }




        void WalkItemServiceQueryRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;

            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            //if we sent only one request, there is only one response, we'll walk the list for this sample
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                //check the status code of the response, 0=ok, >0 is warning
                if (response.StatusCode >= 0)
                {
                    //the request-specific response is in the details, make sure we have some
                    if (response.Detail != null)
                    {
                        //make sure the response is the type we're expecting
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtItemServiceQueryRs)
                        {
                            //upcast to more specific type here, this is safe because we checked with response.Type check above
                            IItemServiceRetList ItemServiceRet = (IItemServiceRetList)response.Detail;
                            WalkItemServiceRet(ItemServiceRet.GetAt(0));
                        }
                    }
                }
            }
        }



        void WalkItemServiceRet(IItemServiceRet ItemServiceRet)
        {
            if (ItemServiceRet == null) return;

            //Go through all the elements of IItemServiceRetList
            //Get value of ListID
            string ListID45 = (string)ItemServiceRet.ListID.GetValue();
            //Get value of TimeCreated
            DateTime TimeCreated46 = (DateTime)ItemServiceRet.TimeCreated.GetValue();
            //Get value of TimeModified
            DateTime TimeModified47 = (DateTime)ItemServiceRet.TimeModified.GetValue();
            //Get value of EditSequence
            string EditSequence48 = (string)ItemServiceRet.EditSequence.GetValue();
            //Get value of Name
            string Name49 = (string)ItemServiceRet.Name.GetValue();
            //Get value of FullName
            string FullName50 = (string)ItemServiceRet.FullName.GetValue();
            //Get value of BarCodeValue
            if (ItemServiceRet.BarCodeValue != null)
            {
                string BarCodeValue51 = (string)ItemServiceRet.BarCodeValue.GetValue();
            }
            //Get value of IsActive
            if (ItemServiceRet.IsActive != null)
            {
                bool IsActive52 = (bool)ItemServiceRet.IsActive.GetValue();
            }
            if (ItemServiceRet.ClassRef != null)
            {
                //Get value of ListID
                if (ItemServiceRet.ClassRef.ListID != null)
                {
                    string ListID53 = (string)ItemServiceRet.ClassRef.ListID.GetValue();
                }
                //Get value of FullName
                if (ItemServiceRet.ClassRef.FullName != null)
                {
                    string FullName54 = (string)ItemServiceRet.ClassRef.FullName.GetValue();
                }
            }
            if (ItemServiceRet.ParentRef != null)
            {
                //Get value of ListID
                if (ItemServiceRet.ParentRef.ListID != null)
                {
                    string ListID55 = (string)ItemServiceRet.ParentRef.ListID.GetValue();
                }
                //Get value of FullName
                if (ItemServiceRet.ParentRef.FullName != null)
                {
                    string FullName56 = (string)ItemServiceRet.ParentRef.FullName.GetValue();
                }
            }
            //Get value of Sublevel
            int Sublevel57 = (int)ItemServiceRet.Sublevel.GetValue();
            if (ItemServiceRet.UnitOfMeasureSetRef != null)
            {
                //Get value of ListID
                if (ItemServiceRet.UnitOfMeasureSetRef.ListID != null)
                {
                    string ListID58 = (string)ItemServiceRet.UnitOfMeasureSetRef.ListID.GetValue();
                }
                //Get value of FullName
                if (ItemServiceRet.UnitOfMeasureSetRef.FullName != null)
                {
                    string FullName59 = (string)ItemServiceRet.UnitOfMeasureSetRef.FullName.GetValue();
                }
            }
            if (ItemServiceRet.SalesTaxCodeRef != null)
            {
                //Get value of ListID
                if (ItemServiceRet.SalesTaxCodeRef.ListID != null)
                {
                    string ListID60 = (string)ItemServiceRet.SalesTaxCodeRef.ListID.GetValue();
                }
                //Get value of FullName
                if (ItemServiceRet.SalesTaxCodeRef.FullName != null)
                {
                    string FullName61 = (string)ItemServiceRet.SalesTaxCodeRef.FullName.GetValue();
                }
            }
            if (ItemServiceRet.ORSalesPurchase != null)
            {
                if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase != null)
                {
                    if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase != null)
                    {
                        //Get value of Desc
                        if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc != null)
                        {
                            string Desc62 = (string)ItemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue();
                        }
                        if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice != null)
                        {
                            if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price != null)
                            {
                                //Get value of Price
                                if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price != null)
                                {
                                    double Price63 = (double)ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue();
                                }
                            }
                            if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent != null)
                            {
                                //Get value of PricePercent
                                if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent != null)
                                {
                                    double PricePercent64 = (double)ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.PricePercent.GetValue();
                                }
                            }
                        }
                        if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef != null)
                        {
                            //Get value of ListID
                            if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID != null)
                            {
                                string ListID65 = (string)ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName != null)
                            {
                                string FullName66 = (string)ItemServiceRet.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.GetValue();
                            }
                        }
                    }
                }
                if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase != null)
                {
                    if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase != null)
                    {
                        //Get value of SalesDesc
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.SalesDesc != null)
                        {
                            string SalesDesc67 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.SalesDesc.GetValue();
                        }
                        //Get value of SalesPrice
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.SalesPrice != null)
                        {
                            double SalesPrice68 = (double)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.SalesPrice.GetValue();
                        }
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.IncomeAccountRef != null)
                        {
                            //Get value of ListID
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.IncomeAccountRef.ListID != null)
                            {
                                string ListID69 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.IncomeAccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.IncomeAccountRef.FullName != null)
                            {
                                string FullName70 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.IncomeAccountRef.FullName.GetValue();
                            }
                        }
                        //Get value of PurchaseDesc
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PurchaseDesc != null)
                        {
                            string PurchaseDesc71 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PurchaseDesc.GetValue();
                        }
                        //Get value of PurchaseCost
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PurchaseCost != null)
                        {
                            double PurchaseCost72 = (double)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PurchaseCost.GetValue();
                        }
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.ExpenseAccountRef != null)
                        {
                            //Get value of ListID
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.ExpenseAccountRef.ListID != null)
                            {
                                string ListID73 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.ExpenseAccountRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.ExpenseAccountRef.FullName != null)
                            {
                                string FullName74 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.ExpenseAccountRef.FullName.GetValue();
                            }
                        }
                        if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PrefVendorRef != null)
                        {
                            //Get value of ListID
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PrefVendorRef.ListID != null)
                            {
                                string ListID75 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PrefVendorRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PrefVendorRef.FullName != null)
                            {
                                string FullName76 = (string)ItemServiceRet.ORSalesPurchase.SalesAndPurchase.PrefVendorRef.FullName.GetValue();
                            }
                        }
                    }
                }
            }
            //Get value of ExternalGUID
            if (ItemServiceRet.ExternalGUID != null)
            {
                string ExternalGUID77 = (string)ItemServiceRet.ExternalGUID.GetValue();
            }
            if (ItemServiceRet.DataExtRetList != null)
            {
                for (int i78 = 0; i78 < ItemServiceRet.DataExtRetList.Count; i78++)
                {
                    IDataExtRet DataExtRet = ItemServiceRet.DataExtRetList.GetAt(i78);
                    //Get value of OwnerID
                    if (DataExtRet.OwnerID != null)
                    {
                        string OwnerID79 = (string)DataExtRet.OwnerID.GetValue();
                    }
                    //Get value of DataExtName
                    string DataExtName80 = (string)DataExtRet.DataExtName.GetValue();
                    //Get value of DataExtType
                    ENDataExtType DataExtType81 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                    //Get value of DataExtValue
                    string DataExtValue82 = (string)DataExtRet.DataExtValue.GetValue();
                }
            }
        }



    }
}