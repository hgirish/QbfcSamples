using QBFC13Lib;
using System;

//The following sample code is generated as an illustration of 
//Creating requests and parsing responses ONLY
//This code is NOT intended to show best practices or ideal code
//Use at your most careful discretion


namespace com.intuit.idn.samples
{
    public class Sample
    {
        public void DoInvoiceAdd()
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

                BuildInvoiceAddRq(requestMsgSet);

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

                WalkInvoiceAddRs(responseMsgSet);

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


        void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet)
        {
            IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
            //Set attributes
            //Set field value for defMacro
            InvoiceAddRq.defMacro.SetValue("IQBStringType");
            //Set field value for ListID
            InvoiceAddRq.CustomerRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.CustomerRef.FullName.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.ClassRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.ClassRef.FullName.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.ARAccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.ARAccountRef.FullName.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.TemplateRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.TemplateRef.FullName.SetValue("ab");
            //Set field value for TxnDate
            InvoiceAddRq.TxnDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for RefNumber
            InvoiceAddRq.RefNumber.SetValue("ab");
            //Set field value for Addr1
            InvoiceAddRq.BillAddress.Addr1.SetValue("ab");
            //Set field value for Addr2
            InvoiceAddRq.BillAddress.Addr2.SetValue("ab");
            //Set field value for Addr3
            InvoiceAddRq.BillAddress.Addr3.SetValue("ab");
            //Set field value for Addr4
            InvoiceAddRq.BillAddress.Addr4.SetValue("ab");
            //Set field value for Addr5
            InvoiceAddRq.BillAddress.Addr5.SetValue("ab");
            //Set field value for City
            InvoiceAddRq.BillAddress.City.SetValue("ab");
            //Set field value for State
            InvoiceAddRq.BillAddress.State.SetValue("ab");
            //Set field value for PostalCode
            InvoiceAddRq.BillAddress.PostalCode.SetValue("ab");
            //Set field value for Country
            InvoiceAddRq.BillAddress.Country.SetValue("ab");
            //Set field value for Note
            InvoiceAddRq.BillAddress.Note.SetValue("ab");
            //Set field value for Addr1
            InvoiceAddRq.ShipAddress.Addr1.SetValue("ab");
            //Set field value for Addr2
            InvoiceAddRq.ShipAddress.Addr2.SetValue("ab");
            //Set field value for Addr3
            InvoiceAddRq.ShipAddress.Addr3.SetValue("ab");
            //Set field value for Addr4
            InvoiceAddRq.ShipAddress.Addr4.SetValue("ab");
            //Set field value for Addr5
            InvoiceAddRq.ShipAddress.Addr5.SetValue("ab");
            //Set field value for City
            InvoiceAddRq.ShipAddress.City.SetValue("ab");
            //Set field value for State
            InvoiceAddRq.ShipAddress.State.SetValue("ab");
            //Set field value for PostalCode
            InvoiceAddRq.ShipAddress.PostalCode.SetValue("ab");
            //Set field value for Country
            InvoiceAddRq.ShipAddress.Country.SetValue("ab");
            //Set field value for Note
            InvoiceAddRq.ShipAddress.Note.SetValue("ab");
            //Set field value for IsPending
            InvoiceAddRq.IsPending.SetValue(true);
            //Set field value for IsFinanceCharge
            InvoiceAddRq.IsFinanceCharge.SetValue(true);
            //Set field value for PONumber
            InvoiceAddRq.PONumber.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.TermsRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.TermsRef.FullName.SetValue("ab");
            //Set field value for DueDate
            InvoiceAddRq.DueDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for ListID
            InvoiceAddRq.SalesRepRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.SalesRepRef.FullName.SetValue("ab");
            //Set field value for FOB
            InvoiceAddRq.FOB.SetValue("ab");
            //Set field value for ShipDate
            InvoiceAddRq.ShipDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for ListID
            InvoiceAddRq.ShipMethodRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.ShipMethodRef.FullName.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.ItemSalesTaxRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.ItemSalesTaxRef.FullName.SetValue("ab");
            //Set field value for Memo
            InvoiceAddRq.Memo.SetValue("ab");
            //Set field value for ListID
            InvoiceAddRq.CustomerMsgRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.CustomerMsgRef.FullName.SetValue("ab");
            //Set field value for IsToBePrinted
            InvoiceAddRq.IsToBePrinted.SetValue(true);
            //Set field value for IsToBeEmailed
            InvoiceAddRq.IsToBeEmailed.SetValue(true);
            //Set field value for ListID
            InvoiceAddRq.CustomerSalesTaxCodeRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            InvoiceAddRq.CustomerSalesTaxCodeRef.FullName.SetValue("ab");
            //Set field value for Other
            InvoiceAddRq.Other.SetValue("ab");
            //Set field value for ExchangeRate
           // InvoiceAddRq.ExchangeRate.SetValue("IQBFloatType");
            //Set field value for ExternalGUID
            InvoiceAddRq.ExternalGUID.SetValue(Guid.NewGuid().ToString());
            //Set field value for LinkToTxnIDList
            //May create more than one of these if needed
            InvoiceAddRq.LinkToTxnIDList.Add("200000-1011023419");
            ISetCredit SetCredit1 = InvoiceAddRq.SetCreditList.Append();
            //Set field value for CreditTxnID
            SetCredit1.CreditTxnID.SetValue("200000-1011023419");
            //Set attributes
            //Set field value for useMacro
          ////  SetCredit1.useMacro.SetValue("IQBStringType");
            //Set field value for AppliedAmount
            SetCredit1.AppliedAmount.SetValue(10.01);
            //Set field value for Override
            SetCredit1.Override.SetValue(true);
            IORInvoiceLineAdd ORInvoiceLineAddListElement2 = InvoiceAddRq.ORInvoiceLineAddList.Append();
            string ORInvoiceLineAddListElementType3 = "InvoiceLineAdd";
            if (ORInvoiceLineAddListElementType3 == "InvoiceLineAdd")
            {
                //Set attributes
                //Set field value for defMacro
                ORInvoiceLineAddListElement2.InvoiceLineAdd.defMacro.SetValue("IQBStringType");
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.ItemRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.ItemRef.FullName.SetValue("ab");
                //Set field value for Desc
                ORInvoiceLineAddListElement2.InvoiceLineAdd.Desc.SetValue("ab");
                //Set field value for Quantity
                ORInvoiceLineAddListElement2.InvoiceLineAdd.Quantity.SetValue(2);
                //Set field value for UnitOfMeasure
                ORInvoiceLineAddListElement2.InvoiceLineAdd.UnitOfMeasure.SetValue("ab");
                string ORRatePriceLevelElementType4 = "Rate";
                if (ORRatePriceLevelElementType4 == "Rate")
                {
                    //Set field value for Rate
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(15.65);
                }
                if (ORRatePriceLevelElementType4 == "RatePercent")
                {
                    //Set field value for RatePercent
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORRatePriceLevel.RatePercent.SetValue(20.00);
                }
                if (ORRatePriceLevelElementType4 == "PriceLevelRef")
                {
                    //Set field value for ListID
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORRatePriceLevel.PriceLevelRef.ListID.SetValue("200000-1011023419");
                    //Set field value for FullName
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORRatePriceLevel.PriceLevelRef.FullName.SetValue("ab");
                }
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.ClassRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.ClassRef.FullName.SetValue("ab");
                //Set field value for Amount
                ORInvoiceLineAddListElement2.InvoiceLineAdd.Amount.SetValue(10.01);
                //Set field value for OptionForPriceRuleConflict
                ORInvoiceLineAddListElement2.InvoiceLineAdd.OptionForPriceRuleConflict.SetValue(ENOptionForPriceRuleConflict.ofprcZero);
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.InventorySiteRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.InventorySiteRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.InventorySiteLocationRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.InventorySiteLocationRef.FullName.SetValue("ab");
                string ORSerialLotNumberElementType5 = "SerialNumber";
                if (ORSerialLotNumberElementType5 == "SerialNumber")
                {
                    //Set field value for SerialNumber
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORSerialLotNumber.SerialNumber.SetValue("ab");
                }
                if (ORSerialLotNumberElementType5 == "LotNumber")
                {
                    //Set field value for LotNumber
                    ORInvoiceLineAddListElement2.InvoiceLineAdd.ORSerialLotNumber.LotNumber.SetValue("ab");
                }
                //Set field value for ServiceDate
                ORInvoiceLineAddListElement2.InvoiceLineAdd.ServiceDate.SetValue(DateTime.Parse("12/15/2007"));
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.SalesTaxCodeRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.OverrideItemAccountRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineAdd.OverrideItemAccountRef.FullName.SetValue("ab");
                //Set field value for Other1
                ORInvoiceLineAddListElement2.InvoiceLineAdd.Other1.SetValue("ab");
                //Set field value for Other2
                ORInvoiceLineAddListElement2.InvoiceLineAdd.Other2.SetValue("ab");
                //Set field value for TxnID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.LinkToTxn.TxnID.SetValue("200000-1011023419");
                //Set field value for TxnLineID
                ORInvoiceLineAddListElement2.InvoiceLineAdd.LinkToTxn.TxnLineID.SetValue("200000-1011023419");
                IDataExt DataExt6 = ORInvoiceLineAddListElement2.InvoiceLineAdd.DataExtList.Append();
                //Set field value for OwnerID
                DataExt6.OwnerID.SetValue(Guid.NewGuid().ToString());
                //Set field value for DataExtName
                DataExt6.DataExtName.SetValue("ab");
                //Set field value for DataExtValue
                DataExt6.DataExtValue.SetValue("ab");
            }
            if (ORInvoiceLineAddListElementType3 == "InvoiceLineGroupAdd")
            {
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.ItemGroupRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.ItemGroupRef.FullName.SetValue("ab");
                //Set field value for Quantity
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.Quantity.SetValue(2);
                //Set field value for UnitOfMeasure
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.UnitOfMeasure.SetValue("ab");
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.InventorySiteRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.InventorySiteRef.FullName.SetValue("ab");
                //Set field value for ListID
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.InventorySiteLocationRef.ListID.SetValue("200000-1011023419");
                //Set field value for FullName
                ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.InventorySiteLocationRef.FullName.SetValue("ab");
                IDataExt DataExt7 = ORInvoiceLineAddListElement2.InvoiceLineGroupAdd.DataExtList.Append();
                //Set field value for OwnerID
                DataExt7.OwnerID.SetValue(Guid.NewGuid().ToString());
                //Set field value for DataExtName
                DataExt7.DataExtName.SetValue("ab");
                //Set field value for DataExtValue
                DataExt7.DataExtValue.SetValue("ab");
            }
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            InvoiceAddRq.IncludeRetElementList.Add("ab");
        }




        void WalkInvoiceAddRs(IMsgSetResponse responseMsgSet)
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
                        if (responseType == ENResponseType.rtInvoiceAddRs)
                        {
                            //upcast to more specific type here, this is safe because we checked with response.Type check above
                            IInvoiceRet InvoiceRet = (IInvoiceRet)response.Detail;
                            WalkInvoiceRet(InvoiceRet);
                        }
                    }
                }
            }
        }



        void WalkInvoiceRet(IInvoiceRet InvoiceRet)
        {
            if (InvoiceRet == null) return;

            //Go through all the elements of IInvoiceRet
            //Get value of TxnID
            string TxnID8 = (string)InvoiceRet.TxnID.GetValue();
            //Get value of TimeCreated
            DateTime TimeCreated9 = (DateTime)InvoiceRet.TimeCreated.GetValue();
            //Get value of TimeModified
            DateTime TimeModified10 = (DateTime)InvoiceRet.TimeModified.GetValue();
            //Get value of EditSequence
            string EditSequence11 = (string)InvoiceRet.EditSequence.GetValue();
            //Get value of TxnNumber
            if (InvoiceRet.TxnNumber != null)
            {
                int TxnNumber12 = (int)InvoiceRet.TxnNumber.GetValue();
            }
            //Get value of ListID
            if (InvoiceRet.CustomerRef.ListID != null)
            {
                string ListID13 = (string)InvoiceRet.CustomerRef.ListID.GetValue();
            }
            //Get value of FullName
            if (InvoiceRet.CustomerRef.FullName != null)
            {
                string FullName14 = (string)InvoiceRet.CustomerRef.FullName.GetValue();
            }
            if (InvoiceRet.ClassRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.ClassRef.ListID != null)
                {
                    string ListID15 = (string)InvoiceRet.ClassRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.ClassRef.FullName != null)
                {
                    string FullName16 = (string)InvoiceRet.ClassRef.FullName.GetValue();
                }
            }
            if (InvoiceRet.ARAccountRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.ARAccountRef.ListID != null)
                {
                    string ListID17 = (string)InvoiceRet.ARAccountRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.ARAccountRef.FullName != null)
                {
                    string FullName18 = (string)InvoiceRet.ARAccountRef.FullName.GetValue();
                }
            }
            if (InvoiceRet.TemplateRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.TemplateRef.ListID != null)
                {
                    string ListID19 = (string)InvoiceRet.TemplateRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.TemplateRef.FullName != null)
                {
                    string FullName20 = (string)InvoiceRet.TemplateRef.FullName.GetValue();
                }
            }
            //Get value of TxnDate
            DateTime TxnDate21 = (DateTime)InvoiceRet.TxnDate.GetValue();
            //Get value of RefNumber
            if (InvoiceRet.RefNumber != null)
            {
                string RefNumber22 = (string)InvoiceRet.RefNumber.GetValue();
            }
            if (InvoiceRet.BillAddress != null)
            {
                //Get value of Addr1
                if (InvoiceRet.BillAddress.Addr1 != null)
                {
                    string Addr123 = (string)InvoiceRet.BillAddress.Addr1.GetValue();
                }
                //Get value of Addr2
                if (InvoiceRet.BillAddress.Addr2 != null)
                {
                    string Addr224 = (string)InvoiceRet.BillAddress.Addr2.GetValue();
                }
                //Get value of Addr3
                if (InvoiceRet.BillAddress.Addr3 != null)
                {
                    string Addr325 = (string)InvoiceRet.BillAddress.Addr3.GetValue();
                }
                //Get value of Addr4
                if (InvoiceRet.BillAddress.Addr4 != null)
                {
                    string Addr426 = (string)InvoiceRet.BillAddress.Addr4.GetValue();
                }
                //Get value of Addr5
                if (InvoiceRet.BillAddress.Addr5 != null)
                {
                    string Addr527 = (string)InvoiceRet.BillAddress.Addr5.GetValue();
                }
                //Get value of City
                if (InvoiceRet.BillAddress.City != null)
                {
                    string City28 = (string)InvoiceRet.BillAddress.City.GetValue();
                }
                //Get value of State
                if (InvoiceRet.BillAddress.State != null)
                {
                    string State29 = (string)InvoiceRet.BillAddress.State.GetValue();
                }
                //Get value of PostalCode
                if (InvoiceRet.BillAddress.PostalCode != null)
                {
                    string PostalCode30 = (string)InvoiceRet.BillAddress.PostalCode.GetValue();
                }
                //Get value of Country
                if (InvoiceRet.BillAddress.Country != null)
                {
                    string Country31 = (string)InvoiceRet.BillAddress.Country.GetValue();
                }
                //Get value of Note
                if (InvoiceRet.BillAddress.Note != null)
                {
                    string Note32 = (string)InvoiceRet.BillAddress.Note.GetValue();
                }
            }
            if (InvoiceRet.BillAddressBlock != null)
            {
                //Get value of Addr1
                if (InvoiceRet.BillAddressBlock.Addr1 != null)
                {
                    string Addr133 = (string)InvoiceRet.BillAddressBlock.Addr1.GetValue();
                }
                //Get value of Addr2
                if (InvoiceRet.BillAddressBlock.Addr2 != null)
                {
                    string Addr234 = (string)InvoiceRet.BillAddressBlock.Addr2.GetValue();
                }
                //Get value of Addr3
                if (InvoiceRet.BillAddressBlock.Addr3 != null)
                {
                    string Addr335 = (string)InvoiceRet.BillAddressBlock.Addr3.GetValue();
                }
                //Get value of Addr4
                if (InvoiceRet.BillAddressBlock.Addr4 != null)
                {
                    string Addr436 = (string)InvoiceRet.BillAddressBlock.Addr4.GetValue();
                }
                //Get value of Addr5
                if (InvoiceRet.BillAddressBlock.Addr5 != null)
                {
                    string Addr537 = (string)InvoiceRet.BillAddressBlock.Addr5.GetValue();
                }
            }
            if (InvoiceRet.ShipAddress != null)
            {
                //Get value of Addr1
                if (InvoiceRet.ShipAddress.Addr1 != null)
                {
                    string Addr138 = (string)InvoiceRet.ShipAddress.Addr1.GetValue();
                }
                //Get value of Addr2
                if (InvoiceRet.ShipAddress.Addr2 != null)
                {
                    string Addr239 = (string)InvoiceRet.ShipAddress.Addr2.GetValue();
                }
                //Get value of Addr3
                if (InvoiceRet.ShipAddress.Addr3 != null)
                {
                    string Addr340 = (string)InvoiceRet.ShipAddress.Addr3.GetValue();
                }
                //Get value of Addr4
                if (InvoiceRet.ShipAddress.Addr4 != null)
                {
                    string Addr441 = (string)InvoiceRet.ShipAddress.Addr4.GetValue();
                }
                //Get value of Addr5
                if (InvoiceRet.ShipAddress.Addr5 != null)
                {
                    string Addr542 = (string)InvoiceRet.ShipAddress.Addr5.GetValue();
                }
                //Get value of City
                if (InvoiceRet.ShipAddress.City != null)
                {
                    string City43 = (string)InvoiceRet.ShipAddress.City.GetValue();
                }
                //Get value of State
                if (InvoiceRet.ShipAddress.State != null)
                {
                    string State44 = (string)InvoiceRet.ShipAddress.State.GetValue();
                }
                //Get value of PostalCode
                if (InvoiceRet.ShipAddress.PostalCode != null)
                {
                    string PostalCode45 = (string)InvoiceRet.ShipAddress.PostalCode.GetValue();
                }
                //Get value of Country
                if (InvoiceRet.ShipAddress.Country != null)
                {
                    string Country46 = (string)InvoiceRet.ShipAddress.Country.GetValue();
                }
                //Get value of Note
                if (InvoiceRet.ShipAddress.Note != null)
                {
                    string Note47 = (string)InvoiceRet.ShipAddress.Note.GetValue();
                }
            }
            if (InvoiceRet.ShipAddressBlock != null)
            {
                //Get value of Addr1
                if (InvoiceRet.ShipAddressBlock.Addr1 != null)
                {
                    string Addr148 = (string)InvoiceRet.ShipAddressBlock.Addr1.GetValue();
                }
                //Get value of Addr2
                if (InvoiceRet.ShipAddressBlock.Addr2 != null)
                {
                    string Addr249 = (string)InvoiceRet.ShipAddressBlock.Addr2.GetValue();
                }
                //Get value of Addr3
                if (InvoiceRet.ShipAddressBlock.Addr3 != null)
                {
                    string Addr350 = (string)InvoiceRet.ShipAddressBlock.Addr3.GetValue();
                }
                //Get value of Addr4
                if (InvoiceRet.ShipAddressBlock.Addr4 != null)
                {
                    string Addr451 = (string)InvoiceRet.ShipAddressBlock.Addr4.GetValue();
                }
                //Get value of Addr5
                if (InvoiceRet.ShipAddressBlock.Addr5 != null)
                {
                    string Addr552 = (string)InvoiceRet.ShipAddressBlock.Addr5.GetValue();
                }
            }
            //Get value of IsPending
            if (InvoiceRet.IsPending != null)
            {
                bool IsPending53 = (bool)InvoiceRet.IsPending.GetValue();
            }
            //Get value of IsFinanceCharge
            if (InvoiceRet.IsFinanceCharge != null)
            {
                bool IsFinanceCharge54 = (bool)InvoiceRet.IsFinanceCharge.GetValue();
            }
            //Get value of PONumber
            if (InvoiceRet.PONumber != null)
            {
                string PONumber55 = (string)InvoiceRet.PONumber.GetValue();
            }
            if (InvoiceRet.TermsRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.TermsRef.ListID != null)
                {
                    string ListID56 = (string)InvoiceRet.TermsRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.TermsRef.FullName != null)
                {
                    string FullName57 = (string)InvoiceRet.TermsRef.FullName.GetValue();
                }
            }
            //Get value of DueDate
            if (InvoiceRet.DueDate != null)
            {
                DateTime DueDate58 = (DateTime)InvoiceRet.DueDate.GetValue();
            }
            if (InvoiceRet.SalesRepRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.SalesRepRef.ListID != null)
                {
                    string ListID59 = (string)InvoiceRet.SalesRepRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.SalesRepRef.FullName != null)
                {
                    string FullName60 = (string)InvoiceRet.SalesRepRef.FullName.GetValue();
                }
            }
            //Get value of FOB
            if (InvoiceRet.FOB != null)
            {
                string FOB61 = (string)InvoiceRet.FOB.GetValue();
            }
            //Get value of ShipDate
            if (InvoiceRet.ShipDate != null)
            {
                DateTime ShipDate62 = (DateTime)InvoiceRet.ShipDate.GetValue();
            }
            if (InvoiceRet.ShipMethodRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.ShipMethodRef.ListID != null)
                {
                    string ListID63 = (string)InvoiceRet.ShipMethodRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.ShipMethodRef.FullName != null)
                {
                    string FullName64 = (string)InvoiceRet.ShipMethodRef.FullName.GetValue();
                }
            }
            //Get value of Subtotal
            if (InvoiceRet.Subtotal != null)
            {
                double Subtotal65 = (double)InvoiceRet.Subtotal.GetValue();
            }
            if (InvoiceRet.ItemSalesTaxRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.ItemSalesTaxRef.ListID != null)
                {
                    string ListID66 = (string)InvoiceRet.ItemSalesTaxRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.ItemSalesTaxRef.FullName != null)
                {
                    string FullName67 = (string)InvoiceRet.ItemSalesTaxRef.FullName.GetValue();
                }
            }
            //Get value of SalesTaxPercentage
            if (InvoiceRet.SalesTaxPercentage != null)
            {
                double SalesTaxPercentage68 = (double)InvoiceRet.SalesTaxPercentage.GetValue();
            }
            //Get value of SalesTaxTotal
            if (InvoiceRet.SalesTaxTotal != null)
            {
                double SalesTaxTotal69 = (double)InvoiceRet.SalesTaxTotal.GetValue();
            }
            //Get value of AppliedAmount
            if (InvoiceRet.AppliedAmount != null)
            {
                double AppliedAmount70 = (double)InvoiceRet.AppliedAmount.GetValue();
            }
            //Get value of BalanceRemaining
            if (InvoiceRet.BalanceRemaining != null)
            {
                double BalanceRemaining71 = (double)InvoiceRet.BalanceRemaining.GetValue();
            }
            if (InvoiceRet.CurrencyRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.CurrencyRef.ListID != null)
                {
                    string ListID72 = (string)InvoiceRet.CurrencyRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.CurrencyRef.FullName != null)
                {
                    string FullName73 = (string)InvoiceRet.CurrencyRef.FullName.GetValue();
                }
            }
            //Get value of ExchangeRate
            if (InvoiceRet.ExchangeRate != null)
            {
                //IQBFloatType ExchangeRate74 = (IQBFloatType)InvoiceRet.ExchangeRate.GetValue();
            }
            //Get value of BalanceRemainingInHomeCurrency
            if (InvoiceRet.BalanceRemainingInHomeCurrency != null)
            {
                double BalanceRemainingInHomeCurrency75 = (double)InvoiceRet.BalanceRemainingInHomeCurrency.GetValue();
            }
            //Get value of Memo
            if (InvoiceRet.Memo != null)
            {
                string Memo76 = (string)InvoiceRet.Memo.GetValue();
            }
            //Get value of IsPaid
            if (InvoiceRet.IsPaid != null)
            {
                bool IsPaid77 = (bool)InvoiceRet.IsPaid.GetValue();
            }
            if (InvoiceRet.CustomerMsgRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.CustomerMsgRef.ListID != null)
                {
                    string ListID78 = (string)InvoiceRet.CustomerMsgRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.CustomerMsgRef.FullName != null)
                {
                    string FullName79 = (string)InvoiceRet.CustomerMsgRef.FullName.GetValue();
                }
            }
            //Get value of IsToBePrinted
            if (InvoiceRet.IsToBePrinted != null)
            {
                bool IsToBePrinted80 = (bool)InvoiceRet.IsToBePrinted.GetValue();
            }
            //Get value of IsToBeEmailed
            if (InvoiceRet.IsToBeEmailed != null)
            {
                bool IsToBeEmailed81 = (bool)InvoiceRet.IsToBeEmailed.GetValue();
            }
            if (InvoiceRet.CustomerSalesTaxCodeRef != null)
            {
                //Get value of ListID
                if (InvoiceRet.CustomerSalesTaxCodeRef.ListID != null)
                {
                    string ListID82 = (string)InvoiceRet.CustomerSalesTaxCodeRef.ListID.GetValue();
                }
                //Get value of FullName
                if (InvoiceRet.CustomerSalesTaxCodeRef.FullName != null)
                {
                    string FullName83 = (string)InvoiceRet.CustomerSalesTaxCodeRef.FullName.GetValue();
                }
            }
            //Get value of SuggestedDiscountAmount
            if (InvoiceRet.SuggestedDiscountAmount != null)
            {
                double SuggestedDiscountAmount84 = (double)InvoiceRet.SuggestedDiscountAmount.GetValue();
            }
            //Get value of SuggestedDiscountDate
            if (InvoiceRet.SuggestedDiscountDate != null)
            {
                DateTime SuggestedDiscountDate85 = (DateTime)InvoiceRet.SuggestedDiscountDate.GetValue();
            }
            //Get value of Other
            if (InvoiceRet.Other != null)
            {
                string Other86 = (string)InvoiceRet.Other.GetValue();
            }
            //Get value of ExternalGUID
            if (InvoiceRet.ExternalGUID != null)
            {
                string ExternalGUID87 = (string)InvoiceRet.ExternalGUID.GetValue();
            }
            if (InvoiceRet.LinkedTxnList != null)
            {
                for (int i88 = 0; i88 < InvoiceRet.LinkedTxnList.Count; i88++)
                {
                    ILinkedTxn LinkedTxn = InvoiceRet.LinkedTxnList.GetAt(i88);
                    //Get value of TxnID
                    string TxnID89 = (string)LinkedTxn.TxnID.GetValue();
                    //Get value of TxnType
                    ENTxnType TxnType90 = (ENTxnType)LinkedTxn.TxnType.GetValue();
                    //Get value of TxnDate
                    DateTime TxnDate91 = (DateTime)LinkedTxn.TxnDate.GetValue();
                    //Get value of RefNumber
                    if (LinkedTxn.RefNumber != null)
                    {
                        string RefNumber92 = (string)LinkedTxn.RefNumber.GetValue();
                    }
                    //Get value of LinkType
                    if (LinkedTxn.LinkType != null)
                    {
                        ENLinkType LinkType93 = (ENLinkType)LinkedTxn.LinkType.GetValue();
                    }
                    //Get value of Amount
                    double Amount94 = (double)LinkedTxn.Amount.GetValue();
                }
            }
            if (InvoiceRet.ORInvoiceLineRetList != null)
            {
                for (int i95 = 0; i95 < InvoiceRet.ORInvoiceLineRetList.Count; i95++)
                {
                    IORInvoiceLineRet ORInvoiceLineRet = InvoiceRet.ORInvoiceLineRetList.GetAt(i95);
                    if (ORInvoiceLineRet.InvoiceLineRet != null)
                    {
                        if (ORInvoiceLineRet.InvoiceLineRet != null)
                        {
                            //Get value of TxnLineID
                            string TxnLineID96 = (string)ORInvoiceLineRet.InvoiceLineRet.TxnLineID.GetValue();
                            if (ORInvoiceLineRet.InvoiceLineRet.ItemRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID != null)
                                {
                                    string ListID97 = (string)ORInvoiceLineRet.InvoiceLineRet.ItemRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName != null)
                                {
                                    string FullName98 = (string)ORInvoiceLineRet.InvoiceLineRet.ItemRef.FullName.GetValue();
                                }
                            }
                            //Get value of Desc
                            if (ORInvoiceLineRet.InvoiceLineRet.Desc != null)
                            {
                                string Desc99 = (string)ORInvoiceLineRet.InvoiceLineRet.Desc.GetValue();
                            }
                            //Get value of Quantity
                            if (ORInvoiceLineRet.InvoiceLineRet.Quantity != null)
                            {
                                int Quantity100 = (int)ORInvoiceLineRet.InvoiceLineRet.Quantity.GetValue();
                            }
                            //Get value of UnitOfMeasure
                            if (ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure != null)
                            {
                                string UnitOfMeasure101 = (string)ORInvoiceLineRet.InvoiceLineRet.UnitOfMeasure.GetValue();
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                                {
                                    string ListID102 = (string)ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                                {
                                    string FullName103 = (string)ORInvoiceLineRet.InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                                }
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.ORRate != null)
                            {
                                if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                                {
                                    //Get value of Rate
                                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate != null)
                                    {
                                        double Rate104 = (double)ORInvoiceLineRet.InvoiceLineRet.ORRate.Rate.GetValue();
                                    }
                                }
                                if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                                {
                                    //Get value of RatePercent
                                    if (ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent != null)
                                    {
                                        double RatePercent105 = (double)ORInvoiceLineRet.InvoiceLineRet.ORRate.RatePercent.GetValue();
                                    }
                                }
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.ClassRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID != null)
                                {
                                    string ListID106 = (string)ORInvoiceLineRet.InvoiceLineRet.ClassRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName != null)
                                {
                                    string FullName107 = (string)ORInvoiceLineRet.InvoiceLineRet.ClassRef.FullName.GetValue();
                                }
                            }
                            //Get value of Amount
                            if (ORInvoiceLineRet.InvoiceLineRet.Amount != null)
                            {
                                double Amount108 = (double)ORInvoiceLineRet.InvoiceLineRet.Amount.GetValue();
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID != null)
                                {
                                    string ListID109 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName != null)
                                {
                                    string FullName110 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                                }
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                                {
                                    string ListID111 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                                {
                                    string FullName112 = (string)ORInvoiceLineRet.InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                                }
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber != null)
                            {
                                if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                                {
                                    //Get value of SerialNumber
                                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                                    {
                                        string SerialNumber113 = (string)ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                                    }
                                }
                                if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                                {
                                    //Get value of LotNumber
                                    if (ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                                    {
                                        string LotNumber114 = (string)ORInvoiceLineRet.InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                                    }
                                }
                            }
                            //Get value of ServiceDate
                            if (ORInvoiceLineRet.InvoiceLineRet.ServiceDate != null)
                            {
                                DateTime ServiceDate115 = (DateTime)ORInvoiceLineRet.InvoiceLineRet.ServiceDate.GetValue();
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                                {
                                    string ListID116 = (string)ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                                {
                                    string FullName117 = (string)ORInvoiceLineRet.InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                                }
                            }
                            //Get value of Other1
                            if (ORInvoiceLineRet.InvoiceLineRet.Other1 != null)
                            {
                                string Other1118 = (string)ORInvoiceLineRet.InvoiceLineRet.Other1.GetValue();
                            }
                            //Get value of Other2
                            if (ORInvoiceLineRet.InvoiceLineRet.Other2 != null)
                            {
                                string Other2119 = (string)ORInvoiceLineRet.InvoiceLineRet.Other2.GetValue();
                            }
                            if (ORInvoiceLineRet.InvoiceLineRet.DataExtRetList != null)
                            {
                                for (int i120 = 0; i120 < ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.Count; i120++)
                                {
                                    IDataExtRet DataExtRet = ORInvoiceLineRet.InvoiceLineRet.DataExtRetList.GetAt(i120);
                                    //Get value of OwnerID
                                    if (DataExtRet.OwnerID != null)
                                    {
                                        string OwnerID121 = (string)DataExtRet.OwnerID.GetValue();
                                    }
                                    //Get value of DataExtName
                                    string DataExtName122 = (string)DataExtRet.DataExtName.GetValue();
                                    //Get value of DataExtType
                                    ENDataExtType DataExtType123 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                                    //Get value of DataExtValue
                                    string DataExtValue124 = (string)DataExtRet.DataExtValue.GetValue();
                                }
                            }
                        }
                    }
                    if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                    {
                        if (ORInvoiceLineRet.InvoiceLineGroupRet != null)
                        {
                            //Get value of TxnLineID
                            string TxnLineID125 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.TxnLineID.GetValue();
                            //Get value of ListID
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID != null)
                            {
                                string ListID126 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.ListID.GetValue();
                            }
                            //Get value of FullName
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName != null)
                            {
                                string FullName127 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue();
                            }
                            //Get value of Desc
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.Desc != null)
                            {
                                string Desc128 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.Desc.GetValue();
                            }
                            //Get value of Quantity
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.Quantity != null)
                            {
                                int Quantity129 = (int)ORInvoiceLineRet.InvoiceLineGroupRet.Quantity.GetValue();
                            }
                            //Get value of UnitOfMeasure
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure != null)
                            {
                                string UnitOfMeasure130 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.UnitOfMeasure.GetValue();
                            }
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef != null)
                            {
                                //Get value of ListID
                                if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID != null)
                                {
                                    string ListID131 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.ListID.GetValue();
                                }
                                //Get value of FullName
                                if (ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName != null)
                                {
                                    string FullName132 = (string)ORInvoiceLineRet.InvoiceLineGroupRet.OverrideUOMSetRef.FullName.GetValue();
                                }
                            }
                            //Get value of IsPrintItemsInGroup
                            bool IsPrintItemsInGroup133 = (bool)ORInvoiceLineRet.InvoiceLineGroupRet.IsPrintItemsInGroup.GetValue();
                            //Get value of TotalAmount
                            double TotalAmount134 = (double)ORInvoiceLineRet.InvoiceLineGroupRet.TotalAmount.GetValue();
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList != null)
                            {
                                for (int i135 = 0; i135 < ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.Count; i135++)
                                {
                                    IInvoiceLineRet InvoiceLineRet = ORInvoiceLineRet.InvoiceLineGroupRet.InvoiceLineRetList.GetAt(i135);
                                    //Get value of TxnLineID
                                    string TxnLineID136 = (string)InvoiceLineRet.TxnLineID.GetValue();
                                    if (InvoiceLineRet.ItemRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.ItemRef.ListID != null)
                                        {
                                            string ListID137 = (string)InvoiceLineRet.ItemRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.ItemRef.FullName != null)
                                        {
                                            string FullName138 = (string)InvoiceLineRet.ItemRef.FullName.GetValue();
                                        }
                                    }
                                    //Get value of Desc
                                    if (InvoiceLineRet.Desc != null)
                                    {
                                        string Desc139 = (string)InvoiceLineRet.Desc.GetValue();
                                    }
                                    //Get value of Quantity
                                    if (InvoiceLineRet.Quantity != null)
                                    {
                                        int Quantity140 = (int)InvoiceLineRet.Quantity.GetValue();
                                    }
                                    //Get value of UnitOfMeasure
                                    if (InvoiceLineRet.UnitOfMeasure != null)
                                    {
                                        string UnitOfMeasure141 = (string)InvoiceLineRet.UnitOfMeasure.GetValue();
                                    }
                                    if (InvoiceLineRet.OverrideUOMSetRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.OverrideUOMSetRef.ListID != null)
                                        {
                                            string ListID142 = (string)InvoiceLineRet.OverrideUOMSetRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.OverrideUOMSetRef.FullName != null)
                                        {
                                            string FullName143 = (string)InvoiceLineRet.OverrideUOMSetRef.FullName.GetValue();
                                        }
                                    }
                                    if (InvoiceLineRet.ORRate != null)
                                    {
                                        if (InvoiceLineRet.ORRate.Rate != null)
                                        {
                                            //Get value of Rate
                                            if (InvoiceLineRet.ORRate.Rate != null)
                                            {
                                                double Rate144 = (double)InvoiceLineRet.ORRate.Rate.GetValue();
                                            }
                                        }
                                        if (InvoiceLineRet.ORRate.RatePercent != null)
                                        {
                                            //Get value of RatePercent
                                            if (InvoiceLineRet.ORRate.RatePercent != null)
                                            {
                                                double RatePercent145 = (double)InvoiceLineRet.ORRate.RatePercent.GetValue();
                                            }
                                        }
                                    }
                                    if (InvoiceLineRet.ClassRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.ClassRef.ListID != null)
                                        {
                                            string ListID146 = (string)InvoiceLineRet.ClassRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.ClassRef.FullName != null)
                                        {
                                            string FullName147 = (string)InvoiceLineRet.ClassRef.FullName.GetValue();
                                        }
                                    }
                                    //Get value of Amount
                                    if (InvoiceLineRet.Amount != null)
                                    {
                                        double Amount148 = (double)InvoiceLineRet.Amount.GetValue();
                                    }
                                    if (InvoiceLineRet.InventorySiteRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.InventorySiteRef.ListID != null)
                                        {
                                            string ListID149 = (string)InvoiceLineRet.InventorySiteRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.InventorySiteRef.FullName != null)
                                        {
                                            string FullName150 = (string)InvoiceLineRet.InventorySiteRef.FullName.GetValue();
                                        }
                                    }
                                    if (InvoiceLineRet.InventorySiteLocationRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.InventorySiteLocationRef.ListID != null)
                                        {
                                            string ListID151 = (string)InvoiceLineRet.InventorySiteLocationRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.InventorySiteLocationRef.FullName != null)
                                        {
                                            string FullName152 = (string)InvoiceLineRet.InventorySiteLocationRef.FullName.GetValue();
                                        }
                                    }
                                    if (InvoiceLineRet.ORSerialLotNumber != null)
                                    {
                                        if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                                        {
                                            //Get value of SerialNumber
                                            if (InvoiceLineRet.ORSerialLotNumber.SerialNumber != null)
                                            {
                                                string SerialNumber153 = (string)InvoiceLineRet.ORSerialLotNumber.SerialNumber.GetValue();
                                            }
                                        }
                                        if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                                        {
                                            //Get value of LotNumber
                                            if (InvoiceLineRet.ORSerialLotNumber.LotNumber != null)
                                            {
                                                string LotNumber154 = (string)InvoiceLineRet.ORSerialLotNumber.LotNumber.GetValue();
                                            }
                                        }
                                    }
                                    //Get value of ServiceDate
                                    if (InvoiceLineRet.ServiceDate != null)
                                    {
                                        DateTime ServiceDate155 = (DateTime)InvoiceLineRet.ServiceDate.GetValue();
                                    }
                                    if (InvoiceLineRet.SalesTaxCodeRef != null)
                                    {
                                        //Get value of ListID
                                        if (InvoiceLineRet.SalesTaxCodeRef.ListID != null)
                                        {
                                            string ListID156 = (string)InvoiceLineRet.SalesTaxCodeRef.ListID.GetValue();
                                        }
                                        //Get value of FullName
                                        if (InvoiceLineRet.SalesTaxCodeRef.FullName != null)
                                        {
                                            string FullName157 = (string)InvoiceLineRet.SalesTaxCodeRef.FullName.GetValue();
                                        }
                                    }
                                    //Get value of Other1
                                    if (InvoiceLineRet.Other1 != null)
                                    {
                                        string Other1158 = (string)InvoiceLineRet.Other1.GetValue();
                                    }
                                    //Get value of Other2
                                    if (InvoiceLineRet.Other2 != null)
                                    {
                                        string Other2159 = (string)InvoiceLineRet.Other2.GetValue();
                                    }
                                    if (InvoiceLineRet.DataExtRetList != null)
                                    {
                                        for (int i160 = 0; i160 < InvoiceLineRet.DataExtRetList.Count; i160++)
                                        {
                                            IDataExtRet DataExtRet = InvoiceLineRet.DataExtRetList.GetAt(i160);
                                            //Get value of OwnerID
                                            if (DataExtRet.OwnerID != null)
                                            {
                                                string OwnerID161 = (string)DataExtRet.OwnerID.GetValue();
                                            }
                                            //Get value of DataExtName
                                            string DataExtName162 = (string)DataExtRet.DataExtName.GetValue();
                                            //Get value of DataExtType
                                            ENDataExtType DataExtType163 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                                            //Get value of DataExtValue
                                            string DataExtValue164 = (string)DataExtRet.DataExtValue.GetValue();
                                        }
                                    }
                                }
                            }
                            if (ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList != null)
                            {
                                for (int i165 = 0; i165 < ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList.Count; i165++)
                                {
                                    IDataExtRet DataExtRet = ORInvoiceLineRet.InvoiceLineGroupRet.DataExtRetList.GetAt(i165);
                                    //Get value of OwnerID
                                    if (DataExtRet.OwnerID != null)
                                    {
                                        string OwnerID166 = (string)DataExtRet.OwnerID.GetValue();
                                    }
                                    //Get value of DataExtName
                                    string DataExtName167 = (string)DataExtRet.DataExtName.GetValue();
                                    //Get value of DataExtType
                                    ENDataExtType DataExtType168 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                                    //Get value of DataExtValue
                                    string DataExtValue169 = (string)DataExtRet.DataExtValue.GetValue();
                                }
                            }
                        }
                    }
                }
            }
            if (InvoiceRet.DataExtRetList != null)
            {
                for (int i170 = 0; i170 < InvoiceRet.DataExtRetList.Count; i170++)
                {
                    IDataExtRet DataExtRet = InvoiceRet.DataExtRetList.GetAt(i170);
                    //Get value of OwnerID
                    if (DataExtRet.OwnerID != null)
                    {
                        string OwnerID171 = (string)DataExtRet.OwnerID.GetValue();
                    }
                    //Get value of DataExtName
                    string DataExtName172 = (string)DataExtRet.DataExtName.GetValue();
                    //Get value of DataExtType
                    ENDataExtType DataExtType173 = (ENDataExtType)DataExtRet.DataExtType.GetValue();
                    //Get value of DataExtValue
                    string DataExtValue174 = (string)DataExtRet.DataExtValue.GetValue();
                }
            }
        }



    }
}