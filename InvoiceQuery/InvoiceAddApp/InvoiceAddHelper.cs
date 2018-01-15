using QBFC13Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceAddApp
{
   public class InvoiceAddHelper
    {
        public void DoInvoiceAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildInvoiceAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Invoice Add Console App");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);


                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkInvoiceAddRs(responseMsgSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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

        private void WalkInvoiceAddRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null)
            {
                return;
            }
            IResponseList responseList = responseMsgSet.ResponseList;

            if (responseList == null)
            {
                return;
            }
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtInvoiceAddRs)
                        {
                            IInvoiceRet invoiceRet = (IInvoiceRet)response.Detail;
                            WalkInvoiceRet(invoiceRet);
                        }
                    }
                }
            }
            Console.WriteLine(responseMsgSet.ToXMLString());
        }

        private void WalkInvoiceRet(IInvoiceRet invoiceRet)
        {
            if (invoiceRet == null)
            {
                return;
            }
            string txtId = invoiceRet.TxnID.GetValue();
            var txtNumber = invoiceRet.TxnNumber?.GetValue();
            var invoiceNumber = invoiceRet.RefNumber?.GetValue();
            Console.WriteLine($"{invoiceNumber}\t{txtId}\t{txtNumber}");

        }

        private void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet)
        {
            IInvoiceAdd invoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
            invoiceAddRq.CustomerRef.FullName.SetValue("City");
            invoiceAddRq.Other.SetValue("12345");
            invoiceAddRq.TxnDate.SetValue(DateTime.Now);

            IORInvoiceLineAdd oRInvoiceLineAddList = invoiceAddRq.ORInvoiceLineAddList.Append();
            string orInvoiceListElementType = "InvoiceLineAdd";

            if (orInvoiceListElementType == "InvoiceLineAdd")
            {
               // oRInvoiceLineAddList.InvoiceLineAdd.ItemRef.FullName.SetValue("Storage");
                oRInvoiceLineAddList.InvoiceLineAdd.Desc.SetValue("January Storage");
               // oRInvoiceLineAddList.InvoiceLineAdd.Quantity.SetValue(2);
               // oRInvoiceLineAddList.InvoiceLineAdd.Amount.SetValue(9.99);
               // oRInvoiceLineAddList.InvoiceLineAdd.UnitOfMeasure.SetValue("each");
              //  oRInvoiceLineAddList.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(9.99);
            }

            // Add blank line
            var invLine3 = invoiceAddRq.ORInvoiceLineAddList.Append();
            var line3 = invLine3.InvoiceLineAdd;
            line3.Desc.SetEmpty();



            IORInvoiceLineAdd oRInvoiceLineAddList2 = invoiceAddRq.ORInvoiceLineAddList.Append();
            var invoiceLineAdd2 = oRInvoiceLineAddList2.InvoiceLineAdd;
            if (orInvoiceListElementType == "InvoiceLineAdd")
            {
                invoiceLineAdd2.ItemRef.FullName.SetValue("Storage");
                invoiceLineAdd2.Desc.SetValue("Test Transaction Item");
                invoiceLineAdd2.Quantity.SetValue(2);
                // oRInvoiceLineAddList.InvoiceLineAdd.Amount.SetValue(9.99);
                // oRInvoiceLineAddList.InvoiceLineAdd.UnitOfMeasure.SetValue("each");
                invoiceLineAdd2.ORRatePriceLevel.Rate.SetValue(9.99);
            }
            //invoiceAddRq.IncludeRetElementList.Add("ab");








            Console.WriteLine(requestMsgSet.ToXMLString());
        }
    }
}
