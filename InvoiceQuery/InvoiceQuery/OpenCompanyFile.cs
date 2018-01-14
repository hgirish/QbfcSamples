using QBFC13Lib;
using System;
using System.Configuration;

namespace InvoiceQuery
{
    public class OpenCompanyFile
    {
        public void OpenQB()
        {
            QBSessionManager sessionManager = null;
            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                sessionManager.OpenConnection("QBAPI", "Quickbooks SDK Demo Test");
                string qbFile = ConfigurationSettings.AppSettings["companyfile"].ToString();
                Console.WriteLine(qbFile);
                sessionManager.BeginSession(
                    qbFile, ENOpenMode.omDontCare);

                ICustomerQuery customerQueryRq = requestMsgSet.AppendCustomerQueryRq();
                customerQueryRq.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue(ENActiveStatus.asAll);


                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }
            catch (Exception ex)
            {

                System.Console.WriteLine(ex.Message);
                sessionManager.EndSession();
                sessionManager.CloseConnection();

            }
        }
    }
}
