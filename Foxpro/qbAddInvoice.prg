#include qbfc.h
Local sessionBegun, connectionOpen, lcqbsdkver, SessionManager
**, requestmsgset
**Local responseMsgSet
Clear



sessionBegun = .F.
connectionOpen = .F.

lcqbsdkver="QBFC13.QBSessionManager" && change to the SDK version you are using
Try
	SessionManager = Createobject(lcqbsdkver)

	requestmsgset= SessionManager.createmsgsetrequest("US",13, 0)
	requestmsgset.Attributes.onerror = roecontinue

	BuildInvoiceAddRq(requestmsgset)

	SessionManager.OpenConnection2("","Foxpro Qbfc", ctLocalQBD)
	connectionOpen = .T.
*Take a look down the qbfc.h file and see if you can find the ctlocalqbd.  Now you see what that file does for you
	SessionManager.BeginSession("", omdontcare)
	sessionBegun = .T.

	responseMsgSet = SessionManager.DoRequests(requestmsgset)

	SessionManager.EndSession()
	sessionBegun = .F.
	SessionManager.CloseConnection()
	connectionOpen = .F.

	WalkInvoiceAddRs(responseMsgSet)





Catch To ex
	?ex.Message
	If sessionBegun
		SessionManager.EndSession()
	Endif
	If connectionOpen
		SessionManager.CloseConnection
	Endif
Finally
	If sessionBegun
		SessionManager.EndSession()
	Endif
	If connectionOpen
		SessionManager.CloseConnection
	Endif
Endtry




* Relinquish the communication channel with QuickBooks
SessionManager.EndSession
SessionManager.CloseConnection
Return

******************
Procedure BuildInvoiceAddRq(requestmsgset)
invoiceAddRq = requestmsgset.AppendInvoiceAddRq()
invoiceAddRq.CustomerRef.FullName.SetValue("City") && Customer name Must exist in system
invoiceAddRq.Other.SetValue("12345")
invoiceAddRq.TxnDate.SetValue(Date())

** First Detail Line
oRInvoiceLineAddList = invoiceAddRq.oRInvoiceLineAddList.Append()
oRInvoiceLineAddList.InvoiceLineAdd.Desc.SetValue("January Storage")

** add blank line
oRInvoiceLineAddList = invoiceAddRq.oRInvoiceLineAddList.Append()
oRInvoiceLineAddList.InvoiceLineAdd.Desc.SetEmpty()

** Third line 
oRInvoiceLineAddList = invoiceAddRq.oRInvoiceLineAddList.Append()
WITH oRInvoiceLineAddList.InvoiceLineAdd 
.ItemRef.FullName.SetValue("Storage") && Item category must exist in system
.Desc.SetValue("Test Transaction Item")
.Quantity.SetValue(2)
.ORRatePriceLevel.Rate.SetValue(9.99)

ENDWITH
xml = requestMsgSet.ToXMLString()
?xml
Endproc
********************
Procedure WalkInvoiceAddRs(responseMsgSet)
xml = responseMsgSet.ToXMLString()
?xml
Endproc

*********************

Function qbfclatestversion
Parameter strxmlversions
Local lastvers, Vers, qbfclatestversion, i
strxmlversions = SessionManager.qbxmlversionsforsession
lastvers = 0
For i = 1 To Alen(strxmlversions)
	Vers = Val(strxmlversions(i))
	If (Vers > lastvers)
		lastvers = Vers
		qbfclatestversion = strxmlversions(i)
	Endif
Next i
Return qbfclatestversion
************************************************
