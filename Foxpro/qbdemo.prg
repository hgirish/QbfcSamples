#include qbfc.h
Local minvno, mfile, fromDate, toDate, sessionBegun, connectionOpen, lcqbsdkver, SessionManager
Local requestmsgset, invoiceQueryRq, responsemsgset, response2, InvoiceRetList, InvoiceRet, transid, editseq, invoicelineretlist, FullName, mtxnlineid, mfullname, mparent, i, k

Clear
minvno  = 307
mfile = 'qbdemo'
fromDate = Date(2018,1,5)
toDate = Date(2018,1,5)

sessionBegun = .F.
connectionOpen = .F.

lcqbsdkver="QBFC13.QBSessionManager" && change to the SDK version you are using
Try
	SessionManager = Createobject(lcqbsdkver)

	SessionManager.OpenConnection2("","Foxpro Qbfc", ctLocalQBD)
	connectionOpen = .T.
*Take a look down the qbfc.h file and see if you can find the ctlocalqbd.  Now you see what that file does for you
	SessionManager.BeginSession("", omdontcare)
	sessionBegun = .T.

*!*	supportedversion = qbfclatestversion(sessionmanager)
*!*	?"version",supportedversion
*Rem request #1 to determine last close date
	requestmsgset= SessionManager.createmsgsetrequest("US",13, 0)

****************
	requestmsgset.Attributes.onerror = roecontinue
* construct the request to get the one invoice, its line items and custom fields

	invoiceQueryRq = requestmsgset.AppendInvoiceQueryRq
	invoiceQueryRq.IncludeLineItems.SetValue(.T.)
	invoiceQueryRq.OwnerIDList.Add ("0")

*!*	invoiceQueryRq.orinvoicequery.invoicefilter.orrefnumberfilter. ;
*!*	refnumberfilter.matchcriterion.setvalue(0)
*!*	** 0 - Starts With, 1 - Contains, 2 - Ends With
*!*	invoiceQueryRq.orinvoicequery.invoicefilter.orrefnumberfilter. ;

*!*	refnumberfilter.refnumber.setvalue(minvno)


	invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(fromDate)
	invoiceQueryRq.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(toDate)

	responsemsgset = SessionManager.DoRequests(requestmsgset)
	response2 = responsemsgset.ResponseList.GetAt(0)



	If Isnull(response2.Detail)
		=Messagebox("No Detail Available for invoice " + minvno)
		Return .F.
	Endif
	InvoiceRetList = response2.Detail
	invoiceCounts =  InvoiceRetList.Count
	?invoiceCounts , " invoices found"
	If Isnull( InvoiceRetList )
		?"NO invoices found"
		Return
	Endif

	Create Cursor curInvoice(InvNum c(20), InvDate Datetime, JobNum c(20), CustName c(50), Amount Y, InvMemo c(150),Name c(50),FullName c(50),CompanyName c(50), Descr c(150))
	For i = 0 To (invoiceCounts -1)
		InvoiceRet = InvoiceRetList.GetAt(i)
		QuickBooksID = InvoiceRet.TxnID.GetValue()
		EditSequence = InvoiceRet.EditSequence.GetValue()
		lcInvoiceNumber = InvoiceRet.RefNumber.GetValue()
		ldInvoiceDate = InvoiceRet.TimeCreated.GetValue()
		lcInvMemo = InvoiceRet.Memo.GetValue()
		lcJobNumber = InvoiceRet.Other.GetValue()
		lcCustomerName = InvoiceRet.CustomerRef.FullName.GetValue()
		lnAmount = InvoiceRet.BalanceRemaining.GetValue()
		If Amount = 0
			lnAmount = InvoiceRet.Subtotal.GetValue()
		Endif

		Insert Into curInvoice(InvNum,InvDate,JobNum, CustName,Amount,InvMemo) Values (lcInvoiceNumber,ldInvoiceDate,Alltrim(lcJobNumber),Alltrim(lcCustomerName),lnAmount,Alltrim(lcInvMemo))

		customerListId = InvoiceRet.CustomerRef.ListID.GetValue()
		If Not Isnull( customerListId)
			
			requestmsgset.ClearRequests()
			customerQueryRq = requestmsgset.AppendCustomerQueryRq()
			customerQueryRq.ORCustomerListQuery.ListIDList.Add(customerListId)

			responsemsgset = SessionManager.DoRequests(requestmsgset)

			response = responsemsgset.ResponseList.GetAt(0)

			customerRetList = response.Detail

			customerRet = customerRetList.GetAt(0)
			?customerRet.ListID.GetValue(),                 customerRet.EditSequence.GetValue()
			lcName = customerRet.Name.GetValue()
			lcFullName = customerRet.FullName.GetValue()
			lcCompanyName = customerRet.CompanyName.GetValue()
			Update curInvoice Set Name = Alltrim(lcName), FullName = Alltrim(lcFullName), CompanyName = Alltrim(lcCompanyName) Where InvNum = lcInvoiceNumber


			If Not Isnull(InvoiceRet.ORInvoiceLineRetList)
				
				lcDescr = ""
				lnI = 0
				lnInvoiceLinecount = InvoiceRet.ORInvoiceLineRetList.Count
				
				If lnInvoiceLinecount  > 0
					Do While  Empty(lcDescr) And lnI < lnInvoiceLinecount
						
						orInvoiceLineRet = InvoiceRet.ORInvoiceLineRetList.GetAt(lnI)
						IF NOT ISNULL(orInvoiceLineRet.InvoiceLineRet.Desc)
						lcDescr = orInvoiceLineRet.InvoiceLineRet.Desc.GetValue()
						ENDIF
						
						
						lnI = lnI + 1
					Enddo

				Endif
				If Not Empty(lcDescr)
					Update curInvoice Set Descr = Alltrim(lcDescr) Where InvNum = lcInvoiceNumber
				Endif

			Endif
		Endif
	Endfor
*!*		InvoiceRet = InvoiceRetList.GetAt(0)
*!*		transid = InvoiceRet.TxnID.GetValue
*!*		editseq = InvoiceRet.EditSequence.GetValue
*!*		invoicelineretlist = InvoiceRet.orinvoicelineretlist
*!*		Create Table (mfile) (Item c(40), qtyshp N(12,2),batchrq c(1),Parent c(40),scanned N(12,2),invno c(8),Batch c(10),scancnt N(12,2),scanitem c(15))
*!*		FullName = ""
*!*		mtxnlineid = " "
*!*		For i = 0 To InvoiceRet.orinvoicelineretlist.Count - 1
*!*			If Not Isnull(invoicelineretlist.GetAt(i).invoicelineret.itemref)
*!*				Append Blank
*!*				mfullname=invoicelineretlist.GetAt(i).invoicelineret.itemref.FullName.GetValue
*!*				mparent=mfullname
*!*				If ":" $ mfullname
*!*					mparent = Substr(mfullname,1,At(":",mfullname)-1)
*!*				Endif
*!*				Replace Item With mfullname,Parent With mparent
*!*				If Not Isnull(invoicelineretlist.GetAt(i).invoicelineret.quantity)
*!*					Replace qtyshp With invoicelineretlist.GetAt(i).invoicelineret.quantity.GetValue
*!*				Endif
*!*				If Not Isnull(invoicelineretlist.GetAt(i).invoicelineret.dataextretlist)
*!*					For k = 0 To invoicelineretlist.GetAt(i).invoicelineret.dataextretlist.Count - 1
*!*						If invoicelineretlist.GetAt(i).invoicelineret.dataextretlist.GetAt(k).dataextname.GetValue="Batch #"
*!*							Replace batchrq With invoicelineretlist.GetAt(i).invoicelineret.dataextretlist.GetAt(k).dataextvalue.GetValue
*!*						Endif
*!*						Replace invno With minvno
*!*					Next
*!*				Endif
*!*			Endif
*!*		Next

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
Select curInvoice
Browse

******************





* Relinquish the communication channel with QuickBooks
SessionManager.EndSession
SessionManager.CloseConnection


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
