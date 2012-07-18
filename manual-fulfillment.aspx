<%@ Page Language="VB" ContentType="text/html" Debug="True"%>
<%@ Register TagPrefix="MenuDisplay" TagName="Menu" src="menu.ascx"%>
<%@ Register TagPrefix="UpStatus" TagName="StatusChange" src="update_status.ascx"%>
<%@ Register TagPrefix="CreateInvoice" TagName="CreateInvoice" src="create-invoice.ascx"%>
<%@ Register Assembly="itextsharp" Namespace="itextsharp" TagPrefix="itextsharp" %>
<%@ Register TagPrefix="SendAppMail" TagName="AppMail" src="send_borrower_fulfillment.ascx"%>
<%@ Register TagPrefix="SendAppMail2" TagName="AppMail2" src="send_borrower_fulfillment-fpf-vp.ascx"%>
<%@ import Namespace="System.Configuration" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data.SQLClient" %>
<%@ import Namespace="System.Security.Cryptography" %>
<%@ import Namespace="System.Text" %>
<%@ import Namespace="itextsharp.text" %>
<%@ import Namespace="itextsharp.text.pdf" %>
<%@ import Namespace="System.Net.Sockets" %>
<%@ import Namespace="System.Net" %>
<%@ import Namespace="System.XML" %>
<%@ import Namespace="System.Web.Mail" %>
<script runat="server">
Sub Page_Load(Src As Object, E As EventArgs)
	Dim loggedIn As Integer=0
	IF Request.Cookies("ID") Is Nothing  OR Request.Cookies("Access") Is Nothing OR Request.Cookies("Flags") Is Nothing  OR Request.Cookies("Name") Is Nothing Then	
		Response.Redirect("default.aspx")
	ELSE 
		IF IsNumeric(Server.URLDecode(CStr(Request.Cookies("ID").Value))) AND (Server.URLDecode(CStr(Request.Cookies("Access").Value))="User" OR Server.URLDecode(CSTR(REquest.Cookies("Access").Value))="Manila") THEN
			loggedIn=1
			IF Server.URLDecode(CStr(Request.Cookies("Access").Value))="User" THEN
				lblAdminLink.Visible="True"
				menuDisplay01.Visible="True"
			ELSE
				lblAdminLink.Visible="False"
				menuDisplay01.Visible="False"							
			END IF			
		ELSE
			Response.Redirect("default.aspx")		
		END IF
	End If	
	IF NOT IsPostBack THEN
'		pnlZooQueue.Visible="False"	
'		chkSendtoZooQueue.Visible="False"
		lblInvoiceLink.Text=""	
		pnlInvoice.Visible="False"	
		lblError.Text=""
		txtEditID.Text="0"
		pnlBorrowerFulfillment.Visible="False"			
		IF loggedIn=1 THEN
menuDisplay01.securityCheck(Server.URLDecode(CStr(Request.Cookies("Flags").Value)),31011,0)
			lblWelcome.Text="Welcome, " & Server.URLDecode(CStr(Request.Cookies("Name").Value))
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnStr As String = MyConn
			Dim dbConnection As New SqlConnection(dbConnStr)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			Dim mailStringInvoice As String=""
			Dim redStr AS String=""
			spnCreateLink.innerHTML="<a href=" & CHR(34) & "create_invoice-fulfillment.aspx?id=" & removeInjection(Request.QueryString("id")) & CHR(34) & "><strong>Create Invoice</strong></a>"
			Try
				dbConnection.Open()
				sqlStr="SELECT  clients.id AS CID, clients.WebService,clients.UCDPClientID, appraisals.id AS APID,* FROM appraisals,clients WHERE clients.id=appraisals.ClientId AND appraisals.id=" & removeInjection(Request.QueryString("id"))
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Appraisals")

				sqlStr="SELECT * FROM products WHERE id=" & dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Product")
				
				Dim ucdpClient As INteger=0
				
				sqlStr="SELECT * FROM clients WHERE id=" & dbSet.Tables("Appraisals").Rows(0).Item("RCLientid")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"RCLient")									
				
				IF dbSet.Tables("RCLient").Rows.Count>0 THEN
					IF dbSet.Tables("RCLient").Rows(0).Item("UCDPProcessing")=1 THEN
						ucdpClient=dbSet.Tables("RCLient").Rows(0).Item("UCDPProcessing")
					END IF
				END IF
				IF ucdpClient=0 THEN
					IF dbSet.Tables("Appraisals").Rows(0).Item("UCDPProcessing")=1 THEN
						ucdpClient=dbSet.Tables("Appraisals").Rows(0).Item("UCDPProcessing")
					END IF					
				END IF
				
				txtBorrowerEmail.Text=dbSet.Tables("Appraisals").Rows(0).Item("BorrowerEMail")	
				chkBorrowerFulfillment.Checked="False"
				IF dbSet.Tables("Appraisals").Rows(0).Item("RequireBorrowerFulfillment")=1 THEN
					chkBorrowerFulfillment.Checked="True"
					chkBorrowerFulfillment.Enabled="False"	
					IF 	trim(dbSet.Tables("Appraisals").Rows(0).Item("BorrowerEMail"))="" THEN
						txtBorrowerEmail.Text="papermail@valuationpartners.com"
					END IF	
					pnlBorrowerFulfillment.Visible="True"								
				END IF 				
				IF dbSet.Tables("Appraisals").Rows(0).Item("BorrowerFulfillment")=1 THEN
					chkBorrowerFulfillment.Checked="True"		
					pnlBorrowerFulfillment.Visible="True"		
				END IF		
				IF  dbSet.Tables("Appraisals").Rows(0).Item("SendUADXML")
					lblSendUADXML.Text="1"
				END IF	
				IF  dbSet.Tables("Appraisals").Rows(0).Item("SendHVCC")
					lblSendHVCC.Text="1"
					chkSendHVCC.Checked="true"
					chkSendHVCC.Enabled="false"
				END IF	
				
				Dim sftpFulfillment AS Integer=0					
				sqlStr="SELECT * FROM tblClientReviewSupportProducts, clients WHERE RequireReviewSupportSelection=1 AND tblClientReviewSupportProducts.clientid=" & dbSet.Tables("Appraisals").Rows(0).Item("ClientID") & " AND productid=" & dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"sftpFulfillment")
				
				IF dbSet.Tables("sftpFulfillment").Rows.Count>0 THEN
					sftpFulfillment=1
				END IF				
				
								
				

				IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=3 OR dbSet.Tables("Appraisals").Rows(0).Item("WebService")=2  OR dbSet.Tables("Appraisals").Rows(0).Item("WebService")=4   OR dbSet.Tables("Appraisals").Rows(0).Item("WebService")=6 OR dbSet.Tables("Appraisals").Rows(0).Item("CID")=432 OR dbSet.Tables("Appraisals").Rows(0).Item("CID")=167 OR dbSet.Tables("Appraisals").Rows(0).Item("WebService")=7 OR sftpFulfillment=1 AND NOT (dbSet.Tables("Appraisals").Rows(0).Item("WebService")=7 AND dbSet.Tables("Appraisals").Rows(0).Item("ClientId")=108 AND dbSet.Tables("Appraisals").Rows(0).Item("VPWSOrder")=0)   THEN
				
					IF ucdpClient>0 AND ( (dbSet.Tables("Product").Rows(0).Item("UADProduct")=1 AND NOT dbSet.Tables("Product").Rows(0).Item("BypassUCDP")=1) OR dbSet.Tables("Appraisals").Rows(0).Item("UADOrder")=1)  AND NOT Request.QueryString("forceFulfillment")=1 AND NOT dbSet.Tables("Appraisals").Rows(0).Item("WebService")=3 AND NOT dbSet.Tables("Appraisals").Rows(0).Item("WebService")=4 THEN
'						response.Write(ucdpClient & "|" & dbSet.Tables("Product").Rows(0).Item("UADProduct") & "|" & dbSet.Tables("Appraisals").Rows(0).Item("UADOrder") & "|" & dbSet.Tables("Product").Rows(0).Item("BypassUCDP"))
						Response.Redirect("UCDP-processing.aspx?id=" & removeInjection(Request.QueryString("id")))
					ELSE				
						IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=3 THEN
							redStr="manual-fulfillment-realec.aspx?id=" &  removeInjection(Request.QueryString("id")) 
						END IF
						IF dbSet.Tables("Appraisals").Rows(0).Item("CID")=432  THEN
							redStr="manual-fulfillment-zip.aspx?id=" &  removeInjection(Request.QueryString("id")) 				
						END IF
						IF dbSet.Tables("Appraisals").Rows(0).Item("CID")=167  THEN
							redStr="manual-fulfillment-sftp.aspx?id=" &  removeInjection(Request.QueryString("id")) 				
						END IF										
						IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=2 THEN 'AND (dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")=122 OR dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")=126) THEN
	
							IF dbSet.Tables("Product").Rows.Count=0 THEN						
								redStr="manual-fulfillment-hudson.aspx?id=" &  removeInjection(Request.QueryString("id")) 
							ELSE
								IF dbSet.Tables("Product").Rows(0).Item("bpo")=1 THEN
									redStr="manual-fulfillment-hudson.aspx?id=" &  removeInjection(Request.QueryString("id")) 								
								ELSE	
									redStr="manual-fulfillment-hudson-std.aspx?id=" &  removeInjection(Request.QueryString("id")) 						
								END IF
							END IF
						END IF
						IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=4 AND dbSet.Tables("Appraisals").Rows(0).Item("APOrder")=1 THEN
							redStr="manual-fulfillment-ap.aspx?id=" &  removeInjection(Request.QueryString("id")) 
							sqlStr="SELECT * FROM products WHERE id=" & dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")
							dc.CommandText=sqlStr
							da.SelectCommand=dc
							da.Fill(dbSet,"Product2") 
							IF dbSet.Tables("Product2").Rows.Count>0 THEN
								SELECT CASE dbSet.Tables("Product2").Rows(0).Item("BPOForm")
										CASE 2
											redStr="manual-fulfillment-ap.aspx?id=" &  removeInjection(Request.QueryString("id")) 
											
										CASE 12
											redStr="manual-fulfillment-ap-fncbpo-2012.aspx?id=" &  removeInjection(Request.QueryString("id")) 																				
										
										CASE 6
											redStr="manual-fulfillment-ap-fmbpo.aspx?id=" &  removeInjection(Request.QueryString("id")) 
											
										CASE ELSE																														
											redStr="manual-fulfillment-ap-pdf.aspx?id=" &  removeInjection(Request.QueryString("id")) 									
														
								END SELECT												
								IF NOT 	dbSet.Tables("Product2").Rows(0).Item("BPOForm")=2 THEN
	'								redStr="manual-fulfillment-ap-pdf.aspx?id=" &  removeInjection(Request.QueryString("id")) 
								END IF
							END IF
						END IF
						IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=6 THEN
							redStr="manual-fulfillment-clp.aspx?id=" &  removeInjection(Request.QueryString("id")) 
						END IF
	'					Response.Write(dbSet.Tables("Appraisals").Rows(0).Item("WebService") & "|" & dbSet.Tables("Appraisals").Rows(0).Item("ClientId") & "|" & dbSet.Tables("Appraisals").Rows(0).Item("VPWSOrder"))
						IF dbSet.Tables("Appraisals").Rows(0).Item("WebService")=7 THEN
							IF dbSet.Tables("Appraisals").Rows(0).Item("ClientId")=108 THEN
								IF dbSet.Tables("Appraisals").Rows(0).Item("VPWSOrder")=1  THEN
									redStr="manual-fulfillment-vpcws-callback.aspx?id=" &  removeInjection(Request.QueryString("id")) 
								END IF
							ELSE
								redStr="manual-fulfillment-vpcws.aspx?id=" &  removeInjection(Request.QueryString("id")) 						
							END IF
						END IF	
						IF sftpFulfillment=1 THEN
							redStr="manual-fulfillment-sftp-pnmac.aspx?id=" &  removeInjection(Request.QueryString("id")) 												
						END IF
					END IF	

									
													
				ELSE
				
					sqlStr="SELECT * FROM files WHERE FRESSR=1 OR FNMSSR=1 AND appraisalid=" & 	removeInjection(Request.QueryString("id")) 
					dc.CommandText=sqlStr
					da.SelectCommand=dc
					da.Fill(dbSet,"SSRReports")
					
					sqlStr="SELECT * FROM tblCLientUADProducts WHERE clientid=" & 	dbSet.Tables("Appraisals").Rows(0).Item("ClientID")  & " AND productid=" & dbSet.Tables("Appraisals").Rows(0).Item("Product1Code")
					dc.CommandText=sqlStr
					da.SelectCommand=dc
					da.Fill(dbSet,"ClientUAD")							
					
'					IF dbSet.tables("SSRReports").Rows.COunt=0 AND dbSet.Tables("Appraisals").Rows(0).Item("UCDPClientID")>0 AND (dbSet.Tables("Product").Rows(0).Item("UADProduct")=1 OR dbSet.Tables("Appraisals").Rows(0).Item("UADOrder")=1)  AND NOT dbSet.Tables("Product").Rows(0).Item("BypassUCDP")=1 AND NOT Request.QueryString("forceF")=1 THEN
					IF ucdpClient>0 AND ( (dbSet.Tables("Product").Rows(0).Item("UADProduct")=1 AND NOT dbSet.Tables("Product").Rows(0).Item("BypassUCDP")=1) OR dbSet.Tables("Appraisals").Rows(0).Item("UADOrder")=1)  AND NOT Request.QueryString("forceFulfillment")=1 THEN
'						response.Write(ucdpClient & "|" & dbSet.Tables("Product").Rows(0).Item("UADProduct") & "|" & dbSet.Tables("Appraisals").Rows(0).Item("UADOrder") & "|" & dbSet.Tables("Product").Rows(0).Item("BypassUCDP"))
						Response.Redirect("UCDP-processing.aspx?id=" & removeInjection(Request.QueryString("id")))
					ELSE
'						Response.Write(dbSet.tables("ClientUAD").Rows.Count & "|" &dbSet.Tables("Product").Rows(0).Item("UADProduct")  & "|" & dbSet.Tables("Appraisals").Rows(0).Item("UADOrder"))

						rptAppraisals.DataSource=dbSet.Tables("Appraisals")
						rptAppraisals.DataBind()
						lblDateDue.Text=dbSet.Tables("Appraisals").Rows(0).Item("DateDue")
						lblDateComplete.Text=dbSet.Tables("Appraisals").Rows(0).Item("DateComplete")				
						lblClientID.Text=dbSet.Tables("Appraisals").Rows(0).Item("CID")			
						lblOrderNum.Text=dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber")	
						lblReturn2Order.Text="<a href=" & CHR(34) & "view_appraisals.aspx?id=" & removeInjection(Request.QueryString("id")) & CHR(34) & ">Return to Appraisal Detail</a>"	
						IF ucdpClient>0 AND (dbSet.Tables("Product").Rows(0).Item("UADProduct")=1 OR dbSet.Tables("Appraisals").Rows(0).Item("UADOrder")=1) AND NOT dbSet.Tables("Product").Rows(0).Item("BypassUCDP")=1   THEN					
							lblReturn2Order.Text=lblReturn2Order.Text & "<br /><a href=" & CHR(34) & "UCDP-Processing.aspx?id=" & removeInjection(Request.QueryString("id")) & CHR(34) & ">UCDP Processing</a>"	
						END IF						
						lblDateComplete.Text=DateTime.Parse(dbSet.Tables("Appraisals").Rows(0).Item("DateComplete")).toString()
						
						Dim additionalClientUserEmail As String=""
						Dim additionalClientUserEmail2 As String=""						
						
						IF dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserFulfillment")=1 AND NOT dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserId")=0 THEN
							sqlStr="SELECT * FROM client_users WHERE id=" & dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserId")
							dc.CommandText=sqlStr
							da.SelectCommand=dc
							da.Fill(dbSet,"AdditionalClientUser")	
							IF dbSet.Tables("AdditionalClientUser").Rows.Count>0 THEN
								additionalClientUserEmail=dbSet.Tables("AdditionalClientUser").Rows(0).Item("EMail")
							END IF					
						END IF
						
						IF dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserFulfillment2")=1 AND NOT dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserId2")=0 THEN
							sqlStr="SELECT * FROM client_users WHERE id=" & dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserId2")
							dc.CommandText=sqlStr
							da.SelectCommand=dc
							da.Fill(dbSet,"AdditionalClientUser2")	

							IF dbSet.Tables("AdditionalClientUser2").Rows.Count>0 THEN
					
								additionalClientUserEmail2=dbSet.Tables("AdditionalClientUser2").Rows(0).Item("EMail")
							END IF					
						END IF						
	
						
		'								Dim anetMail AS New MailMessage
		'
		'								anetMail.To="grm@nexus-enterprises.com"
		'
		'								anetMail.From="appraisal@valuationpartners.com"
		'								anetMail.Subject="Manual Test"
		'								anetMail.Body="grm@nexus-enterprises.com"
		'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
		'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
		'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")
		'		
		'								SmtpMail.SmtpServer = "mail.valuationpartners.com"		
		'								SmtpMail.Send(anetMail)						
		'								anetMail=Nothing					

								lblNeedinvoice.Text=dbSet.Tables("Appraisals").Rows(0).Item("ACIInvoice")
								lblRecipients.Text=makeDeliveryEmail(dbSet.Tables("Appraisals").Rows(0).Item("DeliveryEmail1"),dbSet.Tables("Appraisals").Rows(0).Item("DeliveryEmail2"),dbSet.Tables("Appraisals").Rows(0).Item("Fulfillment1"), dbSet.Tables("Appraisals").Rows(0).Item("Fulfillment2"),dbSet.Tables("Appraisals").Rows(0).Item("OrderedByEmail"),dbSet.Tables("Appraisals").Rows(0).Item("ROrderedByEmail"),dbSet.Tables("Appraisals").Rows(0).Item("RDeliveryEmail"), dbSet.Tables("Appraisals").Rows(0).Item("BorrowerEmail"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment1"),  dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment2"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment3"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment4"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillmentBorrower"),dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserFulfillment") ,additionalClientUserEmail,dbSet.Tables("Appraisals").Rows(0).Item("AdditionalClientUserFulfillment2") ,additionalClientUserEmail2, dbSet.Tables("Appraisals").Rows(0).Item("OrderedByEmail3"),  dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillmentOrderedBy3"))
	'							Response.Write(makeDeliveryEmail(dbSet.Tables("Appraisals").Rows(0).Item("DeliveryEmail1"),dbSet.Tables("Appraisals").Rows(0).Item("DeliveryEmail2"),dbSet.Tables("Appraisals").Rows(0).Item("Fulfillment1"), dbSet.Tables("Appraisals").Rows(0).Item("Fulfillment2"),dbSet.Tables("Appraisals").Rows(0).Item("OrderedByEmail"),dbSet.Tables("Appraisals").Rows(0).Item("ROrderedByEmail"),dbSet.Tables("Appraisals").Rows(0).Item("RDeliveryEmail"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment1"),  dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment2"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment3"), dbSet.Tables("Appraisals").Rows(0).Item("chkFulfillment4")) & "||")
								Dim tempAmount As Double=0	
								
								sqlStr="SELECT Payment As PaymentAmount FROM payments WHERE appraisalid=" & dbSet.Tables("Appraisals").Rows(0).Item("APID")
								dc.CommandText=sqlStr
								da.SelectCommand=dc
								da.Fill(dbSet,"Payments")
		
								
									dim paymentRow As DataRow
									DIm paymentAmount As Double=0
								FOR EACH paymentRow IN dbSet.Tables("Payments").Rows
									paymentAmount=paymentAmount + paymentRow.Item("PaymentAmount")
								
								NEXT											
								
								
								Dim appraisalid As String=dbSet.Tables("Appraisals").Rows(0).Item("APID")							
								Dim dRowTransactions As DataRow
		'						FOR EACH dRowTransactions IN dbSet.Tables("TransactionInformation").Rows
								sqlStr="SELECT * FROM MiscData WHERE OrderId=" & appraisalid & " AND Captured=0 AND OrderType=1"
								dc.CommandText=sqlStr
								da.SelectCommand=dc
								da.Fill(dbSet,"TransactionInformation")	
		
								IF dbSet.Tables("TransactionInformation").Rows.Count>0 THEN
									Try 
										IF dbSet.Tables("Appraisals").Rows.Count>0 THEN
											tempAmount=dbSet.Tables("Appraisals").Rows(0).Item("Fee")-paymentAmount
										ELSE
											tempAmount=dbSet.Tables("TransactionInformation").Rows(0).Item("TransactionAmount")
										END IF						
										IF tempAmount>dbSet.Tables("TransactionInformation").Rows(0).Item("TransactionAmount") THEN
											tempAmount=dbSet.Tables("TransactionInformation").Rows(0).Item("TransactionAmount")
										END IF
										IF tempAmount>0 THEN
											Dim authString AS String=getCapture("https://secure.authorize.net/gateway/transact.dll",tempAmount,dbSet.Tables("TransactionInformation").Rows(0).Item("AuthorizeNetID"))
					'						authString="1|1|1|1|1|1|AN2222"
											Dim authArray As String()=SPLIT(authString,"|")		
					'						Response.Write(authArray(0))					
											IF authArray(0)="1" THEN
					'							Response.Write("2")
												sqlStr="UPDATE MiscData SET Captured=1,TransactionDtCapture='" & DateTime.Now.toString() & "',AuthorizeNetIDCapture='" & authArray(6) & "' WHERE id=" &  dbSet.Tables("TransactionInformation").Rows(0).Item("ID")
												dc.CommandText=sqlStr
												dc.ExecuteNonQuery()
												
												Dim pPayment As New SQLParameter("@Payment",tempAmount)
												Dim pDt As New SQLParameter("@Dt",DateTime.Now.toString())			
												Dim tempStr100 AS String="ANET-" & authArray(6)
				
												Dim pCheckNum As New SQLParameter("@CheckNum",tempStr100)			
									
												sqlStr="INSERT INTO payments (appraisalid, Payment, Dt, CheckNum) VALUES(" & appraisalid & ",@Payment,@Dt,@CheckNum)"
												
												dc.Parameters.Add(pPayment)
												dc.Parameters.Add(pDt)
												dc.Parameters.Add(pCheckNum)			
									
															
												dc.CommandText=sqlStr
												dc.ExecuteNonQuery()								
												
					
											ELSE
												sqlStr="UPDATE MiscData SET Captured=-1 WHERE id=" &  dbSet.Tables("TransactionInformation").Rows(0).Item("ID")
												dc.CommandText=sqlStr
												dc.ExecuteNonQuery()	
						
											END IF	
										END IF	
									Catch exc2000 As Exception
										Dim anetMail AS New MailMessage
		
										anetMail.To="grm@nexus-enterprises.com"
		
										anetMail.From="appraisal@valuationpartners.com"
										anetMail.Subject="Authorize.NET Failure: " & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") 
										anetMail.Body=exc2000.toString()
										anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
										anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
										anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")
				
										SmtpMail.SmtpServer = "mail.valuationpartners.com"		
										SmtpMail.Send(anetMail)						
										anetMail=Nothing								
									Finally
									
									End Try
								END IF
		
							
							CreateInvoice1.CreateInvoice(dbSet.Tables("Appraisals").Rows(0).Item("APID"), 0)					
						
						
						IF dbSet.Tables("Appraisals").Rows(0).Item("ACIInvoice")=1 THEN
							IF File.Exists(Server.MapPath("temp/invoices/" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & ".pdf"))
								pnlInvoice.Visible="True"
								lblInvoiceLink.Text="<a href=" & CHR(34) & "temp/invoices/" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & ".pdf" & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">View Invoice</a>"
								lblInvoiceDate.Text="Last Uploaded:" & getLastDateInvoice(dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber")).toString()
								spnInvoice.innerHTML="Please review the invoice below prior to sending to fulfillment. Click the link above to create a new invoice."						
							ELSE
								pnlInvoice.Visible="False"	
								spnInvoice.innerHTML="<span style=" & CHR(34) & "font-weight:bold;color:#990000;" & CHR(34) & ">This order requires an invoice for fulfillment, and no invoice has been generated. Click the link above to generate an invoice and return to this page.</span>"	
								btnAssignAppraisal.Visible="False"			
							END IF
						
						END IF
						IF dbSet.Tables("Appraisals").Rows(0).Item("StatusCode")<100  THEN
							lblSet2100.text="1"
						ELSE
							lblSet2100.text="1"	
	'						lblSet2100.text="0"										
						END IF
						IF (dbSet.Tables("Appraisals").Rows(0).Item("StatusCode")=110 OR dbSet.Tables("Appraisals").Rows(0).Item("StatusCode") <90) AND NOT dbSet.Tables("Appraisals").Rows(0).Item("StatusCode") =70  THEN				
							btnAssignAppraisal.Visible="false"
							IF dbSet.Tables("Appraisals").Rows(0).Item("StatusCode")=110 THEN
								lblError.Text=lblError.Text & "This order has been cancelled and cannot be sent to fulfillment."
							ELSE
								lblError.Text=lblError.Text & "The order status is not 'Ready for Fulfillment' or 'Complete' and cannot be sent to manual fuilfillment."
							END IF
						END IF
							sqlStr="SELECT * FROM files WHERE aci=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
							dc.CommandText=sqlStr
							da.SelectCommand=dc
							da.Fill(dbSet,"Files")
							
							
		'					rptUsers.DataSource=dbSet.Tables("Files")
		'					rptUsers.DataBind()
							IF dbSet.Tables("Files").Rows.Count>0 THEN
		'						btnAssignAppraisal.Visible="True"
								lblACIDate.Text="Last Uploaded:" & getLastDate(dbSet.Tables("Files").Rows(0).Item("filename")).toString()
								lblFileLink.Text="<a href=" & CHR(34) & "../files/" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & "/" &dbSet.Tables("Files").Rows(0).Item("filename")  & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">" & dbSet.Tables("Files").Rows(0).Item("FileName") & "</a>"
								txtEditID.Text=dbSet.Tables("Files").Rows(0).Item("id")
								txtOldFile.Text=dbSet.Tables("Files").Rows(0).Item("filename")
							ELSE
								btnAssignAppraisal.Visible="False"						
											
							END IF		
							
							sqlStr="SELECT a.*,b.OrderNumber FROM files a, appraisals b WHERE report=1 AND b.id=a.appraisalid AND NOT a.aci=1 AND a.appraisalid=" & removeInjection(Request.QueryString("id"))
							dc.CommandText=sqlStr
							da.SelectCommand=dc
							da.Fill(dbSet,"Documents")							
		
							rptDocs.DataSource=dbSet.tables("Documents")
							rptDocs.DataBind()					
							
							
	'						IF NOT File.Exists(Server.MapPath("temp/invoices/HVCC_" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & ".pdf")) THEN
								CreateInvoice(dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyAddress"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyCity"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyState"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyZipCode"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyCounty"),DateTime.Parse(dbSet.Tables("Appraisals").Rows(0).Item("DateComplete")).toShortDateString(),dbSet.Tables("Appraisals").Rows(0).Item("Fee"),dbSet.Tables("Appraisals").Rows(0).Item("FeeSplit"))					
								
							CreateAIR(dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyAddress"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyCity"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyState"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyZipCode"),dbSet.Tables("Appraisals").Rows(0).Item("PropertyCounty"),DateTime.Parse(dbSet.Tables("Appraisals").Rows(0).Item("DateComplete")).toShortDateString(),dbSet.Tables("Appraisals").Rows(0).Item("Fee"),dbSet.Tables("Appraisals").Rows(0).Item("FeeSplit"))								
								
	'						END IF					
							lblHVCCLink.Text="<a href=" & CHR(34) & "temp/invoices/HVCC_" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & ".pdf" & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">HVCC Statement</a>"

							
							lblAIRLink.Text="<br /><a href=" & CHR(34) & "temp/invoices/AIR_" & dbSet.Tables("Appraisals").Rows(0).Item("OrderNumber") & ".pdf" & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">AIR Certificate</a>"								
						
						END IF	
				END IF
				
			Catch exc AS Exception
				IF NOT IsDbNull(Request.QueryString("tech")) THEN
					IF Request.QueryString("tech")="DIAG" THEN
						lblError.Text=exc.toString() & "<br />"
					ELSE
						lblError.Text="There has been a database error, please contact the administrator.<br />"				
					END IF
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"
				END IF
				
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing				
			End Try
			IF NOT TRIM(redStr)="" THEN
				Response.Redirect(redStr)
			END IF
		ELSE
			lblError.Text= "You have been logged out."				
		
		END IF

	END IF
	IF hdnUpdateDocs.value="1" THEN
		UpdateDocuments()
	END IF
	hdnUpdateDocs.value="0"
End Sub
Sub MarkforFulfillment(Src As Object, E As EventArgs)
		IF chkBorrowerFulfillment.Checked="true" AND NOT checkEMail(txtBorrowerEMail,0)=0 THEN
			lblErrorBorrower.Text="<br />You must provide a valid E-mail address for the borrower to send the report file."
			lblErrorBorrower.CSSClass="formred"
		ELSE
			lblErrorBorrower.CSSClass="form"	
			lblError.Text=""
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnStr As String = MyConn
			Dim dbConnection As New SqlConnection(dbConnStr)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			
			Try
				dbConnection.Open()
	'			sqlStr="UPDATE appraisals SET DateReadyForFulfillment='" & DateTime.Now.toString() & "', Status='Ready for Fulfillment', StatusCode=90  WHERE id=" & removeInjection(Request.QueryString("id"))
	'			dc.CommandText=sqlStr
	'			dc.ExecuteNonQuery()
	'			upStatus1.changeStatus(removeInjection(Request.QueryString("id")), 90,0,1, "Ready for Fulfillment",lblDateDue.Text,lblClientID.Text)								
	
	
	
				
	'			IF chkSendtoZooQueue.Checked="true" THEN
					generateXML()
	'			END IF
	'		pnlZooQueue.Visible="False"
	'			lblError.Text="Appraisal sent to Manual Fulfillment."			
				
			Catch exc AS Exception
				IF NOT IsDbNull(Request.QueryString("tech")) THEN
					IF Request.QueryString("tech")="DIAG" THEN
						lblError.Text=exc.toString() & "<br />"
					ELSE
						lblError.Text="There has been a database error, please contact the administrator.<br />"				
					END IF
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"
				END IF
				
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing				
			End Try
		END IF			
'	ELSE
'		IF chkSendtoZooQueue.Visible="True" THEN
'			lblError.Text="You must check 'Send to ZOOqueue' in in order to proceed."		
'		ELSE
'			lblError.Text="You must upload an ACI file and check 'Send to ZOOqueue' in in order to proceed."
'		END IF
'	END IF
End Sub
Sub AddEditUser(Src As Object, E As EventArgs)
	If txtEditID.Text="0" THEn
		InsertUser()
	ELSE
		ModifyUser()	
	END IF
End Sub
Function UpdateDocuments()
		Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
		Dim dbConnStr As String = MyConn
		Dim dbConnection As New SqlConnection(dbConnStr)
		Dim dc AS New SQLCommand
		Dim sqlStr AS String
		dc.Connection=dbConnection			
		Dim da AS New SQLDataAdapter()
		Dim dbSet AS New DataSet
		
		Try
			dbConnection.Open()
					sqlStr="SELECT a.*,b.OrderNumber FROM files a, appraisals b WHERE b.id=a.appraisalid AND NOT a.aci=1 AND report=1 AND a.appraisalid=" & removeInjection(Request.QueryString("id"))
					dc.CommandText=sqlStr
					da.SelectCommand=dc
					da.Fill(dbSet,"Documents")							

					rptDocs.DataSource=dbSet.tables("Documents")
					rptDocs.DataBind()		
			
		Catch exc AS Exception
			IF NOT IsDbNull(Request.QueryString("tech")) THEN
				IF Request.QueryString("tech")="DIAG" THEN
					lblError.Text=exc.toString() & "<br />"
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"				
				END IF
			ELSE
				lblError.Text="There has been a database error, please contact the administrator.<br />"
			END IF
			
		Finally
			dbConnection.Close()
			dbConnection.Dispose()		
			dbConnection=Nothing	
			dc.Dispose()		
			dc=Nothing	
			da.Dispose()		
			da=Nothing				
			dbSet.Dispose()		
			dbSet=Nothing				
		End Try
End Function
Function InsertUser()
	IF NOT checkFieldsUser() THEN
'	IF 1=2 THEN
		lblError.Text="You must select a JPG, GIF, TIF, BMP, PDF, DOC (Word) or XLS (Excel) file.<br />"
	ELSE
		Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
		Dim dbConnStr As String = MyConn
		Dim dbConnection As New SqlConnection(dbConnStr)
		Dim dc AS New SQLCommand
		Dim sqlStr AS String
		dc.Connection=dbConnection			
		Dim da AS New SQLDataAdapter()
		Dim dbSet AS New DataSet
		Dim fileName, tempStr AS String
		Dim slashPos, I, report AS Integer
		report=0
		
		Try
			fileName=File1.PostedFile.FileName
			dbConnection.Open()
			slashPos=InStrRev(fileName,"\")
			fileName=MID(fileName, slashPos +1, LEN(fileName)-slashPos)
			fileName=removeInjection(fileName)			
			tempStr=fileName			

			I=1
			WHILE File.Exists(Server.MapPath("\files\" & lblOrderNum.Text & "\" & lblOrderNum.Text & "_" & tempStr))
				tempStr=CSTR(I) & "_" & tempStr
				I=I+1
			END WHILE
			fileName=lblOrderNum.Text & "_" & tempStr
'			Response.Write(fileName)
	
			IF NOT Directory.Exists(Server.MapPath("..\files\" & lblOrderNum.Text & "\")) THEN
				Directory.CreateDirectory(Server.MapPath("..\files\" & lblOrderNum.Text & "\"))
			END IF
			File1.PostedFile.SaveAs(Server.MapPath("..\files\" & lblOrderNum.Text & "\" & fileName))		
			File1.Dispose()							
'			IF chkReport.Checked="True" THEN
'				report=1
'			END IF
			Dim pfilename As New SQLParameter("@fileName",fileName)
'			Dim preport As New SQLParameter("@report",report)			

			sqlStr="INSERT INTO files (appraisalid, ordnum, filename, report, dateuploaded, aci, permissiontype) VALUES(" & removeInjection(Request.QueryString("id")) & ",'" & lblOrderNum.Text & "',@filename, 0,'" & DateTime.Now.toString & "',1,2)"
			
			dc.Parameters.Add(pfilename)
'			dc.Parameters.Add(preport)

						
			dc.CommandText=sqlStr
			dc.ExecuteNonQuery()

'			pnlAddEdit.Visible="False"
			lblError.Text="File Uploaded"
			txtOldFile.Text=fileName
			
			sqlStr="SELECT * FROM files WHERE aci=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"Files")			
			
			IF dbSet.Tables("Files").Rows.Count>0 THEN			
				lblACIDate.Text="Last Uploaded:" & getLastDate(dbSet.Tables("Files").Rows(0).Item("filename")).toString()
						lblFileLink.Text="<a href=" & CHR(34) & "../files/" & lblOrderNum.Text & "/" &dbSet.Tables("Files").Rows(0).Item("filename")  & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">" & dbSet.Tables("Files").Rows(0).Item("FileName") & "</a>"
				
				btnAssignAppraisal.Visible="True"				
			ELSE
				btnAssignAppraisal.Visible="False"						
			END IF			
			
		Catch exc AS Exception
			IF NOT IsDbNull(Request.QueryString("tech")) THEN
				IF Request.QueryString("tech")="DIAG" THEN
					lblError.Text=exc.toString() & "<br />"
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"				
				END IF
			ELSE
				lblError.Text="There has been a database error, please contact the administrator.<br />"
			END IF
			
		Finally
			dbConnection.Close()
			dbConnection.Dispose()		
			dbConnection=Nothing	
			dc.Dispose()		
			dc=Nothing	
			da.Dispose()		
			da=Nothing				
			dbSet.Dispose()		
			dbSet=Nothing				
		End Try		
	END IF
End Function

Function ModifyUser()
	IF NOT checkFieldsUser() THEN
'	IF 1=2 THEN
		lblError.Text="You must select a JPG, GIF, TIF, BMP, PDF, DOC (Word) or XLS (Excel) file.<br />"
	ELSE
		lblError.Text=""	
		Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
		Dim dbConnStr As String = MyConn
		Dim dbConnection As New SqlConnection(dbConnStr)
		Dim dc AS New SQLCommand
		Dim sqlStr AS String
		dc.Connection=dbConnection			
		Dim da AS New SQLDataAdapter()
		Dim dbSet AS New DataSet
		Dim fileName, tempStr AS String
		Dim slashPos, I, report AS Integer
		report=0
		
		Try
			dbConnection.Open()
'			IF chkReport.Checked="True" THEN
'				report=1
'			END IF			
			
			IF NOT File1.PostedFile Is Nothing AND NOT File1.PostedFile.FileName="" THEN
				fileName=File1.PostedFile.FileName
				slashPos=InStrRev(fileName,"\")
				fileName=MID(fileName, slashPos +1, LEN(fileName)-slashPos)
				fileName=removeInjection(fileName)			
				tempStr=fileName			
				IF NOT lblOrderNum.Text & "_" & fileName=txtOldFile.Text THEN
					I=1
					WHILE File.Exists(Server.MapPath("\files\" & lblOrderNum.Text & "\" & lblOrderNum.Text & "_" & tempStr))
						tempStr=CSTR(I) & "_" & tempStr
						I=I+1
					END WHILE
				END IF
				fileName=lblOrderNum.Text & "_" & tempStr
	'			Response.Write(fileName)
		
				IF NOT Directory.Exists(Server.MapPath("..\files\" & lblOrderNum.Text & "\")) THEN
					Directory.CreateDirectory(Server.MapPath("..\files\" & lblOrderNum.Text & "\"))
				END IF
				File1.PostedFile.SaveAs(Server.MapPath("..\files\" & lblOrderNum.Text & "\"  & fileName))		
				IF NOT fileName=txtOldFile.Text THEN
					IF File.Exists(Server.MapPath("..\files\" & lblOrderNum.Text & "\" & txtOldFile.Text)) THEN
						File.Delete(Server.MapPath("..\files\" & lblOrderNum.Text & "\" & txtOldFile.Text))
					END IF
				END IF	
				File1.Dispose()	
				txtOldFile.Text=fileName						
				Dim pfilename As New SQLParameter("@fileName",fileName)
				Dim preport As New SQLParameter("@report",report)			
				sqlStr="UPDATE files SET filename=@filename  WHERE id=" & txtEditID.Text			
				dc.Parameters.Add(pfilename)
				dc.Parameters.Add(preport)	
				dc.CommandText=sqlStr
				dc.ExecuteNonQuery()							
				
				pnlZooQueue.Visible="true"
				sqlStr="SELECT * FROM files WHERE aci=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Files")
				
				
'				rptUsers.DataSource=dbSet.Tables("Files")
'				rptUsers.DataBind()
				IF dbSet.Tables("Files").Rows.Count>0 THEN
					lblACIDate.Text="Last Uploaded:" & getLastDate(dbSet.Tables("Files").Rows(0).Item("filename")).toString()
						lblFileLink.Text="<a href=" & CHR(34) & "../files/" & lblOrderNum.Text & "/" &dbSet.Tables("Files").Rows(0).Item("filename")  & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">" & dbSet.Tables("Files").Rows(0).Item("FileName") & "</a>"


					
					btnAssignAppraisal.Visible="True"				
				ELSE
					btnAssignAppraisal.Visible="False"						
				END IF				
			ELSE
				lblError.Text="You must select a JPG, GIF, TIF, BMP, PDF, DOC (Word) or XLS (Excel) file.<br />"
'				Dim preport As New SQLParameter("@report",report)				
'				sqlStr="UPDATE files SET report=@report  WHERE id=" & txtEditID.Text			
'				dc.Parameters.Add(preport)				
			END IF



			sqlStr="SELECT * FROM files WHERE aci=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"Files")
'			rptUsers.DataSource=dbSet.Tables("Files")
'			rptUsers.DataBind()
			
			IF dbSet.Tables("Files").Rows.Count>0 THEN
				lblACIDate.Text="Last Uploaded:" & getLastDate(dbSet.Tables("Files").Rows(0).Item("filename")).toString()
				lblFileLink.Text="<a href=" & CHR(34) & "../files/" & lblOrderNum.Text & "/" &dbSet.Tables("Files").Rows(0).Item("filename")  & CHR(34) & " target=" & CHR(34) & "_blank" & CHR(34) & ">" & dbSet.Tables("Files").Rows(0).Item("FileName") & "</a>"


				
				btnAssignAppraisal.Visible="True"				
			ELSE
				btnAssignAppraisal.Visible="False"						
			END IF			
			lblError.Text="File Uploaded"
			
		Catch exc AS Exception
			IF NOT IsDbNull(Request.QueryString("tech")) THEN
				IF Request.QueryString("tech")="DIAG" THEN
					lblError.Text=exc.toString() & "<br />"
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"				
				END IF
			ELSE
				lblError.Text="There has been a database error, please contact the administrator.<br />"
			END IF
			
		Finally
			dbConnection.Close()
			dbConnection.Dispose()		
			dbConnection=Nothing	
			dc.Dispose()		
			dc=Nothing	
			da.Dispose()		
			da=Nothing				
			dbSet.Dispose()		
			dbSet=Nothing				
		End Try		
	END IF
End Function

Function checkFieldsUser() As Boolean
	Dim rptItem AS RepeaterItem
	Dim errFlag AS Integer=0
	Dim ctrl AS Control
	Dim tempStr  AS String
	Dim tempVal AS Integer
	IF (File1.PostedFile Is Nothing) OR File1.PostedFile.FileName="" Then	

			errFlag=1	
			File1.Attributes("class")="formred"	

	ELSE
		IF NOT (File1.PostedFile.ContentType="image/gif" OR File1.PostedFile.ContentType="image/jpg" OR File1.PostedFile.ContentType="image/tiff" OR File1.PostedFile.ContentType="image/jpeg" OR File1.PostedFile.ContentType="image/pjpeg" OR File1.PostedFile.ContentType="image/tif" OR File1.PostedFile.ContentType="image/bmp" OR File1.PostedFile.ContentType="application/msword" OR File1.PostedFile.ContentType="application/vnd.ms-excel" OR File1.PostedFile.ContentType="application/pdf" OR File1.PostedFile.ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document" OR File1.PostedFile.ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" OR File1.PostedFile.ContentType="application/download" OR File1.PostedFile.ContentType="text/html") THEN
'			Response.Write("2 " & File1.PostedFile.ContentType)
			errFlag=1			
			File1.Attributes("class")="formred"
		ELSE
			File1.Attributes("class")="form"	
		END IF
	END IF
	IF errFlag=1 THEN
		checkFieldsUser="False"
	ELSE
		checkFieldsUser="True"
	END IF
End Function


Function removeInjection(dataVAl As String)
	dataVal=REPLACE(dataVal,CHR(60),"")
	dataVal=REPLACE(dataVal,CHR(62),"")
	dataVal=REPLACE(dataVal,CHR(39),"''")
	dataVal=REPLACE(dataVal,CHR(37),"")
	dataVal=REPLACE(dataVal,CHR(59),"")
	dataVal=REPLACE(dataVal,CHR(41),"")
	dataVal=REPLACE(dataVal,CHR(40),"")
	dataVal=REPLACE(dataVal,CHR(38),"")
	dataVal=REPLACE(dataVal,CHR(43),"")
	dataVal=REPLACE(dataVal,"-","")		
	dataVal=REPLACE(dataVal,"|","")	
	dataVal=REPLACE(dataVal,";","")					
	removeInjection=REPLACE(dataVal,CHR(34),"")																		
End Function
Function removeInjectionFileName(dataVAl As String)
	dataVal=REPLACE(dataVal,CHR(60),"")
	dataVal=REPLACE(dataVal,CHR(62),"")
	dataVal=REPLACE(dataVal,CHR(39),"''")
	dataVal=REPLACE(dataVal,CHR(37),"")
	dataVal=REPLACE(dataVal,CHR(59),"")
	dataVal=REPLACE(dataVal,CHR(41),"")
	dataVal=REPLACE(dataVal,CHR(40),"")
	dataVal=REPLACE(dataVal,CHR(38),"")
	dataVal=REPLACE(dataVal,CHR(43),"")
'	dataVal=REPLACE(dataVal,"-","")		
	dataVal=REPLACE(dataVal," ","-")		
	dataVal=REPLACE(dataVal,"|","")	
	dataVal=REPLACE(dataVal,";","")					
	removeInjectionFileName=REPLACE(dataVal,CHR(34),"")																		
End Function

	Function generateXML()
		Dim excThrown As Integer=0
		lblError.Text=""			
		Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
		Dim dbConnection As New SqlConnection(MyConn)
		Dim sqlStr As String
		sqlStr="SELECT * FROM appraisals WHERE id=" & removeInjection(Request.QueryString("id"))
		Dim dc As New SqlCommand
		Dim da AS New SQLDataAdapter()
		Dim dbSet AS New DataSet		
		dc.Connection = dbConnection
		dc.CommandText = sqlStr	
		Dim mailString As String=""
		Dim mailbody As New StringBuilder		
		
		Try
		
			dbConnection.Open()
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"OrderTransaction")
			

			sqlStr="SELECT * FROM clients WHERE id=" & dbSet.Tables("OrderTransaction").Rows(0).Item("ClientId")
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"ClientInfo")		
			
			Dim additionalClientUserEmail As String=""
			Dim additionalClientUserEmail2 As String=""			
			
			IF dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserFulfillment")=1 AND NOT dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserId")=0 THEN
				sqlStr="SELECT * FROM client_users WHERE id=" & dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserId")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"AdditionalClientUser")	
				IF dbSet.Tables("AdditionalClientUser").Rows.Count>0 THEN
					additionalClientUserEmail=dbSet.Tables("AdditionalClientUser").Rows(0).Item("EMail")
				END IF					
			END IF
			
			IF dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserFulfillment2")=1 AND NOT dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserId2")=0 THEN
				sqlStr="SELECT * FROM client_users WHERE id=" & dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserId2")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"AdditionalClientUser2")	
				IF dbSet.Tables("AdditionalClientUser2").Rows.Count>0 THEN
					additionalClientUserEmail2=dbSet.Tables("AdditionalClientUser2").Rows(0).Item("EMail")
				END IF					
			END IF			
				

	mailString=makeDeliveryEmail(dbSet.Tables("OrderTransaction").Rows(0).Item("DeliveryEmail1"),dbSet.Tables("OrderTransaction").Rows(0).Item("DeliveryEmail2"),dbSet.Tables("OrderTransaction").Rows(0).Item("Fulfillment1"), dbSet.Tables("OrderTransaction").Rows(0).Item("Fulfillment2"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderedByEmail"),dbSet.Tables("OrderTransaction").Rows(0).Item("ROrderedByEmail"),dbSet.Tables("OrderTransaction").Rows(0).Item("RDeliveryEmail"),dbSet.Tables("OrderTransaction").Rows(0).Item("BorrowerEmail"), dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillment1"),  dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillment2"), dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillment3"), dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillment4"), dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillmentBorrower"),  dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserFulfillment"),additionalClientUserEmail,dbSet.Tables("OrderTransaction").Rows(0).Item("AdditionalClientUserFulfillment2") ,additionalClientUserEmail2, dbSet.Tables("OrderTransaction").Rows(0).Item("OrderedByEmail3"),  dbSet.Tables("OrderTransaction").Rows(0).Item("chkFulfillmentOrderedBy3"))
'	Response.Write(mailString)
				IF mailString="" THEN
					enterMailError(dbSet.Tables("OrderTransaction").Rows(0))					
				ELSE
					sqlStr="SELECT * FROM files WHERE aci=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
					dc.CommandText=sqlStr
					da.SelectCommand=dc
					da.Fill(dbSet,"Files")	
					
					sqlStr="SELECT * FROM files WHERE NOT aci=1 AND report=1 AND appraisalid=" & removeInjection(Request.QueryString("id"))
					dc.CommandText=sqlStr
					da.SelectCommand=dc
					da.Fill(dbSet,"Documents")	

					
					
					Dim fileRow As DataRow
					Dim newMail AS New MailMessage
						Dim newMail2 AS New MailMessage					
					IF lblClientID.text="2" THEN

						newMail.To="app@ncslenders.com"
'						newMail.To="grm@nexus-enterprises.com;cwarr@williamfallgroup.com"
						
						newMail2.to=mailString
						newMail2.bcc="delivery@valuationpartners.com"		
						newMail2.From="appraisal@valuationpartners.com"				
		'				newMail.Add( "Reply-To", "appraisal@valuationpartners.com" );
						newMail2.Subject="Report for File # " & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & " : " & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyAddress")						
					ELSE
						newMail.To=mailString
					END IF				
					IF ConfigurationSettings.appSettings("SendEMail")="False" THEN
						newMail.To="grm@nexus-enterprises.com"
					END IF		
					newMail.bcc="delivery@valuationpartners.com"	
					IF lblClientID.text="519" THEN
						newMail.From="ls@valuationpartners.com"									
					ELSE					
						newMail.From="appraisal@valuationpartners.com"															
					END IF
	'				newMail.Add( "Reply-To", "appraisal@valuationpartners.com" );
					IF lblClientID.text="519" THEN
						newMail.Subject="Report for File # " & dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber") & " : " & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyAddress")
					ELSE
						newMail.Subject="Report for File # " & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & " : "  & addNonBlank(" ",dbSet.Tables("OrderTransaction").Rows(0).Item("ClientPONumber")," ") & addNonBlank(" ",dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber")," ") & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyAddress")
					END IF
'					newMail.Body="Valuation Partners Order Number " & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber")

					mailBody.Append("<!DOCTYPE html PUBLIC " & CHR(34) & "-//W3C//DTD XHTML 1.0 Transitional//EN" & CHR(34) & " " & CHR(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & CHR(34) & ">" & vbCrLf)
					mailBody.Append("<html xmlns=" & CHR(34) & "http://www.w3.org/1999/xhtml" & CHR(34) & ">" & vbCrLf)
					mailBody.Append("<head>" & vbCrLf)
					mailBody.Append("<meta http-equiv=" & CHR(34) & "Content-Type" & CHR(34) & " content=" & CHR(34) & "text/html; charset=utf-8" & CHR(34) & " />" & vbCrLf)
					mailBody.Append("<title>Valuation Partners</title>" & vbCrLf)
					mailBody.Append("<style type=" & CHR(34) & "text/css" & CHR(34) & ">" & vbCrLf & "<!--" & vbCrLf & "body {	margin-left: 0px;	margin-top: 0px;	margin-right: 0px;	margin-bottom: 0px;	width:680px;	margin-right:auto;	margin-left:auto;	color:#000000;	font-family: Verdana, Arial, Helvetica, sans-serif;		font-size:1.0em;}" & vbCrLf & "table#data tr td{	font-size:0.9em;	}" & vbCrLf & ".sectionhead {	background-color:#DDDDDD;	font-size:1.1em; font-weight:bold;}" & vbCrLf & "-->" & vbCrLf & "</style>" & vbCrLf)
					mailBody.Append("</head>" & vbCrLf)
					
					mailBody.Append("<body>" & vbCrLf)
								mailBody.Append("<table width=" & CHR(34) & "680" & CHR(34) & " border=" & CHR(34) & "0" & CHR(34) & " cellspacing=" & CHR(34) & "0" & CHR(34) & " cellpadding=" & CHR(34) & "2" & CHR(34) & ">" & vbCrLf & "  <tr>" & vbCrLf & "    <td bgcolor=" & CHR(34) & "#89242A" & CHR(34) & "><img src=" & CHR(34) & "http://www.valuationpartners.com/vendors/images/top_left.jpg" & CHR(34) & " width=" & CHR(34) & "171" & CHR(34) & " height=" & CHR(34) & "62" & CHR(34) & " alt=" & CHR(34) & "Valuation Partners" & CHR(34) & " border=" & CHR(34) & "0" & CHR(34) & " /></td>" & vbCrLf & "  </tr>" & vbCrLf & "  <tr>" & vbCrLf & "    <td align=" & CHR(34) & "center" & CHR(34) & "><table id=" & CHR(34) & "data" & CHR(34) & " width=" & CHR(34) & "670" & CHR(34) & " border=" & CHR(34) & "0" & CHR(34) & " cellspacing=" & CHR(34) & "0" & CHR(34) & " cellpadding=" & CHR(34) & "2" & CHR(34) & ">" & vbCrLf )


					mailBody.Append(addTableHeader("Report Details",0))												
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),"Valuation Partners Order #:") & vbCrLf)
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber"),"Client File #:") & vbCrLf)
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("ClientPONumber"),"Client Loan #:") & vbCrLf)	
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("Product1"),"Product:") & vbCrLf)										
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("BorrowerFName") & " " & dbSet.Tables("OrderTransaction").Rows(0).Item("BorrowerName"),"Borrower:") & vbCrLf)										
					mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyAddress") & "<br />" & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyCity") & ", " & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyState") & " " & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyZipCode") & "<br />" & dbSet.Tables("OrderTransaction").Rows(0).Item("PropertyCounty"),"Property Details:") & vbCrLf)
			
					mailBody.Append(addTableHeader("Attachment(s)",0))															
					
					FOR EACH fileRow IN dbSet.Tables("Files").Rows
							File.Copy(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & fileRow.Item("filename")), Server.MapPath("temp/fulfillment/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber"))),1)
	
'							newMail.Attachments.Add(new MailAttachment(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber"))))		
		
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("temp/fulfillment/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber")))))	

							mailbody.Append(addToBody(makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber")),"Report File:") & vbCrLf)	
							IF lblClientID.text="2" THEN
								newMail2.Attachments.Add(new MailAttachment(Server.MapPath("temp/fulfillment/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber")))))	
						
							END IF							
																		
					NEXT
					FOR EACH fileRow IN dbSet.Tables("Documents").Rows
							IF lblClientID.text="2" THEN
								newMail2.Attachments.Add(new MailAttachment(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & fileRow.Item("filename"))))	
			
						
							END IF
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & fileRow.Item("filename"))))	
		
							mailbody.Append(addToBody(fileRow.Item("filename"),"Report Document:") & vbCrLf)											
					NEXT	
					Dim invoicecopy As String="" 									
					IF lblNeedinvoice.Text="1" THEN
							File.Copy(Server.MapPath("temp/invoices/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf"),Server.MapPath("temp/fulfillment/Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf"),1)
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("temp/fulfillment/Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")))	
							mailbody.Append(addToBody("Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf","Invoice File:") & vbCrLf)
						invoicecopy=addToBody("Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf","Invoice File:")																							
					END IF
					IF chkSendHVCC.Checked=true THEN
						IF File.Exists(Server.MapPath("temp/invoices/HVCC_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")) THEN
								IF lblClientID.text="2" THEN
									newMail2.Attachments.Add(new MailAttachment(Server.MapPath("temp/invoices/HVCC_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")))	
								END IF						
								newMail.Attachments.Add(new MailAttachment(Server.MapPath("temp/invoices/HVCC_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")))					
								mailbody.Append(addToBody("HVCC_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf","HVCC Statement:") & vbCrLf)								
						END IF	
					END IF	
					IF File.Exists(Server.MapPath("temp/invoices/AIR_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")) THEN
							IF lblClientID.text="2" THEN
								newMail2.Attachments.Add(new MailAttachment(Server.MapPath("temp/invoices/AIR_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")))	
							END IF						
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("temp/invoices/AIR_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")))					
							mailbody.Append(addToBody("AIR_" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf","AIR Certificate of Assurance:") & vbCrLf)								
					END IF										
					IF lblSendUADXML.Text="1" THEN
						sqlStr="SELECT * FROM files WHERE UAD=1 AND appraisalid=" & dbSet.Tables("OrderTransaction").Rows(0).Item("id")
						dc.CommandText=sqlStr
						da.SelectCommand=dc
						da.Fill(dbSet,"UADXML")	
						IF dbSet.Tables("UADXML").Rows.Count>0 THEN
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & dbSet.Tables("UADXML").Rows(0).Item("filename") )))					
							mailbody.Append(addToBody(dbSet.Tables("UADXML").Rows(0).Item("filename"),"UAD XML:") & vbCrLf)														
						END IF
					END IF
					
					IF File.Exists(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "_Email-Certification.pdf")) AND  dbSet.Tables("OrderTransaction").Rows(0).Item("CLientID") = 877 THEN
			
							newMail.Attachments.Add(new MailAttachment(Server.MapPath("../files/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "/" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "_Email-Certification.pdf")))					
							mailbody.Append(addToBody(dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & "_Email-Certification.pdf","Borrower E-mail Certificate:") & vbCrLf)								
					END IF						
					
					newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
					newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
					newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")
	
					SmtpMail.SmtpServer = "mail.valuationpartners.com"
					

					mailbody.Append("  <tr><td>&nbsp;</td><td>&nbsp;</td></tr>")
					mailBody.Append("  </table></td></tr><tr><td bgcolor=" & CHR(34) & "#89242A" & CHR(34) & ">&nbsp;</td></tr></table>")					
					newMail.Body=mailBody.toString()
					newMail.BodyFormat = MailFormat.HTML						
					SmtpMail.Send(newMail)
					newMail=Nothing	
					IF lblClientID.text="2" THEN
						newMail2.Body=REPLACE(mailBody.toString(),invoicecopy,"")
						newMail2.BodyFormat = MailFormat.HTML	
						newMail2.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
						newMail2.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
						newMail2.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")											
						SmtpMail.Send(newMail2)
						newMail2=Nothing
					END IF						
'					response.Write("sent")
'								Dim anetMail AS New MailMessage
'
'								anetMail.To="grm@nexus-enterprises.com"
'
'								anetMail.From="appraisal@valuationpartners.com"
'								anetMail.Subject="Manual Test"
'								anetMail.Body=mailString
'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
'								anetMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")
'		
'								SmtpMail.SmtpServer = "mail.valuationpartners.com"		
'								SmtpMail.Send(anetMail)						
'								anetMail=Nothing					
					IF File.Exists(Server.MapPath("temp/fulfillment/Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf")) THEN
						File.Delete(Server.MapPath("temp/fulfillment/Invoice-" & dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber") & ".pdf"))
					END IF	
					FOR EACH fileRow IN dbSet.Tables("Files").Rows		
						IF File.Exists(Server.MapPath("temp/fulfillment/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber")))) THEN
							File.Delete(Server.MapPath("temp/fulfillment/" & makeReportName(fileRow.Item("filename"),dbSet.Tables("ClientInfo").Rows(0).Item("Filename"),dbSet.Tables("OrderTransaction").Rows(0).Item("OrderNumber"),dbSet.Tables("OrderTransaction").Rows(0).Item("ClientFileNumber"))))
						END IF
					
					NEXT
					
					lblErrorBorrower.Text=""
					IF chkBorrowerFulfillment.Checked="true" OR dbSet.Tables("OrderTransaction").Rows(0).Item("ClientId")=877  THEN	
						Dim bFulfillment As String="Failure"
						IF dbSet.Tables("OrderTransaction").Rows(0).Item("ClientId")=877 THEN
							bFulfillment=aMail2.changeStatus(removeInjection(Request.QueryString("id")),txtBorrowerEMail.Text)			
						ELSE
							bFulfillment=aMail1.changeStatus(removeInjection(Request.QueryString("id")),txtBorrowerEMail.Text)										
						END IF	
						
						IF bFulfillment="Success" THEN
							lblErrorBorrower.Text="<br />Report sent to borrower."
						ELSE
							lblErrorBorrower.Text="<br />Error sending report sent to borrower."									
						END IF
						IF rbUpdateDB.Checked="true" THEN
							sqlStr="UPDATE appraisals SET BorrowerEMail='" & Replace(txtBorrowerEMail.Text,"'","''") & "' WHERE id=" & removeInjection(Request.QueryString("id"))
							dc.COmmandText=sqlStr
							dc.ExecuteNonQuery()								
						END IF									
					END IF					
					
					enterZooQueueCOmment(dbSet.Tables("OrderTransaction").Rows(0),1)	
					lblError.Text="Appraisal sent to Manual Fulfillment."								
				END IF
				
'									MarkAsComplete(dbSet.Tables("OrderTransaction").Rows(0))									
			Catch exc AS Exception
				excThrown=1
'				Response.Write(exc.Message() & "||")
				enterZooQueueCOmment(dbSet.Tables("OrderTransaction").Rows(0),0)
				IF INSTR(exc.Message(),"because it is being used by another process")>0 THEN
					lblError.Text="Manual Fulfillment Failed. One of the files required for fulfillment is currently in use by another process on the server. Wait 15 - 30 seconds and try the fulfillment process again.  If the problem persists, please contact the administrator"								
				ELSE
					lblError.Text="Manual Fulfillment Failed." & exc.Message()	
'				Response.Write(exc.toString() & "||")	
				END IF
				
			Finally
				IF lblSet2100.Text="1" AND excThrown=0 THEN
					MarkAsComplete(dbSet.Tables("OrderTransaction").Rows(0))
				END IF
				
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing				
			End Try	
	
	End Function									
								
								
	Private Function enterZooQueueCOmment(datRow AS DataRow, success As Integer)
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnStr As String = MyConn
			Dim dbConnection As New SqlConnection(dbConnStr)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			Dim commentVal AS String="Manual Fulfillment Successful."
			IF success=0 THEN
				commentVal="Manual Fulfillment Failed."
			END IF
			DIm apprID As Integer
			apprID=datRow.Item("id")			
			
			Try
				dbConnection.Open()
				sqlStr="INSERT INTO appraisal_comments (LinkId, Comment, ClientView, DateAdded, CommentType, Conveyed, ConveyedMessage, ConveyedDate, NeedsConveyance, UserID, DateEntered, AddedBy,ConveyanceType) VALUES(" & apprID & ", '" & commentVal & "','No', '" & DateTime.Now.toString() & "',0,0,'','1/1/1970',0,0,'1/1/1970','" & Server.URLDecode(CStr(Request.Cookies("Name").Value)) & "','other')"	
				dc.CommandText=sqlStr					
				dc.ExecuteNonQuery()				
				
				
				IF success=1 THEN
					sqlStr="UPDATE appraisals SET Up2Zoo=1 WHERE id=" &  apprID
					dc.CommandText=sqlStr					
					dc.ExecuteNonQuery()
				ELSE
					IF datRow.Item("StatusCode")=100 OR  datRow.Item("StatusCode")=70 THEN							
						upStatus1.changeStatus(removeInjection(Request.QueryString("id")), 90,0,1, "Ready for Fulfillment",lblDateDue.Text,lblClientID.Text)
					END IF
				END IF
			
			Catch exc AS Exception
					
				
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing				
			End Try	
	
	End Function					
	
	
	Function MarkAsComplete(dRowA AS DataRow) AS String

			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnection As New SqlConnection(MyConn)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			DIm penaltyVal AS String="0"
			DIm apprID As Integer
			Dim dateFieldVal As String
			apprID=dRowA.Item("id")
			Dim commentStr AS String=""
			Dim newStatus AS Integer=100


			Try
				dbConnection.Open()
				



				DIm miscData AS String="Completed"		
				dateFieldVal="DateComplete"
				commentStr="Appraisal Completed & Fulfilled."
'				DIm comDateValue AS String=lblDateComplete.Text
'				DIm testDate As dateTime
'				testDate=DateTime.Parse(comDateValue)
'				IF testDate.Month=1 AND testDate.Year=1970 AND testDate.Day=1 THEN
'					comDateValue=DateTime.Now.toString()
'				END IF
				IF DAteTime.Parse(lblDateComplete.Text).Month=1 AND DateTime.Parse(lblDateComplete.Text).Day=1 AND DAteTime.Parse(lblDateComplete.Text).Year=1970 THEN
'					sqlStr="UPDATE appraisals SET DateComplete='" & DateTime.Now.toString() & "', Status='Completed', StatusCode=100,MReviewID=0,MReviewName='',Escalated=0, Deescalated=0, ForceIHReview=0 WHERE id=" & apprID
					upStatus1.changeStatus(apprID,100,1,1, "Completed", lblDateDue.Text,dRowA.Item("ClientId"))
				ELSE
'					sqlStr="UPDATE appraisals SET Status='Completed', StatusCode=100,MReviewID=0,MReviewName='',Escalated=0, Deescalated=0, ForceIHReview=0 WHERE id=" & apprID
					upStatus1.changeStatus(apprID,100,1,1, "Completed", "NOUPDATE",dRowA.Item("ClientId"))				
				END IF
'				sqlStr="UPDATE appraisals SET Status='" & miscData & "', StatusCode=" & newStatus & "," &  DateFieldVal & "='" & comDateValue & "',MReviewID=0,MReviewName='',Escalated=0, Deescalated=0, ForceIHReview=0 WHERE id=" & apprID

				dc.CommandText=sqlStr					
				dc.ExecuteNonQuery()				
				
	


' DETERMINE CONVEYANCE REQUIREMENTS FOR APPRAISAL RECEIVED COMMENT / INSERT COMMENT
'					Dim needConveyance As String="0"
'					sqlStr="SELECT * FROM clients WHERE id=" & dRowA.Item("ClientId") & " AND OConveyance=1 AND completed=1"	
'					dc.CommandText=sqlStr
'					da.SelectCommand=dc
'					da.Fill(dbSet,"ClientConvey")	
'					
'					IF dbSet.Tables("ClientConvey").Rows.Count>0 THEN
'						needConveyance="1"	
'					END IF								
					
'					sqlStr="INSERT INTO appraisal_comments (LinkId, Comment, ClientView, DateAdded, CommentType, Conveyed, ConveyedMessage, ConveyedDate, NeedsConveyance, UserID, DateEntered, AddedBy,ConveyanceType) VALUES(" & apprID & ", 'Appraisal Completed.','Yes', '" & DateTime.Now.toString() & "'," & needConveyance & ",0,'','1/1/1970',0,0,'1/1/1970','" & Server.URLDecode(CStr(Request.Cookies("Name").Value)) & "','completed')"	
'					dc.CommandText=sqlStr					
'					dc.ExecuteNonQuery()		
					
					



											
						
' SEND NECESSARY MESSAGES TO CLIENTS USERS
'			Dim flagVal As String="completed=1"
			
'			sqlStr="SELECT * FROM client_users,appraisals WHERE EmailNotifications='Yes' AND (appraisals.OrderedById=client_users.id OR appraisals.ROrderedById=client_users.id) AND appraisals.id=" & apprID & " AND " & flagVal
'			dc.CommandText=sqlStr
'			da.SelectCommand=dc
'			da.Fill(dbSet,"Appraisers")		

			
'			Dim dRow As DataRow
'			Dim I As Integer=0
'			FOR EACH dRow IN dbSet.Tables("Appraisers").Rows
'				Dim newMail AS New MailMessage
'				newMail.To=dRow.Item("Email")		
'				IF ConfigurationSettings.appSettings("SendEMail")="False" THEN
'					newMail.To="grm@nexus-enterprises.com"
'				END IF
'				newMail.From="appraisal@valuationpartners.com"
'				newMail.Add( "Reply-To", "appraisal@valuationpartners.com" );
'				newMail.Subject="Status Change for Valuation Partners File #:" & dRow.Item("OrderNumber") & " - " & dRow.Item("PropertyAddress")
'				newMail.Body="Status Change for Valuation Partners File #: " &  dRow.Item("OrderNumber") & vbCrLf & "Address: " & dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") &  vbCrLf & "Borrower: " & dRow.Item("BorrowerName") & addtoBody2("PO Number",dRow.Item("ClientPONumber")) & addtoBody2("File Number",dRow.Item("ClientFileNumber")) & addtoBody2("Case Number",dRow.Item("ClientCaseNumber")) & vbCrLf & "Status Details:" & vbCrLf  & commentStr
'				newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1")
'				newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "appraisal@valuationpartners.com")
'				newMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "vpt2009")

'				SmtpMail.SmtpServer = "mail.valuationpartners.com"	
'				SmtpMail.Send(newMail)						
'				newMail=Nothing		
'			NEXT													
	
		


		
			Catch exc As Exception
				
		
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing	
			End Try

		MarkASComplete="True"
	End Function				
	
	Function makeDeliveryEmail(delEM1 AS String, delEM2 AS String, ful1 AS String, ful2 AS String, OBEM AS String, ROBEM AS String, rdelEM AS String, delEMBorrower AS String, chkFul1 AS Integer, chkFul2 AS Integer, chkFul3 AS Integer, chkFul4 AS Integer, chkFulB AS Integer, chkAdditionalFulfillment As Integer, additionalClientUserEMail As String, chkAdditionalFulfillment2 As Integer, additionalClientUserEMail2 As String, obEmail3 AS String, chkFulOb AS Integer) AS String	
		makeDeliveryEmail=""
		Try
		
		Dim semiColonStr AS String=""
		IF INSTR(delEM1,"@")>0 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & delEM1
			semiColonStr=";"
		END IF
		IF INSTR(delEM2,"@")>0 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & delEM2
			semiColonStr=";"
		END IF
		IF INSTR(rdelEM,"@")>0 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & rdelEM
			semiColonStr=";"
		END IF		
		IF INSTR(ful1,"@")>0 AND chkFul1=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & ful1
			semiColonStr=";"
		END IF		
		IF INSTR(ful2,"@")>0 AND chkFul2=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & ful2
			semiColonStr=";"
		END IF			
		IF INSTR(OBEM,"@")>0 AND chkFul3=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & OBEM
			semiColonStr=";"
		END IF		
		IF INSTR(ROBEM,"@")>0 AND chkFul4=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & ROBEM
			semiColonStr=";"
		END IF
		IF INSTR(delEMBorrower,"@")>0 AND chkFulB=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & delEMBorrower
			semiColonStr=";"
		END IF	
		IF INSTR(additionalClientUserEMail,"@")>0 AND chkAdditionalFulfillment=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & additionalClientUserEMail
			semiColonStr=";"
		END IF	
		IF INSTR(additionalClientUserEMail2,"@")>0 AND chkAdditionalFulfillment2=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & additionalClientUserEMail2
			semiColonStr=";"
		END IF	
		IF INSTR(obEmail3,"@")>0 AND chkFulOb=1 THEN
			makeDeliveryEmail=makeDeliveryEmail & semiColonStr & obEmail3
			semiColonStr=";"
		END IF											
		
			Catch exc As Exception
'				Response.Write(exc.toString())	
		
			Finally

			End Try	
	
	End Function	
	Private Function enterMailError(datRow AS DataRow)
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnStr As String = MyConn
			Dim dbConnection As New SqlConnection(dbConnStr)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			Dim commentVal AS String="Fulfillment Failed - no Delivery addresses available."

			DIm apprID As Integer
			apprID=datRow.Item("id")			
			
			Try
				dbConnection.Open()
				sqlStr="INSERT INTO appraisal_comments (LinkId, Comment, ClientView, DateAdded, CommentType, Conveyed, ConveyedMessage, ConveyedDate, NeedsConveyance, UserID, DateEntered, AddedBy,ConveyanceType) VALUES(" & apprID & ", '" & commentVal & "','No', '" & DateTime.Now.toString() & "',0,0,'','1/1/1970',0,0,'1/1/1970','" & Server.URLDecode(CStr(Request.Cookies("Name").Value)) & "','other')"	
				dc.CommandText=sqlStr					
				dc.ExecuteNonQuery()				
				
			lblError.Text="No e-mail addresses for fulfillment exist in the system. Return to the appraisal detail to add fulfillment addresses."							
			
			Catch exc AS Exception
					
				
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing				
			End Try	
	
	End Function								

'	Function addToBody(dataLabel AS String, dataVal AS String)
'		IF dataVal="" THEN
'			addToBody=""
'		ELSE
'			addToBody=vbCrLf & dataLabel & ": "	& dataVal
'		END IF
'	
'	End Function		
	Function getLastDate(fileName AS String)
		dim pathVal AS String
		pathVal=Server.MapPath("../files/" & lblOrderNum.Text & "/" & fileName)
		getLastDate=File.GetLastWriteTime(pathVal)
		Dim retDate AS DateTime
		retDate=DateTime.Parse(getLastDate)
		getLastDate=retDate.toString()
	End Function	
	Function getLastDateInvoice(fileName AS String)
		dim pathVal AS String
		pathVal=Server.MapPath("temp/invoices/" & fileName & ".pdf")

		IF NOT File.Exists(pathVal) THEN
'		response.Write(pathVal)
		END IF
		getLastDateInvoice=File.GetLastWriteTime(pathVal)
'		response.Write(getLastDateInvoice)		
		Dim retDate AS DateTime
		retDate=DateTime.Parse(getLastDateInvoice)
		getLastDateInvoice=retDate.toString()
		
	End Function
	Function makeReportName(fileName As String, fName AS Integer, orderNum AS String, clientFileNumber As String)
		DIm extension AS String
		DIm lastPeriod AS Integer=INSTRREV(fileName,".")
		Dim fNameLength AS Integer=LEN(fileName)
		extension=MID(fileName,lastPeriod,fNameLength-lastPeriod+1)
		IF fName=0 THEN
			makeReportName=orderNum & extension
		ELSE
			makeReportName=removeInjectionFileName(clientFileNumber) & extension		
		END IF	
	End Function
	
	Function addNonBlank(prefix AS String, dataVal As String, suffix As String)
		IF TRIM(dataVal)="" THEN
			addNonBlank=""
		ELSE
			addNonBlank=prefix & dataval & suffix
		END IF	
	End Function
Function CreateInvoice(orderNumber AS String, propertyAddress AS String, propertyCity As String, propertyState As String, propertyZipCode As String, propertyCounty AS String, dateComplete AS String, clientFee As String, feeSplit As String)
			Dim viewPDF As Integer=1
			DIm fileName AS String=""
			lblError.Text=""	
	
	'		theDoc.Rect.String = "10 0 832 1190"
	'	    theDoc.FrameRect()
			
	'		theDoc.Rect.String = "50 50 200 950"
	'	    theDoc.FrameRect()		
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnection As New SqlConnection(MyConn)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			Dim yCoord AS Integer
			Dim dRow As DataRow
			Try
				dbConnection.Open()
				sqlStr="SELECT * FROM appraisals WHERE id=" & Request.QueryString("id")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Appraisal")	
	
				dRow=dbSet.Tables("Appraisal").rows(0)
				
				sqlStr="SELECT * FROM products WHERE id=" & drow.Item("Product1Code")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Product")	
				
				sqlStr="SELECT * FROM tblClientSpecialConfigurations WHERE clientid=" & drow.Item("ClientId") & " AND ConfigurationType='CnCHVCC'"
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"SpecialHVCC")									
									
	'
				DIm reader  as PdfReader = new PdfReader(Server.MapPath("temp/invoices/HVCC-Template-EL-2011-CnC.pdf"))
	'			DIm reader  as PdfReader = new PdfReader(Server.MapPath("temp/invoices/Invoice.pdf"))
				dIM size  as Rectangle= reader.GetPageSizeWithRotation(1)
				DIm doc AS Document = new Document(size)
				Dim  writer  As PdfWriter =  PdfWriter.getInstance(doc, new FileStream(Server.MapPath("temp/invoices/HVCC_" & orderNumber & ".pdf"), FileMode.Create, FileAccess.Write))
	'			fileName="temp/invoices/INVOICE-TEST.pdf"
				fileName="temp/invoices/HVCC_" & dRow.item("orderNumber") & ".pdf"
				doc.Open()
				
				doc.NewPage()
				DIm bf AS BaseFont = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)	
				Dim  cb As PdfContentByte = writer.DirectContent
				Dim page AS PdfImportedPage
'				IF dRow.item("ClientId")=195 OR dRow.item("RClientId")=195 THEN
				IF dbSet.Tables("SpecialHVCC").rows.Count>0 THEN
				 	page= writer.GetImportedPage(reader, 2)
					cb.AddTemplate(page, 0, 0)
					
								
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,703,9)
					
					Dim tempdate As DateTime			
					tempdate=DateTime.Parse(dRow.Item("DateComplete"))
					IF tempDate.Month=1 AND tempDate.Day=1 AND tempDate.Year=1970 THEN
						addToPDF(DateTime.Now.toShortDateString() ,bf,cb,143,685,9)				 	
					ELSE
						addToPDF(DateTime.Parse(dRow.Item("DateComplete")).toShortDateString() ,bf,cb,143,685,9)  
					END IF
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,668.5,9)					
				ELSE
					page= writer.GetImportedPage(reader, 1)
					cb.AddTemplate(page, 0, 0)
					
								
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,647,9)
					
					Dim tempdate As DateTime			
					tempdate=DateTime.Parse(dRow.Item("DateComplete"))
					IF tempDate.Month=1 AND tempDate.Day=1 AND tempDate.Year=1970 THEN
						addToPDF(DateTime.Now.toShortDateString() ,bf,cb,143,621.5,9)				 	
					ELSE
						addToPDF(DateTime.Parse(dRow.Item("DateComplete")).toShortDateString() ,bf,cb,143,621.5,9)  
					END IF
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,594.5,9)					
				END IF
										
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 3)
				cb.AddTemplate(page, 0, 0)
				
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 4)
				cb.AddTemplate(page, 0, 0)
				
				addToPDF(FormatCurrency(dRow.Item("FeeSplit"),2,-2,-2,-2) ,bf,cb,215,135,9)	
				addToPDF(FormatCurrency(dRow.Item("Fee"),2,-2,-2,-2) ,bf,cb,215,124,9)	
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 5)
				cb.AddTemplate(page, 0, 0)									
				
				IF File.Exists(Server.MapPath("../files/" & dRow.Item("OrderNumber") & "/" & dRow.Item("OrderNumber") & "-DistanceMap.jpg"))  AND dbSet.Tables("Product").rows(0).Item("AppraiserProximityMapBypass")=0 THEN
					doc.NewPage()			
					page=writer.GetImportedPage(reader, 6)
					cb.AddTemplate(page, 0, 0)	
					
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,647,9)
					
		
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,621.5,9)	
					
					addToPDF(FormatNumber(dRow.Item("DistanceToSubject"),1,-2,-2,-2) & " miles" ,bf,cb,216,541,9)										
					
					
					Dim jpg As itextsharp.text.Image = itextsharp.text.Image.getInstance(Server.MapPath("../files/" & dRow.Item("OrderNumber") & "/" & dRow.Item("OrderNumber") & "-DistanceMap.jpg"))
					jpg.setAbsolutePosition(85, 85)
					jpg.scaleToFit(440,440)
					doc.Add(jpg)
				END IF	
	
				
								
				
				doc.Close()	
				doc=Nothing	
				writer.Close()
				writer=Nothing					
		
		
	
	
	
			Catch exc As Exception
				IF NOT IsDbNull(Request.QueryString("tech")) THEN
					IF Request.QueryString("tech")="DIAG" THEN
						lblError.Text=exc.toString() & "<br />"
					ELSE
						lblError.Text="There has been a database error, please contact the administrator.<br />"				
					END IF
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"
				END IF
			
			
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing	
			End Try			
			IF viewPDF=1 THEN
	'			Response.clear()
	'			Response.Redirect(fileName)			
			END IF


		IF viewPDF=1 THEN
'			ResponseRedirect(fileName)			
		END IF
End Function
Function CreateAIR(orderNumber AS String, propertyAddress AS String, propertyCity As String, propertyState As String, propertyZipCode As String, propertyCounty AS String, dateComplete AS String, clientFee As String, feeSplit As String)
			Dim viewPDF As Integer=1
			DIm fileName AS String=""
			lblError.Text=""	
	
	'		theDoc.Rect.String = "10 0 832 1190"
	'	    theDoc.FrameRect()
			
	'		theDoc.Rect.String = "50 50 200 950"
	'	    theDoc.FrameRect()		
			Dim MyConn AS String = ConfigurationSettings.appSettings("WFGResConn")
			Dim dbConnection As New SqlConnection(MyConn)
			Dim dc AS New SQLCommand
			Dim sqlStr AS String
			dc.Connection=dbConnection			
			Dim da AS New SQLDataAdapter()
			Dim dbSet AS New DataSet
			Dim yCoord AS Integer
			Dim dRow As DataRow
			Try
				dbConnection.Open()
				sqlStr="SELECT * FROM appraisals WHERE id=" & Request.QueryString("id")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Appraisal")	
	
				dRow=dbSet.Tables("Appraisal").rows(0)
				
				sqlStr="SELECT * FROM products WHERE id=" & drow.Item("Product1Code")
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"Product")	
				
				sqlStr="SELECT * FROM tblClientSpecialConfigurations WHERE clientid=" & drow.Item("ClientId") & " AND ConfigurationType='CnCHVCC'"
				dc.CommandText=sqlStr
				da.SelectCommand=dc
				da.Fill(dbSet,"SpecialHVCC")									
									
	'
				DIm reader  as PdfReader = new PdfReader(Server.MapPath("temp/invoices/AIR-Template-v1.pdf"))
	'			DIm reader  as PdfReader = new PdfReader(Server.MapPath("temp/invoices/Invoice.pdf"))
				dIM size  as Rectangle= reader.GetPageSizeWithRotation(1)
				DIm doc AS Document = new Document(size)
				Dim  writer  As PdfWriter =  PdfWriter.getInstance(doc, new FileStream(Server.MapPath("temp/invoices/AIR_" & orderNumber & ".pdf"), FileMode.Create, FileAccess.Write))
	'			fileName="temp/invoices/INVOICE-TEST.pdf"
				fileName="temp/invoices/AIR_" & dRow.item("orderNumber") & ".pdf"
				doc.Open()
				
				doc.NewPage()
				DIm bf AS BaseFont = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)	
				Dim  cb As PdfContentByte = writer.DirectContent
				Dim page AS PdfImportedPage
'				IF dRow.item("ClientId")=195 OR dRow.item("RClientId")=195 THEN
				IF 1=2 THEN 'dbSet.Tables("SpecialHVCC").rows.Count>0 THEN
				 	page= writer.GetImportedPage(reader, 2)
					cb.AddTemplate(page, 0, 0)
					
								
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,650,9)
					
					Dim tempdate As DateTime			
					tempdate=DateTime.Parse(dRow.Item("DateComplete"))
					IF tempDate.Month=1 AND tempDate.Day=1 AND tempDate.Year=1970 THEN
						addToPDF(DateTime.Now.toShortDateString() ,bf,cb,143,685,9)				 	
					ELSE
						addToPDF(DateTime.Parse(dRow.Item("DateComplete")).toShortDateString() ,bf,cb,143,685,9)  
					END IF
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,668.5,9)					
				ELSE
					page= writer.GetImportedPage(reader, 1)
					cb.AddTemplate(page, 0, 0)
					
								
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,627,9)
					
					Dim tempdate As DateTime			
					tempdate=DateTime.Parse(dRow.Item("DateComplete"))
					IF tempDate.Month=1 AND tempDate.Day=1 AND tempDate.Year=1970 THEN
						addToPDF(DateTime.Now.toShortDateString() ,bf,cb,143,613,9)				 	
					ELSE
						addToPDF(DateTime.Parse(dRow.Item("DateComplete")).toShortDateString() ,bf,cb,143,613,9)  
					END IF
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,600,9)					
				END IF
										
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 3)
				cb.AddTemplate(page, 0, 0)
				
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 4)
				cb.AddTemplate(page, 0, 0)
				
				addToPDF(FormatCurrency(dRow.Item("FeeSplit"),2,-2,-2,-2) ,bf,cb,215,135,9)	
				addToPDF(FormatCurrency(dRow.Item("Fee"),2,-2,-2,-2) ,bf,cb,215,124,9)	
				
				doc.NewPage()			
				page=writer.GetImportedPage(reader, 5)
				cb.AddTemplate(page, 0, 0)									
				
				IF File.Exists(Server.MapPath("../files/" & dRow.Item("OrderNumber") & "/" & dRow.Item("OrderNumber") & "-DistanceMap.jpg"))  AND dbSet.Tables("Product").rows(0).Item("AppraiserProximityMapBypass")=0 THEN
					doc.NewPage()			
					page=writer.GetImportedPage(reader, 6)
					cb.AddTemplate(page, 0, 0)	
					
					addToPDF(dRow.Item("PropertyAddress") & " " & dRow.Item("PropertyCity") & ", " & dRow.Item("PropertyState") & " " & dRow.Item("PropertyZipCode") & " " & dRow.Item("PropertyCounty") ,bf,cb,155,647,9)
					
		
					
					addToPDF(dRow.Item("OrderNumber") ,bf,cb,160,621.5,9)	
					
					addToPDF(FormatNumber(dRow.Item("DistanceToSubject"),1,-2,-2,-2) & " miles" ,bf,cb,216,541,9)										
					
					
					Dim jpg As itextsharp.text.Image = itextsharp.text.Image.getInstance(Server.MapPath("../files/" & dRow.Item("OrderNumber") & "/" & dRow.Item("OrderNumber") & "-DistanceMap.jpg"))
					jpg.setAbsolutePosition(85, 85)
					jpg.scaleToFit(440,440)
					doc.Add(jpg)
				END IF	
	
				
								
				
				doc.Close()	
				doc=Nothing	
				writer.Close()
				writer=Nothing					
		
		
	
	
	
			Catch exc As Exception
				IF NOT IsDbNull(Request.QueryString("tech")) THEN
					IF Request.QueryString("tech")="DIAG" THEN
						lblError.Text=exc.toString() & "<br />"
					ELSE
						lblError.Text="There has been a database error, please contact the administrator.<br />"				
					END IF
				ELSE
					lblError.Text="There has been a database error, please contact the administrator.<br />"
				END IF
			
			
			Finally
				dbConnection.Close()
				dbConnection.Dispose()		
				dbConnection=Nothing	
				dc.Dispose()		
				dc=Nothing	
				da.Dispose()		
				da=Nothing				
				dbSet.Dispose()		
				dbSet=Nothing	
			End Try			
			IF viewPDF=1 THEN
	'			Response.clear()
	'			Response.Redirect(fileName)			
			END IF


		IF viewPDF=1 THEN
'			ResponseRedirect(fileName)			
		END IF
End Function		
Function addToBody(dataVal AS String, labelVal AS String)
	IF dataVal="" THEN
		addToBody=""
	ELSE
	    addToBody="<tr><td align=" & CHR(34) & "left" & CHR(34) & " valign=" & CHR(34) & "top" & CHR(34) & ">" & labelVal & "</td><td>" & dataVal & "</td></tr>"
	END IF

End Function
Function addToBody2(labelVal AS String,dataVal  AS String)
	IF dataVal="" THEN
		addToBody2=""
	ELSE
	    addToBody2=labelVal & ": " & dataVal 
	END IF

End Function

Function displayDate(dtString AS String, displayType As Integer)
	IF NOT dtString="" THEN
		Dim tempDt As DateTime
		tempDt=DateTime.Parse(dtString)
		IF NOT (tempDt.Month=1 ANd tempDt.Day=1 AND tempDt.Year=1970) THEN
			IF displayType=0 THEN
				displayDate=tempDt.toShortDateString()
			ELSE
				displayDate=tempDt.toString()		
			END IF
		ELSE
			displayDate=""
		END IF
	ELSE
		displayDate=""
	END IF
End Function
Function addTableHeader(message As String, firstPass AS Integer)
	IF firstPass>0 THEN
		addTableHeader="<tr><td align=" & CHR(34) & "left" & CHR(34) & " colspan=" & CHR(34) & "2" & CHR(34) & ">&nbsp;</td></tr>"	
	END IF
	addTableHeader=addTableHeader & "<tr><td align=" & CHR(34) & "left" & CHR(34) & " class=" & CHR(34) & "sectionhead" & CHR(34) & " colspan=" & CHR(34) & "2" & CHR(34) & ">" & message & "</td></tr>"
End Function	
Function setDateDue(dueDate As String)
	lblDateDue.Text=dueDate
End Function
Private Function getCapture(url As String, amountVal AS Double, transID As String) As String
   Dim result As String = ""
'   IF 1=2 THEN
	'   Dim strPost As String = "x_login=7Y2x569PhHVy&x_tran_key=9um7GK72L98uk6h7&x_method=CC&x_type=AUTH_ONLY&x_amount=" & FormatNumber(amountVal,2,0,0,0) & "&x_delim_data=TRUE&x_delim_char=|&x_relay_response=FALSE&x_card_num=" & lblccReal.Text & "&x_exp_date=" & makeExpDate("") & "&x_card_code=" & txtcusccid.Text & "&x_test_request=FALSE&x_version=3.1&x_first_name=" & Server.URLEncode(txtbillfname.Text) & "&x_last_name=" & Server.URLEncode(txtbilllname.Text) & "&x_address=" & makeAddress(txtbillstreet1.Text) & "&x_city=" & Server.URLEncode(txtbillcity.Text) & "&x_state=" & ddlbillstate.SelectedItem.Value & "&x_zip=" & Server.URLEncode(txtbillzipcode.Text) & "&x_phone=" & Server.URLEncode(txtPhone.Text) & "&x_email=" & Server.URLEncode(txtEMail.Text) & "&x_email_customer=FALSE"
'	   Dim strPost As String = "x_login=7Y2x569PhHVy&x_tran_key=9um7GK72L98uk6h7&x_method=CC&x_type=PRIOR_AUTH_CAPTURE&x_trans_id=" & transID & "&x_amount=" & FormatNumber(amountVal,2,0,0,0) & "&x_delim_data=TRUE&x_delim_char=|&x_relay_response=FALSE&x_card_num=" & lblccReal.Text & "&x_exp_date=" & makeExpDate("") & "&x_card_code=" & txtcusccid.Text & "&x_test_request=FALSE&x_version=3.1&x_first_name=" & Server.URLEncode(txtbillfname.Text) & "&x_last_name=" & Server.URLEncode(txtbilllname.Text) & "&x_address=" & makeAddress(txtbillstreet1.Text) & "&x_city=" & Server.URLEncode(txtbillcity.Text) & "&x_state=" & ddlbillstate.SelectedItem.Value & "&x_zip=" & Server.URLEncode(txtbillzipcode.Text) & "&x_email_customer=FALSE"
   Dim strPost As String = "x_login=27zE9Srv5&x_tran_key=8Eb7t73QyAT7V7dc&x_method=CC&x_type=PRIOR_AUTH_CAPTURE&x_trans_id=" & transID & "&x_amount=" & FormatNumber(amountVal,2,0,0,0) & "&x_delim_data=TRUE&x_delim_char=|&x_relay_response=FALSE&x_test_request=FALSE&x_version=3.1&x_email_customer=FALSE"	   
	   Dim myWriter As StreamWriter = Nothing
	   
	   Dim objRequest As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
	   objRequest.Method = "POST"
	   objRequest.ContentLength = strPost.Length
	   objRequest.ContentType = "application/x-www-form-urlencoded"
	   
	   Try
		  myWriter = New StreamWriter(objRequest.GetRequestStream())
		  myWriter.Write(strPost)
	   Catch e As Exception
		  Return e.Message
	   Finally
		  myWriter.Close()
	   End Try
	   
	   Dim objResponse As HttpWebResponse = CType(objRequest.GetResponse(), HttpWebResponse)
	   Dim sr As New StreamReader(objResponse.GetResponseStream())
      result = sr.ReadToEnd()
      
      ' Close and clean up the StreamReader
      sr.Close()
'	ELSE
'		result="(1|1|1|1|1|1|AN11111|1|1|1|1|1|1|1|1|1|1|)"
'	END IF	  
   Return result
End Function
Function addChangeComment(dcF AS SQLCommand, commentStr AS String, apprslID AS Integer)
	Dim addName As String
	IF Server.URLDecode(CSTR(REquest.Cookies("Name").Value))="" THEN
		addName="PlatData Processing"
	ELSE
		addName=Server.URLDecode(CSTR(REquest.Cookies("Name").Value))
	END IF
		dcF.COmmandText="INSERT INTO appraisal_comments (LinkId, Comment, ClientView, DateAdded, CommentType, Conveyed, ConveyedMessage, ConveyedDate, NeedsConveyance, UserID, DateEntered, AddedBy,ConveyanceType) VALUES(" & CSTR(apprslID) & ", '" & commentStr & "','No', '" & DateTime.Now.toString() & "',1,0,'','1/1/1970',0,0,'1/1/1970','" & addName & "','other')"
		
		dcF.ExecuteNonQuery()
End Function
	Sub addToPDF(ByVal stringValue As String, ByRef basef AS BaseFont,  ByRef contbyte AS PdfContentByte, xPos As integer, yPos As Integer, fontSize As Integer)
		contbyte.beginText()
		contbyte.setFontAndSize(basef, fontSize)
		contbyte.setTextMatrix(xPos, yPos)
		contbyte.showText(stringValue)
		contbyte.endText() 
	End Sub
	Function getRadians(degVal As Double)
		getRadians=(Math.PI / 180) * degVal
	End Function
	Function getDegrees(radVal As Double)
		getDegrees=(180 / Math.PI) * radVal
	End Function
Function checkEMail(ctrl1 AS Control, errorFlag As Integer)
	checkEMail=errorFlag
	IF CType(ctrl1,textbox).Text = "" OR LEN(CType(ctrl1,textbox).Text)=0 OR NOT INSTR(CType(ctrl1,textbox).Text," ")=0 OR INSTR(CType(ctrl1,textbox).Text,"@")=0 OR INSTR(CType(ctrl1,textbox).Text,".")=0 OR LEN(CType(ctrl1,textbox).Text) < 5 THEN
		checkEMail=1
		CType(ctrl1,textbox).CSSCLass="formred"
	ELSE
		CType(ctrl1,textbox).CSSCLass="form"	
	END IF
End Function	
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<html>
<head>
<title>Valuation Partners - Admin Area</title>
<link rel=StyleSheet href="../style.css" type="text/css" media=screen />
<link href="layout.css" rel="stylesheet" type="text/css" />
<link href="menu.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
<!--

function popUp(url,width,height) {
	if (window.featWin && !window.featWin.closed) {
		window.featWin.close();
		featWin=window.open(url,"win",'toolbar=0,location=1,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width='+width+',height='+height);
		selfWin = self;
		featWin.focus();
	}
	else {
		featWin=window.open(url,"win",'toolbar=0,location=1,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,width='+width+',height='+height);
		selfWin = self;
		featWin.focus();
	}
}
function closeWin() {
	if (window.featWin && !window.featWin.closed) {
	window.featWin.close();
	}
}
function updateDocuments() {
	document.form1.hdnUpdateDocs.value=1;
	document.form1.submit();
}
function setform() {
	document.form1.disabled=false;

}
//-->
</script>
</head>
<body text="black" alink="blue" vlink="purple" link="blue" topmargin="0" leftmargin="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="setform();">
<form name="form1" id="form1" runat="server" style="width:100%;" disabled="disabled">
<div id="header"><MenuDisplay:Menu id="menuDisplay01" runat="server" /></div>
<div id="contentholder"><div id="contentbox"><p>&nbsp;</p>
    <asp:label id="lblAdminLink" runat="server" ><a href="default.aspx">Admin Home</a>&nbsp;&bull;&nbsp;</asp:label>
    &nbsp;<a href="logout.aspx">Logout</a><br />
    &nbsp;<br />
     
<asp:label id="lblWelcome" runat="server" class="contentbold" />    
    <br />
    <strong>Send Appraisal to Manual Fulfillment</strong> - 
    <asp:label id="lblReturn2Order" runat="server" class="content" />
    <br />
    <span id="spnCreateLink" runat="server"></span><br />
     Verify the date of the Report file, or upload a new file through the interface
    provided<br />
    <span id="spnInvoice" runat="server"></span><br />
    &nbsp;<br />
    <asp:label id="lblError" runat="server" class="contentred" />
    <asp:label id="lblErrorBorrower" runat="server" class="contentred" />    
    <asp:label id="lblDateDue" runat="server" Visible="False"/>
    <asp:label id="lblClientID" runat="server" Visible="False"/> 
    <span class="form">
    <asp:label id="lblOrderNum" runat="server"  Visible="False"/>     
    <asp:label id="lblDateComplete" runat="server"  Visible="False"/>    </span>    
    <span class="form">
    <asp:label id="lblSet2100" runat="server"  Visible="False"/>    </span>
    <span class="form">
    <asp:label id="lblNeedinvoice" runat="server"  Visible="False"/>    </span>
    <span class="form">
    <asp:label id="lblSendUADXML" runat="server"  Visible="False"/>    </span>
    <span class="form">
    <asp:label id="lblSendHVCC" runat="server"  Visible="False"/>    
    </span>
    <asp:panel id="pnlForm" runat="server" >
    </asp:panel>      
    <table width="100%"  border="0" cellspacing="0" cellpadding="2">
      <tr>
        <td colspan="4" class="tableheader">Appraisal Details </td>
        </tr>
      <tr bgcolor="#CCCCCC">
        <td width="21%" class="content"><strong>Order Number </strong></td>
        <td width="28%" class="content"><strong>Client</strong></td>
        <td width="12%" class="content"><strong>Date Ordered <br />
        Date Due</strong></td>
        <td width="39%" class="contentbold">Property Details </td>
      </tr>
		<asp:repeater id="rptAppraisals" runat="server"><ITEMTEMPLATE>
      <tr valign="top">
        <td class="content"><%# Container.DataItem("OrderNumber") %></td>
        <td class="content"><%# Container.DataItem("CompanyName") %></td>
        <td class="content"><%# DateTime.Parse(Container.DataItem("DateOrdered")).toShortDateString() %><br />
          <%# DateTime.Parse(Container.DataItem("DateDue")).toShortDateString() %></td>
        <td align="left" class="content"><%# Container.DataItem("PropertyAddress") %><br /><%# Container.DataItem("PropertyAddress2") %><br /><%# Container.DataItem("PropertyCity") %>, <%# Container.DataItem("PropertyState") %> <%# Container.DataItem("PropertyZipCode") %></td>
      </tr></ITEMTEMPLATE></asp:repeater>
	  <asp:panel id="pnlZooQueue" runat="server">
	  <asp:panel id="pnlBorrowerFulfillment" runat="server">        
        <tr valign="top">
          <td class="content"><span class="style1">
            <asp:CheckBox ID="chkBorrowerFulfillment" Text="Borrower Fulfillment" runat="server" />            
          </span></td>
          <td colspan="3" class="content"><asp:TextBox ID="txtBorrowerEMail" runat="server" TextMode="SingleLine" MaxLength="75" Columns="35" />
            <asp:RadioButton ID="rbUpdateDB" Text='Update E-mail in Order' runat="server" GroupName="rg01" Checked="true" />
            &nbsp;&nbsp;
            <asp:RadioButton ID="rbNoUpdateDB" Text='Do NOT Update E-mail in Order' runat="server" GroupName="rg01" /></td>
        </tr>
        <tr valign="top">
          <td class="content">&nbsp;</td>
          <td colspan="3" class="content">&nbsp;</td>
        </tr>
        </asp:panel>
        <tr valign="top">
          <td class="content"><strong>Recipient List</strong></td>
          <td colspan="3" class="content"><asp:label id="lblRecipients" runat="server" CSSClass="form"></asp:label></td>
        </tr>
      
        <tr valign="top">
          <td class="content"><strong>Document Attachment List</strong><br />
          <a   href="javascript:popUp('file_management.aspx?id=4&fulfill=1&tech=DIAG',800,400)">Document Management<br />
          </a><br />
          (<a href="javascript:updateDocuments()">Update Document list</a>)</td>
          <td colspan="3" class="content"><asp:repeater id="rptDocs" runat="server"><itemtemplate><a target="_blank" href="../files/<%# Container.Dataitem("OrderNumber") %>/<%# Container.Dataitem("filename") %>"><%# Container.Dataitem("filename") %></a><br /></itemtemplate></asp:repeater></td>
        </tr>
    
        <tr valign="top">
          <td class="content">&nbsp;</td>
          <td colspan="3" class="content">&nbsp;</td>
        </tr>
	  <asp:panel id="pnlInvoice" runat="server">          
        <tr valign="top">
            <td class="content"><strong>Current Invoice: 
              <asp:label id="lblInvoiceLink" runat="server" CSSClass="form"></asp:label>
              <br />
              <asp:label id="lblInvoiceDate" runat="server" CSSClass="form"></asp:label>
            </strong></td>
            <td colspan="3" class="content">&nbsp;</td>
        </tr>
        </asp:panel>    
        <tr valign="top">
        <td class="content"><strong>&nbsp;<br />
          Date of Current Report File</strong><br />
          <asp:label id="lblACIDate" runat="server" CSSClass="form"></asp:label></td>
        <td colspan="3" rowspan="4" class="content"><table width="550" border="0" cellspacing="0" cellpadding="3">
            <tr>
              <td colspan="2" class="tableheader"><asp:Label ID="lblFormName" runat="server" />        
        File
          <asp:TextBox ID="txtEditID" Visible="False" Columns="35" MaxLength="50" CSSCLass="form" TextMode="SingleLine" runat="server" />  
          <asp:TextBox ID="txtOldFile" Visible="False" Columns="35" MaxLength="50" CSSCLass="form" TextMode="SingleLine" runat="server" />  
          <asp:TextBox ID="txtDateUploaded" Visible="False" Columns="35" MaxLength="50" CSSCLass="form" TextMode="SingleLine" runat="server" /></td>
            </tr>
            <tr>
              <td width="285" class="content"><strong>
                <asp:label id="lblFileLink" runat="server" CSSClass="form"></asp:label>
&nbsp;</strong></td>
              <td width="253"><input name="File1" type="file"  class="form" id="File1" runat="server" /></td>
            </tr>
            <tr align="center" bgcolor="#D4D0C8">
              <td colspan="2" class="content"><asp:button id="btnEditUser" runat="server" Text="Upload Report File" OnClick="AddEditUser" />        
                  <br />
        Do NOT submit the form more than once - it may take a minute<br />
        or more to upload your file, depending on the file size.</td>
            </tr>
        </table></td>
        </tr>
      <tr valign="top">
        <td class="content">&nbsp;</td>
        </tr>
      <tr valign="top">
        <td class="content"><strong>
          <asp:label id="lblHVCCLink" runat="server" CSSClass="form"></asp:label></strong><strong>
          <asp:CheckBox ID="chkSendHVCC" Text='Send HVCC with Report' runat="server" />          
          </strong><br /></td>
        </tr>
      <tr valign="top">
        <td class="content"><strong>
          <asp:label id="lblAIRLink" runat="server" CSSClass="form"></asp:label>
        </strong></td>
        </tr>
      <tr valign="top">
        <td class="content"><input name="hdnUpdateDocs" type="hidden" id="hdnUpdateDocs" value="0" runat="server" /></td>
        <td colspan="3" class="content">&nbsp;</td>
      </tr>	  
      <tr valign="top">
        <td colspan="4" class="content"><asp:Button ID="btnAssignAppraisal" Text="Send Appraisal to Manual Fulfillment" onClick="MarkforFulfillment" runat="server" />
          <br />
          <strong>Do NOT click the above button more than once. It may take a minute or more to send your files, depending on the size.</strong></td>
        </tr></asp:panel>
    </table>
    <p><br />
    </p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
	        <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p></div>
	  <div id="footer"><!--#include file="footer.htm" --></div></div>

</form>
<UpStatus:StatusChange id="upStatus1" runat="server" />
<CreateInvoice:CreateInvoice id="CreateInvoice1" runat="server" />
<SendAppMail:AppMail id="aMail1" runat="server" />
<SendAppMail2:AppMail2 id="aMail2" runat="server" />
</body>
</html>