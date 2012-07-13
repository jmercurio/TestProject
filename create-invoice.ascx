<%@ Control Language="VB"  Debug="True"  %>
<%@ import Namespace="System.Configuration" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data.SQLClient" %>
<%@ import Namespace="System.Text" %>
<%@ Register Assembly="itextsharp" Namespace="itextsharp" TagPrefix="itextsharp" %>
<%@ import Namespace="itextsharp.text" %>
<%@ import Namespace="itextsharp.text.pdf" %>

<script runat="server">
Sub Page_Load(Src As Object, E As EventArgs)

End Sub
Public Function CreateInvoice(appraisalId As Integer, fileLoc As Integer)

		Dim pathStarter2 As String=""
		Dim pathStarter As String=""		
		IF fileLoc=1 THEN
			pathStarter="admin/"
			pathStarter2=""			
		END IF
		Dim viewPDF As Integer=1
		DIm fileName AS String=""
'		lblError.Text=""	

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
		Try
			dbConnection.Open()
			sqlStr="SELECT * FROM appraisals WHERE id=" & Request.QueryString("id")
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"Appraisal")	
			
			
			Dim tempPayments AS Double=0
			sqlStr="SELECT SUM(Payment) AS TotPayments FROM payments WHERE appraisalid=" & CSTR(removeInjection(Request.QueryString("id")))
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"Payments")				
			IF NOT dbSet.Tables("Payments").Rows.Count=0 THEN
				IF NOT ISDbNull(dbSet.Tables("Payments").Rows(0).Item("TotPayments")) THEN
					tempPayments=dbSet.Tables("Payments").Rows(0).Item("TotPayments")
				END IF
			END IF
			sqlStr="SELECT * FROM payments WHERE appraisalid=" & CSTR(removeInjection(Request.QueryString("id")))
			dc.CommandText=sqlStr
			da.SelectCommand=dc
			da.Fill(dbSet,"Payments2")									

			dim dRow As DataRow
			Dim strPaymentDates As String
			Dim sepString As String=""
			FOR EACH DRow IN dbSet.Tables("Payments2").Rows
				strPaymentDates=strPaymentDates & sepString & DateTime.Parse(dRow.item("Dt")).toShortDateString()	
				sepString=","
			NEXT	
							
            Dim reader As PdfReader = New PdfReader(Server.MapPath(pathStarter & "temp/invoices/Invoice-VP3.pdf"))
			dIM size  as Rectangle= reader.GetPageSizeWithRotation(1)
			DIm doc AS Document = new Document(size)
			Dim  writer  As PdfWriter =  PdfWriter.getInstance(doc, new FileStream(Server.MapPath(pathStarter & "temp/invoices/" & dbSet.Tables("Appraisal").Rows(0).Item("OrderNumber") & ".pdf"), FileMode.Create, FileAccess.Write))
'			fileName="temp/invoices/INVOICE-TEST.pdf"
			fileName=pathStarter & "temp/INVOICE-TEST.pdf"
			doc.Open()
			
			doc.NewPage()
			DIm bf AS BaseFont = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED)	
			Dim  cb As PdfContentByte = writer.DirectContent
			Dim page AS PdfImportedPage = writer.GetImportedPage(reader, 1)
			cb.AddTemplate(page, 0, 0)
						
			addToPDF("Valuation Partners",bf,cb,121,972,9)			
			addToPDF("12808 West Airport Blvd., Suite 330",bf,cb,121,962,9)			
			addToPDF("Sugar Land, TX 77478",bf,cb,121,952,9)
            addToPDF("281.313.1571", bf, cb, 121, 942, 9)
			addToPDF("Federal Tax#: 26-1272692",bf,cb,121,932,9)
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientIDCode"),bf,cb,63,867,9)	
			yCoord=857
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("OrderedByName"))="" THEN	
				addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderedByName"),bf,cb,63,yCoord,9)	
				yCoord=yCoord-10				
			END IF							
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientName"),bf,cb,63,yCoord,9)	
			yCoord=yCoord-10						
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress"),bf,cb,63,yCoord,9)	
			yCoord=yCoord-10						
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress2"))="" THEN	
				addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress2"),bf,cb,63,yCoord,9)	
				yCoord=yCoord-10				
			END IF	
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientCity") & ", " & dbSet.Tables("Appraisal").Rows(0).Item("ClientState") & " " & dbSet.Tables("Appraisal").Rows(0).Item("ClientZipCode"),bf,cb,63,yCoord,9)
			yCoord=yCoord-10	
			
            yCoord = 215
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientIDCode"), bf, cb, 63, yCoord, 9)
            yCoord = yCoord - 9
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("OrderedByName"))="" THEN	
				addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderedByName"),bf,cb,63,yCoord,9)	
				yCoord=yCoord-9				
			END IF							
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientName"),bf,cb,63,yCoord,9)	
			yCoord=yCoord-9						
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress"),bf,cb,63,yCoord,9)	
			yCoord=yCoord-9						
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress2"))="" THEN	
				addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientAddress2"),bf,cb,63,yCoord,9)	
				yCoord=yCoord-9				
			END IF	
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientCity") & ", " & dbSet.Tables("Appraisal").Rows(0).Item("ClientState") & " " & dbSet.Tables("Appraisal").Rows(0).Item("ClientZipCode"),bf,cb,63,yCoord,9)
			yCoord=yCoord-9																								
			
			addToPDF("Valuation Partners",bf,cb,74,69,9)			
			addToPDF("12808 West Airport Blvd., Suite 330",bf,cb,74,59,9)			
			addToPDF("Sugar Land, TX 77478",bf,cb,74,49,9)	
			
			
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderNumber"), bf, cb, 395, 65, 9)
            addToPDF(DateTime.Now.ToString("MM/dd/yyyy"), bf, cb, 408, 54, 9)
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderNumber"), bf, cb, 403, 42, 9)
            'addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientCaseNumber"),bf,cb,389,46,9)		
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderNumber"),bf,cb,420,736,9)	
			addToPDF(DateTime.Now.toString("MM/dd/yyyy"),bf,cb,433,725,9)	
            addToPDF(DateTime.Parse(dbSet.Tables("Appraisal").Rows(0).Item("DateOrdered")).ToString("MM/dd/yyyy"), bf, cb, 428.5, 713, 9)
            addToPDF(DateTime.Now.AddDays(30).ToString("MM/dd/yyyy"), bf, cb, 457, 703, 9)
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("OrderNumber"), bf, cb, 430, 691, 9)
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientCaseNumber"), bf, cb, 414, 679.5, 9)
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientFileNumber"), bf, cb, 431, 668, 9)
            addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("ClientPONumber"), bf, cb, 408, 657, 9)

			addToPDF(FormatNumber(dbSet.Tables("Appraisal").Rows(0).Item("Fee"),2,-2,-2,-2),bf,cb,503,495,9)
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("BorrowerName"),bf,cb,104,587,9)
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("PropertyAddress") & " " & dbSet.Tables("Appraisal").Rows(0).Item("PropertyAddress2"),bf,cb,136,575,9)
			
			addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("PropertyCity")  & ", "  & dbSet.Tables("Appraisal").Rows(0).Item("PropertyState") & " " & dbSet.Tables("Appraisal").Rows(0).Item("PropertyZipCode"),bf,cb,136,563,9)																																					
				
            addToPDF("Net 30", bf, cb, 96, 351, 9)
            
'Response.Write("123")	

			
			
			Dim tempStr AS String
			tempStr=dbSet.Tables("Appraisal").Rows(0).Item("Product1")
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("Product2"))="" THEN
				tempStr=tempStr & ", " & dbSet.Tables("Appraisal").Rows(0).Item("Product2")
			END IF	
			yCoord=500		
			IF NOT TRIM(dbSet.Tables("Appraisal").Rows(0).Item("Product3"))="" THEN
				IF LEN(tempStr) + LEN(dbSet.Tables("Appraisal").Rows(0).Item("Product3"))+2>80 THEN
					addToPDF(tempStr & ",",bf,cb,64,500,9)	
					yCoord=yCoord-10									
					addToPDF(dbSet.Tables("Appraisal").Rows(0).Item("Product3"),bf,cb,64,yCoord,9)					
					yCoord=yCoord-10														
				ELSE
					tempStr=tempStr & ", " & dbSet.Tables("Appraisal").Rows(0).Item("Product3")		
					addToPDF(tempStr,bf,cb,64,yCoord,9)					
					yCoord=yCoord-10							
				END IF	
			ELSE
				addToPDF(tempStr,bf,cb,64,yCoord,9)	 			
				yCoord=yCoord-10				
			END IF	

			
'			IF NOT txtDiscount.Text="" THEN
'				addToPDF("Discount - " & Server.HTMLEncode(txtDiscountDescription.Text),bf,cb,64,yCoord,9)
''				yCoord=yCoord-12
	'			addToPDF("-" & FormatNumber(txtDiscount.Text,2,-2,-2,-2),bf,cb,503,yCoord,9)				
	'			yCoord=yCoord-12
	'		END IF
			
	
'			IF NOT TRIM(txtComment1.Text)="" THEN
'				addToPDF(Server.HTMLEncode(txtComment1.Text),bf,cb,64,yCoord,9)
'				yCoord=yCoord-10										
'			END IF	
'			IF NOT TRIM(txtComment2.Text)="" THEN
'				addToPDF(Server.HTMLEncode(txtComment2.Text),bf,cb,64,yCoord,9)
'				yCoord=yCoord-10		
'			END IF			

			Dim tempFloat As Double=dbSet.Tables("Appraisal").Rows(0).Item("Fee")
'			IF IsNumeric(txtDiscount.Text) THEN
'				tempFloat=tempFloat-CDBL(txtDiscount.Text)
'			END IF
			addToPDF(FormatNumber(tempFloat,2,-2,-2,-2),bf,cb,503,432,9)			

			addToPDF(FormatNumber(tempFloat,2,-2,-2,-2),bf,cb,503,403,9)								
			
			IF tempPayments>0 THEN
				tempFloat=tempFloat-CDBL(tempPayments)
				addToPDF(FormatNumber(tempPayments,2,-2,-2,-2),bf,cb,507,384,9)	
				addToPDF(strPaymentDates,bf,cb,323,385,9)								
			END IF		
			addToPDF(FormatNumber(tempFloat,2,-2,-2,-2),bf,cb,503,351,9)		
			

			addToPDF(FormatNumber(tempFloat,2,-2,-2,-2),bf,cb,503,242,9)								
			
			doc.Close()	
			doc=Nothing	
			writer.Close()
			writer=Nothing					
	
	



		Catch exc As Exception
			IF NOT IsDbNull(Request.QueryString("tech")) THEN
				IF Request.QueryString("tech")="DIAG" THEN
'					lblError.Text=exc.toString() & "<br />"
					rESPONSE.wRITE(EXC.TOsTRING())	
				ELSE
'					lblError.Text="There has been a database error, please contact the administrator.<br />"				
				END IF
			ELSE
'				lblError.Text="There has been a database error, please contact the administrator.<br />"
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

End Function
Function checkText(ctrl1 AS Control, errorFlag As Integer)
	checkText=errorFlag
	IF CType(ctrl1,textbox).Text="" THEN
		checkText=1
		CType(ctrl1,textbox).CSSCLass="formred"
	ELSE
		CType(ctrl1,textbox).CSSCLass="form"	
	END IF
End Function
Function checkhtmlinputtextDate(ctrl1 AS Control, errorFlag As Integer)
	checkhtmlinputtextDate=errorFlag
	IF CType(ctrl1,htmlinputtext).value="" AND NOT IsDate(CType(ctrl1,htmlinputtext).value) THEN
		checkhtmlinputtextDate=1
		CType(ctrl1,htmlinputtext).Attributes("class")="formred"
	ELSE
		CType(ctrl1,htmlinputtext).Attributes("class")="form"	
	END IF
End Function
Function checkInteger(ctrl1 AS Control, errorFlag As Integer)
	checkInteger=errorFlag
	IF CType(ctrl1,textbox).Text="" THEN
'		checkInteger=
'		CType(ctrl1,textbox).CSSCLass="formred"
	ELSE
		IF NOT IsNumeric(CType(ctrl1,textbox).Text) THEN
			checkInteger=1
			CType(ctrl1,textbox).CSSCLass="formred"
		ELSE
			CType(ctrl1,textbox).CSSCLass="form"	
		END IF	
	END IF
End Function

Function clearField(fieldType As String, controlID AS Control)
	SELECT fieldType
		CASE "T"
			CType(controlID,textbox).Text=""
			CType(controlID,textbox).CSSCLass="form"
		CASE "H"
			CType(controlID,htmlinputtext).Value=""
			CType(controlID,htmlinputtext).attributes("class")="form"

		CASE "F"
'			File1.Dispose()
'			File1.Attributes("class")="form"
'			CType(controlID,htmlfileinput).FileName=""
'			CType(controlID,htmlfileinput).Attributes("class")="form"			
		CASE "P"	
			CType(controlID,textbox).attributes.add("value", "")
			CType(controlID,textbox).CSSClass="form"			
		CASE "C"
			CType(controlID,checkbox).Checked="False"
			CType(controlID,checkbox).CSSClass="form"			
		CASE "D"
			Dim lItem AS System.Web.UI.WebControls.ListItem
			FOR EACH lItem IN CType(controlID,dropdownlist).Items
				lItem.Selected="False"
			NEXT
			CType(controlID,dropdownlist).SelectedIndex=0
			CType(controlID,dropdownlist).CSSClass="form"					
	END SELECT		
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
	dataVal=REPLACE(dataVal," ","_")							
	removeInjection=REPLACE(dataVal,CHR(34),"")																		
End Function
Function makeLink(fName As String)
'	makeLink="<a target=" & CHR(34) & "_blank" & CHR(34) & " href=" & CHR(34) & "\files\" & lblOrderNum.Text & "\" & fName &   CHR(34) & ">View</a>"
End Function
Function makeReport(repVal As Integer)
	IF repVal=1 THEN
		makeReport="X"
	ELSE
		makeReport=""	
	END IF
End Function
	Function deductPayment(paymentVal As Double)
'		tdGrandTotal.Value=CDBL(tdGrandTotal.Value)	- paymentVal
	End Function
	Sub addToPDF(ByVal stringValue As String, ByRef basef AS BaseFont,  ByRef contbyte AS PdfContentByte, xPos As integer, yPos As Integer, fontSize As Integer)
		contbyte.beginText()
		contbyte.setFontAndSize(basef, fontSize)
		contbyte.setTextMatrix(xPos, yPos)
		contbyte.showText(stringValue)
		contbyte.endText() 
	End Sub		
</script>
