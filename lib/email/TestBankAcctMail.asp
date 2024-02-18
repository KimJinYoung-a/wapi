<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/classes/order/bankacctcls.asp" -->
<!-- #include virtual="lib/classes/order/new_ordercls.asp"-->
<%
Dim oordermaster, oorderdetail, strPrice, i

dim fileContents, fs,dirPath,fileName,objFile, mailcontent, mailItems, tempItems, strItems

dim iorderserial : iorderserial = "11041731871"

		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.mappath("/lib/email")
		fileName = dirPath&"\\email_bank_resend.htm"
		Set objFile = fs.OpenTextFile(fileName,1)
		fileContents = objFile.readall

			If Len(iorderserial) = 11 Then

				mailcontent = fileContents

				mailItems = Mid(mailcontent,InStr(mailcontent,"<!-- 상품목록 -->"),InstrRev(mailcontent,"<!-- 상품목록 -->") + 13 - InStr(mailcontent,"<!-- 상품목록 -->") )
				mailcontent = Replace(mailcontent,mailItems,"|itemList|")

				Set objFile = Nothing
				Set fs = Nothing

				Set oordermaster = new COrderMaster
				oordermaster.FRectOrderSerial = iorderserial
				oordermaster.QuickSearchOrderMaster

				Set oorderdetail = new COrderMaster
				oorderdetail.FRectOrderSerial = iorderserial
				oorderdetail.QuickSearchOrderDetail

				strItems = ""
				For i = 0 To oorderdetail.FResultCount-1
					If oorderdetail.FItemList(i).Fitemid <> 0 Then
						tempItems = mailItems
					
						tempItems = Replace(tempItems, "|ItemID|", oorderdetail.FItemList(i).Fitemid)
						tempItems = Replace(tempItems, "|SmallImage|", oorderdetail.FItemList(i).FSmallImage)
						tempItems = Replace(tempItems, "|ItemName|", oorderdetail.FItemList(i).FItemName & "]" & oorderdetail.FItemList(i).FItemoptionName & "]")
						tempItems = Replace(tempItems, "|ItemNo|", oorderdetail.FItemList(i).FItemNo)
						
						tempItems = Replace(tempItems, "|ItemCost|", FormatNumber(oorderdetail.FItemList(i).FItemCost,0))
						tempItems = Replace(tempItems, "|ItemMileage|", FormatNumber(oorderdetail.FItemList(i).Fmileage * oorderdetail.FItemList(i).FItemNo,0))
						tempItems = Replace(tempItems, "|ItemPrice|", FormatNumber(oorderdetail.FItemList(i).FItemCost * oorderdetail.FItemList(i).FItemNo,0))

						strItems = strItems & tempItems


					End If 
				Next 

				strPrice = "상품주문금액 " &  FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) & "원&nbsp;&nbsp; + &nbsp;&nbsp;배송비 " &  FormatNumber(oorderdetail.BeasongPay,0) & "원&nbsp;&nbsp; - &nbsp;&nbsp;마일리지 " &  FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) & "원&nbsp;&nbsp; - &nbsp;&nbsp;할인 " &  FormatNumber((oordermaster.FOneItem.Ftencardspend + oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership),0) & "원&nbsp;&nbsp; = &nbsp;&nbsp;<span style=""COLOR: #c3080a; TEXT-DECORATION: none; font-weight:bold;"">" &  FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0) & "</span>원"

				mailcontent = Replace(mailcontent, "|itemList|", strItems)
				mailcontent = Replace(mailcontent, "|strPrice|", strPrice)

				mailcontent = Replace(mailcontent, "|buyName|", oordermaster.FOneItem.FBuyName)
				mailcontent = Replace(mailcontent, "|orderSerial|", oordermaster.FOneItem.ForderSerial)

				mailcontent = Replace(mailcontent, "|JumunMethodName|", oordermaster.FOneItem.JumunMethodName)
				mailcontent = Replace(mailcontent, "|RegDate|", oordermaster.FOneItem.FRegDate)
				mailcontent = Replace(mailcontent, "|miletotalprice|", FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0))
				mailcontent = Replace(mailcontent, "|tencardspend|", FormatNumber(oordermaster.FOneItem.Ftencardspend,0))
				mailcontent = Replace(mailcontent, "|etcDiscount|", FormatNumber(oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership,0))
				mailcontent = Replace(mailcontent, "|subtotalPrice|", FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0))
				mailcontent = Replace(mailcontent, "|accountno|", oordermaster.FOneItem.Faccountno)
				mailcontent = Replace(mailcontent, "|AccountName|", oordermaster.FOneItem.FAccountName)

				'rw mailcontent

				dim mailtitle
				mailtitle = "[텐바이텐] 주문에 대한 입금확인(미입금) 안내메일입니다"

				'rw oordermaster.FOneItem.Fbuyemail
				sendmail "텐바이텐<customer@10x10.co.kr>", oordermaster.FOneItem.Fbuyemail, mailtitle, mailcontent

				sqlStr = " insert into [db_temp].[dbo].tbl_bankmail_sendlist(orderserial)"
				sqlStr = sqlStr + " values('" & iorderserial & "')"
				'rw sqlStr
				dbget.execute(sqlStr)

				Set oordermaster = Nothing 
				Set oorderdetail = Nothing 

			End If 
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->