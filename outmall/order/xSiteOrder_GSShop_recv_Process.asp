<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 180
%>
<%
'###########################################################
' Description :
'###########################################################
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<%

dim IS_TEST_MODE : IS_TEST_MODE = False

Dim refIP : refIP = Request.ServerVariables ("REMOTE_ADDR")

Dim sqlStr
dim obj
Dim NowHMS
''NowHMS = Hour(Time())&Minute(Time())&Second(Time())&"0"  ''??? 이게 왜 필요한지..?
''if (NOT IS_TEST_MODE) then NowHMS="" '' ??
'response.write replace(FormatDateTime(now(),4),":","")&RIGHT("0"&second(Time()),2)&Int((Rnd * 9) + 0)
Randomize
NowHMS = replace(FormatDateTime(now(),4),":","")&RIGHT("0"&second(Time()),2)&Int((Rnd * 9) + 0)


obj = RequestArrayToArray(Request("ordNo"))

sqlStr = "insert into db_temp.dbo.tbl_tmp_gsOrder"
sqlStr = sqlStr&" (regdate,refip,xmlData)"
sqlStr = sqlStr&" values(getdate(),'"&refIP&"','" & replace(Server.HTMLEncode(Request.Form),"'","''") & "')"
dbget.Execute sqlStr

''Response.end

Call GetOrderFrom_gseshop_Recv()

if (IS_TEST_MODE = False) then
	response.write "<?xml version=""1.0"" encoding=""utf-8"" ?>"
	response.write "<PurchaseOrder_V01_00>" + vbCrLf
	response.write "<MessageHeader>" + vbCrLf
	response.write "	<Sender>10X10</Sender>" + vbCrLf
	response.write "	<Receiver>GS SHOP</Receiver>" + vbCrLf
	response.write "	<MessageID>"&Request.Form("MessageID")(1)&NowHMS&"</MessageID>" + vbCrLf
	response.write "	<DateTime>"&Request.Form("DateTime")(1)&NowHMS&"</DateTime>" + vbCrLf
	response.write "	<ProcessType>S</ProcessType>" + vbCrLf
	response.write "	<DocumentID>"&Request.Form("DocumentID")(1)&"</DocumentID>" + vbCrLf
	response.write "	<UniqueID>"&Request.Form("UniqueID")(1)&NowHMS&"</UniqueID>" + vbCrLf
	response.write "	<ErrorOccur></ErrorOccur>" + vbCrLf
	response.write "	<ErrorMessage></ErrorMessage>" + vbCrLf
	response.write "</MessageHeader>" + vbCrLf
	response.write "<MessageBody>" + vbCrLf
	response.write "	<PurchaseOrders>" + vbCrLf
	response.write "		<ordItemNo>"&Request.Form("ordItemNo")(1)&"</ordItemNo>" + vbCrLf
	response.write "		<ordNo>"&Request.Form("ordNo")(1)&"</ordNo>" + vbCrLf
	response.write "		<OrderGenerationDate>" & Left(Now(), 10) & "</OrderGenerationDate>" + vbCrLf
	response.write "		<ProductLineItem>" + vbCrLf
	response.write "			<ConfirmedDeliveryDate>" & Left(Now(), 10) & "</ConfirmedDeliveryDate>" + vbCrLf
	response.write "			<sendFg>S</sendFg>" + vbCrLf
	response.write "		</ProductLineItem>" + vbCrLf
	response.write "	</PurchaseOrders>" + vbCrLf
	response.write "</MessageBody>" + vbCrLf
	response.write "</PurchaseOrder_V01_00>" + vbCrLf
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->