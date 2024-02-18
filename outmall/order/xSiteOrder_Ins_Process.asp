<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 600 ''�ʴ���
%>
<%
'###########################################################
' Description : ���޸� API �ֹ��Է�
'###########################################################
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/shintvshopping/inc_authCheck.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/auction/auctionItemcls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%

dim IS_TEST_MODE : IS_TEST_MODE = False

Dim sqlStr, sellsite, selldate, selldateStr, mode
dim isSuccess
dim i, j, k, lp
dim orderObjArr, tmpObjArr
dim nowdate, fromdate, todate, currdate, hasMoreData, nvCount
Dim chgCode, lastOrderNo, lastTime, gubunCode
Dim isOrderComplete
Dim maxPg, xmlCheck
hasMoreData = "N"
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
mode		= requestCheckVar(html2db(request("mode")),32)
chgCode		= requestCheckVar(html2db(request("chgCode")),32)
xmlCheck	= requestCheckVar(html2db(request("xmlCheck")),1)
gubunCode	= requestCheckVar(html2db(request("gubunCode")),1)
If chgCode = "" Then
	chgCode = "PAYED"
End If

If xmlCheck <> "Y" Then
	xmlCheck = "N"
End If

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// ���ñ��� �ϰ��� ��������
	Call GetCheckStatus(sellsite, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)

	if (fromdate = todate) then
		selldateStr = fromdate
	else
		selldateStr = fromdate & " ~ " & todate
	end if
else
	fromdate = selldate
	todate = selldate
	selldateStr = fromdate
	IS_SELLDATE_FIXED = True
end if

If (sellsite = "shintvshopping" or sellsite = "skstoa") AND request("redSsnKey") = "system" Then
	IsAutoScript = True
End If

select case sellsite
	case "interpark"
		if (selldate < Left(Now(), 10)) then
			if IsAutoScript then
				response.write "interpark : " & selldate & "<br />"
			else
				response.write "interpark : " & selldate & "<br />"
				response.write "<script>alert('interpark : " & selldate & "');</script>"
			end if

			Call GetOrderFromExtSite(sellsite, selldate)
			response.write "Ȯ���ؾ��� �ֹ������� �� �ֽ��ϴ�.<br />"
		else
			fromdate = Left(DateAdd("d", -3, Now()), 10)
			todate = Left(Now, 10)

			currdate = fromdate
			do while (currdate <= todate)
				response.write "interpark : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate)
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			loop
		end if

		response.flush
	case "auction1010"
		if IsAutoScript then
			''response.write "auction1010 : " & selldateStr & "<br />"
		else
			''response.write "auction1010 : " & selldateStr & "<br />"
			response.write "<script>alert('auction1010 : " & selldateStr & "');</script>"
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "auction1010 : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "lfmall"
		if IsAutoScript then
			''response.write "auction1010 : " & selldateStr & "<br />"
		else
			''response.write "auction1010 : " & selldateStr & "<br />"
			response.write "<script>alert('lfmall : " & selldateStr & "');</script>"
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "lfmall : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "nvstorefarm", "nvstoremoonbangu", "nvstoregift", "Mylittlewhoopee"
		' IS_TEST_MODE = TRUE
		if IsAutoScript then
			''response.write "nvstorefarm : " & selldateStr & "<br />"
		else
			''response.write "nvstorefarm : " & selldateStr & "<br />"
			response.write "<script>alert('"& sellsite &" : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		'1. ���� ���̺� ����
		sqlStr = ""
		sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] WHERE sellsite = '"& sellsite &"' "
		dbget.Execute sqlStr
		do while (currdate <= todate)
			If (sellsite = "nvstoremoonbangu") Then
				response.write  sellsite & " : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate)
			ElseIf (sellsite = "Mylittlewhoopee") Then
				response.write  sellsite & " : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate)
			ElseIf (sellsite = "nvstoregift") Then
				response.write  sellsite & " : " & currdate & "<br />"
				Call GetOrderFromExtSite(sellsite, currdate)
			ElseIf (sellsite = "nvstorefarm") AND currdate <> "2020-08-09" Then
				isOrderComplete = "N"
				response.write  sellsite & " : " & currdate & "<br />"
				Do Until isOrderComplete = "Y"
					Call GetOrder_nvstorefarm(sellsite, currdate, hasMoreData, chgCode, lastOrderNo, lastTime, xmlCheck)
					If hasMoreData = "N" Then
						isOrderComplete = "Y"
					End If
					response.flush
				Loop

				Call getStorefarmOrderNumUpd()
				maxPg = getMaxPageStorefarm()

				If maxPg <> 0 Then
					For lp = 1 to maxPg
						'2. �� call�ؼ� ���� �ֹ���ȣ�� �󼼳����� �ٽ� ���Ѵ�.
						Call GetOrderFrom_NewCall_nvstorefarm(sellsite, currdate, lp, xmlCheck)
						rw "API ȣ���� �Դϴ�.............." & lp
						response.flush
					Next
				End If
			End If
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "ezwel"
		'// ���������
		'// IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "ezwel : " & selldateStr & "<br />"
		else
			''response.write "ezwel : " & selldateStr & "<br />"
			response.write "<script>alert('ezwel : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "ezwel : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "lotteCom"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "lotteCom : " & selldateStr & "<br />"
		else
			''response.write "lotteCom : " & selldateStr & "<br />"
			response.write "<script>alert('lotteCom : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "lotteCom : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			Call GetOrderFromExtSiteConfirmlist(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "wetoo1300k"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "wetoo1300k : " & selldateStr & "<br />"
		else
			''response.write "wetoo1300k : " & selldateStr & "<br />"
			response.write "<script>alert('wetoo1300k : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "wetoo1300k : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "lotteon"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "lotteon : " & selldateStr & "<br />"
		else
			''response.write "lotteon : " & selldateStr & "<br />"
			response.write "<script>alert('lotteon : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "lotteon : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "skstoa"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "skstoa : " & selldateStr & "<br />"
		else
			''response.write "skstoa : " & selldateStr & "<br />"
			response.write "<script>alert('skstoa : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -3, Now()), 10)
		end if

		'gubunCode
		'Case 1		doFlag=25
		'	�������ô�� �˻� �� tbl_xSite_TMPOrder_skstoa�� �����͸� �����Ѵ�
		'Case 2
		'	tbl_xSite_TMPOrder_shintvshopping�� ����� �����͸� ��������ó���Ѵ�
		'Case 3		doFlag=30
		'	doFlag=30�� ����� �ֹ��� ���� ������ ��ȸ�ϴ� ���̴�. �̰��� �ؾ� ������ ��ȭ��ȣ/�ּҸ� ���� �� �ִ�.
		If gubunCode = "1" Then
			sqlStr = ""
			sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] WHERE sellsite = '"& sellsite &"' "
			dbget.Execute sqlStr

			currdate = fromdate
			do while (currdate <= todate)
				response.write "skstoa : " & currdate & "<br />"
				Call GetOrderFrom_skstoa_Gubun1(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
				response.flush
			loop
			rw "======== �������ô�� ���� ó�� �� ========"
			dbget.close : response.end
		ElseIf gubunCode = "2" Then
			tmpObjArr = getSkstoaGubun1()
			If isArray(tmpObjArr) Then
				For i = 0 To UBound(tmpObjArr, 2)
					Call GetOrderFrom_skstoa_Gubun2(tmpObjArr(0, i), tmpObjArr(1, i), tmpObjArr(2, i), tmpObjArr(3, i), tmpObjArr(4, i), tmpObjArr(5, i))
				Next
				rw "======== ��ǰ �غ� �� ó�� �� ========"
				dbget.close : response.end
			Else
				rw "�������� ó���� �����Ͱ� �����ϴ�."
				response.end
			End If
		ElseIf gubunCode = "3" Then
			currdate = fromdate
			do while (currdate <= todate)
				response.write "skstoa : " & currdate & "<br />"
				Call GetOrderFrom_skstoa(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
				response.flush
			loop
		Else
			response.write "�߸��� �����Դϴ�. gubunCode���� �����ϴ�."
			dbget.close : response.end
		End If
		'http://localhost:11117/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=skstoa&gubunCode=1&selldate=2022-09-23
	case "shintvshopping"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "shintvshopping : " & selldateStr & "<br />"
		else
			''response.write "shintvshopping : " & selldateStr & "<br />"
			response.write "<script>alert('shintvshopping : " & selldateStr & "');</script>"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -2, Now()), 10)			'Left(DateAdd("d", -3, Now()), 10)���� -2�� ����
		end if

		'gubunCode
		'Case 1		doFlag=25
		'	�������ô�� �˻� �� tbl_xSite_TMPOrder_shintvshopping�� �����͸� �����Ѵ�
		'Case 2
		'	tbl_xSite_TMPOrder_shintvshopping�� ����� �����͸� ��������ó���Ѵ�
		'Case 3		doFlag=30
		'	doFlag=30�� ����� �ֹ��� ���� ������ ��ȸ�ϴ� ���̴�. �̰��� �ؾ� ������ ��ȭ��ȣ/�ּҸ� ���� �� �ִ�.
		If gubunCode = "1" Then
			sqlStr = ""
			sqlStr = sqlStr & " DELETE FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] WHERE sellsite = '"& sellsite &"' "
			dbget.Execute sqlStr

			currdate = fromdate
			do while (currdate <= todate)
				response.write "shintvshopping : " & currdate & "<br />"
				Call GetOrderFrom_shintvshopping_Gubun1(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
				response.flush
			loop
			rw "======== �������ô�� ���� ó�� �� ========"
			dbget.close : response.end
		ElseIf gubunCode = "2" Then
			tmpObjArr = getShintvshoppingGubun1()
			If isArray(tmpObjArr) Then
				For i = 0 To UBound(tmpObjArr, 2)
					Call GetOrderFrom_shintvshopping_Gubun2(tmpObjArr(0, i), tmpObjArr(1, i), tmpObjArr(2, i), tmpObjArr(3, i), tmpObjArr(4, i), tmpObjArr(5, i))
				Next
				rw "======== ��ǰ �غ� �� ó�� �� ========"
				dbget.close : response.end
			Else
				rw "�������� ó���� �����Ͱ� �����ϴ�."
				response.end
			End If
		ElseIf gubunCode = "3" Then
			currdate = fromdate
			do while (currdate <= todate)
				response.write "shintvshopping : " & currdate & "<br />"
				Call GetOrderFrom_shintvshopping(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
				response.flush
			loop
		Else
			response.write "�߸��� �����Դϴ�. gubunCode���� �����ϴ�."
			dbget.close : response.end
		End If
	case "gseshop"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "gseshop : " & selldateStr & "<br />"
		else
			''response.write "gseshop : " & selldateStr & "<br />"
			response.write "<script>alert('gseshop : " & selldateStr & "');</script>"
		end if

		'' �ּ�ó�� by eastone 2018/11/08
		' if (selldate = Left(Now(), 10)) then
		' 	fromdate = Left(DateAdd("d", -3, Now()), 10)
		' end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "gseshop : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop

		''�Ϸ羿�� ����.
		'response.write "<script>var popwin = window.open('http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=1003890&sdDt="&replace(currdate,"-","")&"&tnsType=S','popGsOrdReceiv','width=300,height=300,scrollbars=yes,resizable=yes');popwin.focus();</script>"
	case "gseshopNew"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "gseshop : " & selldateStr & "<br />"
		else
			''response.write "gseshop : " & selldateStr & "<br />"
			response.write "<script>alert('gseshop : " & selldateStr & "');</script>"
		end if

		'' �ּ�ó�� by eastone 2018/11/08
		' if (selldate = Left(Now(), 10)) then
		' 	fromdate = Left(DateAdd("d", -3, Now()), 10)
		' end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "gseshop : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop

		''�Ϸ羿�� ����.
		'response.write "<script>var popwin = window.open('http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=1003890&sdDt="&replace(currdate,"-","")&"&tnsType=S','popGsOrdReceiv','width=300,height=300,scrollbars=yes,resizable=yes');popwin.focus();</script>"
		'sellsite = "gseshop"
	case "sabangnet"
		''IS_TEST_MODE = True
		if IsAutoScript then
			''response.write "gseshop : " & selldateStr & "<br />"
		else
			''response.write "gseshop : " & selldateStr & "<br />"
			response.write "<script>alert('sabangnet : " & selldateStr & "');</script>"

			response.write "**�������<br />"
			response.write "1. �����θ��� �ֹ����� ���ݿ� ������ �Ǿ��� �־�� �մϴ�.<br />"
			response.write "2. ���θ� �ֹ������� ���� �޴� [�ֹ�����] >> [�ֹ�������(�ڵ�)] �� ������ư�� ���� �����Ͻø� �˴ϴ�.<br />"
			response.write "3. ������ �ֹ��ǿ� ���ؼ� [�ֹ�����] >> [�ֹ��� Ȯ������] �޴����� Ȯ��ó���� �ϼž� ���� ��ũ��Ʈ �������� �ֹ��� �������� �� �ֽ��ϴ�.<br />"
		end if

		if (selldate = Left(Now(), 10)) then
			fromdate = Left(DateAdd("d", -5, Now()), 10)
		end if

		currdate = fromdate
		do while (currdate <= todate)
			response.write "sabangnet : " & currdate & "<br />"
			Call GetOrderFromExtSite(sellsite, currdate)
			selldate = currdate
			currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
		loop
	case "11st1010"
		response.write "11st1010 " & selldate & "<br />"
		dbget.close : response.end
		Call GetOrderFromExtSite(sellsite, selldate)
	case "gmarket1010"
		response.write "gmarket1010 " & selldate & "<br />"
		dbget.close : response.end
		Call GetOrderFromExtSite(sellsite, selldate)
	case "coupang"
		response.write "coupang " & selldate & "<br />"
		dbget.close : response.end
		Call GetOrderFromExtSite(sellsite, selldate)
	case else
		response.write "�߸��� �����Դϴ�."
		dbget.close : response.end
end select

if (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) then
	if (selldate < Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	elseif (selldate = Left(Now(), 10)) then
		Call SetCheckStatus(sellsite, selldate, "Y")
	end if
end if

''ǰ��/���� ����üũ
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<% if  (IsAutoScript) then  %>
<% rw "OK" %>
<% else %>
<script>alert('����Ǿ����ϴ�.');</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
