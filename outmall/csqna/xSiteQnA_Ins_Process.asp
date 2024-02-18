<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 180 ''초단위
'###########################################################
' Description : 제휴몰 API CS Q&A
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/shintvshopping/inc_authCheck.asp"-->
<!-- #include virtual="/outmall/skstoa/skstoaItemcls.asp"-->
<!-- #include virtual="/outmall/csqna/lib/xSiteQnALib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim IS_TEST_MODE : IS_TEST_MODE = False
Dim sellsite, selldate, csGubun, isSuccess, arrRows, arrRows2
Dim i, j, k
Dim nowdate, fromdate, todate, currdate
Dim sqlStr, msg

sellsite	= requestCheckVar(html2db(request("sellsite")),32)
selldate	= requestCheckVar(html2db(request("selldate")),32)
csGubun		= requestCheckVar(html2db(request("mode")),32)

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")
dim redKey  : redKey = requestCheckVar(request("redSsnKey"),32)

If (sellsite = "") or (csGubun = "") then
	response.write "잘못된 접근입니다."
	dbget.close : response.end
end if

dim IS_SELLDATE_FIXED : IS_SELLDATE_FIXED = False
if (selldate = "") then
	'// 오늘까지 일괄로 가져오기
	Call GetCSCheckStatus(sellsite, csGubun, selldate, isSuccess)
	fromdate = selldate
	todate = Left(Now, 10)
else
	fromdate = selldate
	todate = selldate
	IS_SELLDATE_FIXED = True
end if

Select Case sellsite
	Case "nvstorefarm", "nvstoremoonbangu", "nvstoregift", "Mylittlewhoopee"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate

			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_nvstorefarm(currdate, sellsite)
				Call GetCSOrderQnA_nvstorefarm(currdate, sellsite)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getNvstorefarmCSAnswerComplete(sellsite, "item")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_nvstorefarm(arrRows(0, i), arrRows(1, i), sellsite)
				Next
			Else
				rw "None(item)"
			End If

			arrRows2 = getNvstorefarmCSAnswerComplete(sellsite, "order")
			If IsArray(arrRows2) Then
				For i = 0 To Ubound(arrRows2, 2)
					Call resCSOrderQnA_nvstorefarm(arrRows2(0, i), arrRows2(1, i), sellsite)
				Next
			Else
				rw "None(order)"
			End If
		Else
			response.write "잘못된 접근입니다.."
			dbget.close : response.end
		End If
	Case "sabangnet"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate
			Do While (currdate <= todate)
'				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_sabangnet(currdate)

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		End If
	Case "lotteCom"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate
			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_lotteCom(currdate)

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_lotteCom(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "lotteimall"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate
			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_lotteimall(currdate)

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_lotteimall(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "shintvshopping"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate

			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_shintvshopping(currdate)
				Call GetCSQnA_shintvshopping_complete(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getShintvshoppingCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_shintvshopping(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case "skstoa"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate

			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_skstoa(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getSkstoaCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_skstoa(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
		'http://localhost:11117/outmall/csqna/xSiteQnA_Ins_Process.asp?sellsite=skstoa&mode=resQnA
	Case "11st1010"
		Dim sugi
		sugi = requestCheckVar(html2db(request("v")),1)
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate
			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_11st1010(currdate, sugi)

				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = get11stCSAnswerComplete(sellsite)
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_11st1010(arrRows(0, i), arrRows(1, i), arrRows(2, i))
				Next
			Else
				rw "None"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
		sqlStr = ""
		sqlStr = sqlStr & " EXEC [db_item].[dbo].[sp_Ten_MyQna_Get_ExtQna] "
		dbget.Execute sqlStr
	Case "lotteon"
		If (csGubun = "reqQnA") Then
			If (selldate = Left(Now(), 10)) Then
				fromdate = Left(DateAdd("d", -1, Now()), 10)
			End If
			currdate = fromdate

			Do While (currdate <= todate)
				response.write "<br />" & sellsite & " : " & currdate & "<br />"
				Call GetCSQnA_lotteon(currdate)
				Call GetCSSellerQnA_lotteon(currdate)
				selldate = currdate
				currdate = Left(DateAdd("d", 1, CDate(currdate)), 10)
			Loop
		ElseIf (csGubun = "resQnA") Then
			arrRows = getLotteonCSAnswerComplete(sellsite, "item")
			If IsArray(arrRows) Then
				For i = 0 To Ubound(arrRows, 2)
					Call resCSQnA_lotteon(arrRows(0, i), arrRows(1, i))
				Next
			Else
				rw "None(item)"
			End If

			arrRows2 = getLotteonCSAnswerComplete(sellsite, "seller")
			If IsArray(arrRows2) Then
				For i = 0 To Ubound(arrRows2, 2)
					Call resCSSellerQnA_lotteon(arrRows2(0, i), arrRows2(1, i))
				Next
			Else
				rw "None(seller)"
			End If
		Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
		End If
	Case Else
		response.write "잘못된 접근입니다.s"
		dbget.close : response.end
End Select

' 2022-04-18 김진영 주석..중복등록된다고 함;
' sqlStr = ""
' sqlStr = sqlStr & " EXEC [db_item].[dbo].[sp_Ten_MyQna_Get_ExtQna] "
' dbget.Execute sqlStr

If (IS_TEST_MODE = False) and (IS_SELLDATE_FIXED = False) Then
	If (selldate < Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, Left(DateAdd("d", 1, CDate(selldate)), 10), "N")
	ElseIf (selldate = Left(Now(), 10)) Then
		Call SetCSCheckStatus(sellsite, csGubun, selldate, "Y")
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
