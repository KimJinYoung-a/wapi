<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 180 ''초단위
'###########################################################
' Description : 제휴몰 API 정산
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/outmall/jungsan/lib/xSiteJungsanLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sellsite : sellsite	= requestCheckVar(html2db(request("sellsite")),32)
Dim reqDate : reqDate = request("reqDate")
Dim nextToken
Dim vPage, hasnext, isJungsanComplete, vTotalPage

If sellsite = "ezwel" Then
	If (reqDate = "") Then
		reqDate = Replace(Left(DateAdd("m", -1, NOW()), 7), "-", "")
	End If
	If Len(reqDate) <> 6 Then
		rw "날짜 형식이 잘 못 되었습니다."
		response.end
	End If
ElseIf sellsite = "coupang" OR sellsite = "WMP" Then
	If (reqDate = "") Then
		reqDate = DATE() - 1
	End If

	If Len(reqDate) = "8" Then
		reqDate = Left(reqDate,4) & "-" & Mid(reqDate,5,2) & "-" & mid(reqDate,7,2)
	End If

	If Len(reqDate) <> 10 Then
		rw "날짜 형식이 잘 못 되었습니다."
		response.end
	End If
Else
	response.write "잘못된 접근입니다."
	dbget.close : response.end
End If

isJungsanComplete = "N"
Select Case sellsite
	Case "ezwel"
		vPage = 1
		rw "호출월 : " & reqDate
		Do Until isJungsanComplete = "Y"
			Call GetJungsan_ezwel(reqDate, hasnext, vPage, vTotalPage)
			If hasnext = "N" Then
				isJungsanComplete = "Y"
				rw "완료 ("& vPage & "/" & vTotalPage & ")"
			Else
				rw "API 호출 중 입니다. ("& vPage & "/" & vTotalPage & ")"
				rw "-------------------------"
			End If
			response.flush
		Loop
	Case "coupang"
		rw "호출일 : " & reqDate
		Do Until isJungsanComplete = "Y"
			Call GetJungsan_coupang(reqDate, hasnext, nextToken)
			If hasnext = "N" Then
				isJungsanComplete = "Y"
				rw "완료"
			Else
				rw "API 호출 중 입니다."
				rw "-------------------------"
			End If
			response.flush
		Loop
	Case "WMP"
		rw "호출일 : " & reqDate
		rw "상품 정산 API 호출 중 입니다."
		Call GetJungsan_WMP(reqDate)
		response.flush
		rw "배송비 정산 API 호출 중 입니다."
		Call GetJungsan_WMPbeasongpay(reqDate)
		rw "완료"
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
