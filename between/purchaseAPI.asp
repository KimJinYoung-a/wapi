<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

'Response.ContentType = "application/json"
'Response.AddHeader "Accept", "application/json"
%>
<%
'#######################################################
'	Description : 비트윈 주문/결제
'	History	:  2015.05.08 한용민 생성
'	문서 : http://between-gift-gateway-dev.vcnc.co.kr/#/		' dev@10x10.co.kr		' 1q2w3e
'	태스트 경로 : https://commerce.mintnote.com/login		' giftshop1@vcnc.co.kr, giftshop2@vcnc.co.kr	' qwerty
'#######################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/between/shoppingbagDBcls.asp"-->
<!-- #include virtual="/between/betweenCommFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sData, reqValue, oResult, i, betID
Dim itemid, itemoption, itemea, requiredetail
Dim report_url, return_url, fail_url, token, usersn
Dim itemarr, itemoptionarr, itemeaarr, requiredetailarr
Dim orgreqval, oJSON
dim ref_result_str, ref_Status, ref_accessKey, ref_secretKey, ref_Key_str, vKey_confirm
	orgreqval = request("value")
	'orgreqval = "%7B%22fail_url%22%3A+%22https%3A%2F%2Fcommerce.mintnote.com%2Fpurchases%2F140%2F10x10%2Fsuccess%22%2C+%22return_url%22%3A+%22https%3A%2F%2Fcommerce.mintnote.com%2Fpurchases%2F140%2F10x10%2Fsuccess%22%2C+%22requests%22%3A+%5B%7B%22amount%22%3A+1%2C+%22text%22%3A+null%2C+%22options%22%3A+%5B%5D%2C+%22code%22%3A+%221267898%22%7D%2C+%7B%22amount%22%3A+1%2C+%22text%22%3A+null%2C+%22options%22%3A+%5B%220012%22%5D%2C+%22code%22%3A+%221267897%22%7D%5D%2C+%22token%22%3A+%22eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJwYXlsb2FkIjoie1wiaWRcIjogMTQwLCBcInVzZXJfaWRcIjogXCI3M21jRXNmMlwifSIsImV4cCI6MTQzMjAxMzUxMjMzMH0.2SFOjc90DmH4FMRcOoJOL2AYYCA5XODdPL01rg7Ccno%22%2C+%22report_url%22%3A+%22https%3A%2F%2Fcommerce.mintnote.com%2F10x10%2Fpurchases%2F140%2Freports%22%7D"

'response.write orgreqval & "<br><br>"
'response.end

if orgreqval="" then
	response.write "<script type='text/javascript'>alert('NOT INFO.');</script>"
	dbget.close()	:	response.end
end if

'/파라메타 내에 필드값을 체크해서 그부분만 가져오는 방식일경우
'SET oJSON = New aspJSON
'	oJSON.loadJSON(orgreqval)
'	reqValue = oJSON.data("value")
'set oJSON = nothing
'reqValue = FnURLDecode(reqValue)

reqValue = FnURLDecode(orgreqval)

'response.write reqValue & "<br><br>"
'response.end

Dim GoMobileURL
IF application("Svr_Info")="Dev" THEN
	GoMobileURL = "http://testm.10x10.co.kr"
Else
	GoMobileURL = "https://m.10x10.co.kr"
End If

On Error Resume Next
SET oResult = JSON.parse(reqValue)
	i = 0
	For i = 0 to oResult.requests.length-1
		If i > 0 Then
			itemid			= itemid & "||" & oResult.requests.get(i).code
			
			if oResult.requests.get(i).options <> "" then
				itemoption		= itemoption & "||" & oResult.requests.get(i).options
			else
				itemoption	= itemoption & "||" & "0000"
			end if
			
			itemea			= itemea & "||" & oResult.requests.get(i).amount
			
			if oResult.requests.get(i).text <> "" then
				requiredetail	= requiredetail & "||" & html2db(oResult.requests.get(i).text)
			else
				requiredetail	= requiredetail & "||" & "제작문구없음"
			end if
		Else
			itemid			= oResult.requests.get(i).code
			
			if oResult.requests.get(i).options <> "" then
				itemoption		= oResult.requests.get(i).options
			else
				itemoption = "0000"
			end if

			itemea			= oResult.requests.get(i).amount
			
			if oResult.requests.get(i).text <> "" then
				requiredetail	= html2db(oResult.requests.get(i).text)
			else
				requiredetail = "제작문구없음"
			end if
		End If
	Next
	report_url	= oResult.report_url
	return_url	= oResult.return_url
	fail_url	= oResult.fail_url
	token		= oResult.token

	If (Err) Then
		response.write "<script type='text/javascript'>alert('알수없는 에러가 발생 되었습니다[1].');</script>"
		dbget.close()	:	response.end
	Else
		If (itemid = "") Then
			response.write "<script type='text/javascript'>alert('상품코드가 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		ElseIf (itemea = "") Then
			response.write "<script type='text/javascript'>alert('판매수량이 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		ElseIf (report_url = "") Then
			response.write "<script type='text/javascript'>alert('통보[1] 경로가 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		ElseIf (return_url = "") Then
			response.write "<script type='text/javascript'>alert('통보[2] 경로가 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		ElseIf (fail_url = "") Then
			response.write "<script type='text/javascript'>alert('통보[3] 경로가 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		ElseIf (token = "") Then
			response.write "<script type='text/javascript'>alert('토큰 값이 누락 되었습니다.');</script>"
			dbget.close()	:	response.end
		End If
	End If
SET oResult = nothing
On Error Goto 0

'response.write itemid & " / " & itemoption & " / " & itemea & " / " & requiredetail
'response.end

''//헤더에 키값이 정상인지 체크
'vKey_confirm = getkey_confirm(ref_accessKey, ref_secretKey, ref_Key_str)
''//키값이 안맞으면 팅겨냄
'if not(vKey_confirm) then
'	response.write "<script type='text/javascript'>alert('"& ref_Key_str &"');</script>"
'	response.redirect fail_url
'	dbget.close()	:	response.end
'end if

'response.write token
'response.end

'/비트윈서버와 통신후 아이디 가져오기
betID = getBetweenID(token, ref_Status, ref_result_str)
'response.write betID & "/" & ref_Status & "/" & ref_result_str
'response.end

'/정상 통신인지 체크후 비정상이면 팅겨냄
if ref_Status<>"200" then
	response.write "<script type='text/javascript'>alert('"& ref_result_str &"');</script>"
	response.redirect fail_url
	dbget.close()	:	response.end
end if

If betID <> "" Then
	Call getTenUserSn(betID, usersn)	'가져온 비트윈 ID로 usersn가져오기
End If

Const sitename = "10x10"
Dim oShoppingBag
set oShoppingBag = new CShoppingBag
	oShoppingBag.FRectUserSn    = "BTW_USN_" & usersn
	oShoppingBag.FRectUserID    = ""
	oshoppingbag.FRectSessionID = ""
	oShoppingBag.FRectSiteName  = sitename

'	If (itemid <> "") and (itemoption <> "") and (itemea <> "") and (usersn <> "") Then	
	If (itemid <> "") and (itemoption <> "") and (itemea <> "") Then
		itemarr				= Split(itemid, "||")
		itemoptionarr		= Split(itemoption, "||")
		itemeaarr			= Split(itemea, "||")
		requiredetailarr	= Split(requiredetail, "||")

'	######################## 장바구니는 비트윈이 갖고 있고 userinfo는 우리가 갖고있음 ########################
'	1. 우리가 장바구니가 없으므로 해당 BTW_USN_usersn이 갖고 있는 데이터들 [db_my10x10].[dbo].tbl_my_baguni 에서 전부 삭제
		Call DeleteBaguniData("BTW_USN_" & usersn)
'	2. 넘어온 데이터만 인서트 
		i = 0
		For i=Lbound(itemarr) to Ubound(itemarr)
			'response.write itemarr(i) & " / " & itemoptionarr(i) & " / " & itemeaarr(i) & " / " & requiredetailarr(i)
			oshoppingbag.AddshoppingBagDB itemarr(i),itemoptionarr(i),itemeaarr(i),requiredetailarr(i)
		Next

'	3. chkOrder를 Y로
		Call UpdateChkOrderYBaguniData("BTW_USN_" & usersn)

'	4.폼 액션 Userinfo.asp
%>
			<form name="frm" method="post" action="<%= GoMobileURL %>/apps/appCom/between/inipayAPI/userinfo.asp">
			<input type="hidden" name="usersn" value="<%= usersn %>">
			<input type="hidden" name="report_url" value="<%= report_url %>">
			<input type="hidden" name="return_url" value="<%= return_url %>">
			<input type="hidden" name="fail_url" value="<%= fail_url %>">
			<input type="hidden" name="token" value="<%= token %>">
			<input type="hidden" name="signdata" value="<%= orgreqval %>">
			<input type="hidden" name="betID" value="<%= betID %>">
			</form>

			<script type="text/javascript">
				document.frm.submit();	
			</script>
<%
'	##########################################################################################################
	End If

set oShoppingBag = nothing
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->