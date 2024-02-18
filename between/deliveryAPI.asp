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
'	Description : 비트윈 주문/배송
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
Dim sData, reqtoken, oResult, i, betID
Dim itemid, itemoption, itemea, requiredetail
Dim report_url, return_url, fail_url, token, usersn
Dim itemarr, itemoptionarr, itemeaarr, requiredetailarr
Dim orgreqval, oJSON
dim ref_result_str, ref_Status, ref_accessKey, ref_secretKey, ref_Key_str, vKey_confirm
	orgreqval = request("token")
	'orgreqval = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJwYXlsb2FkIjoie1wiaWRcIjogMTQ1LCBcInVzZXJfaWRcIjogXCI3M21jRXNmMlwifSIsImV4cCI6MTQzMjAxNzM0NzEzM30.NzBOMSNKdOj3gnVOgkbDwCkVSvilmAji9VDqoFf5tOU"

'response.write orgreqval & "<br><br>"
'response.end

if orgreqval="" then
	response.write "<script type='text/javascript'>alert('NOT INFO.');</script>"
	dbget.close()	:	response.end
end if

'/파라메타 내에 필드값을 체크해서 그부분만 가져오는 방식일경우
'SET oJSON = New aspJSON
'	oJSON.loadJSON(orgreqval)
'	reqtoken = oJSON.data("token")
'set oJSON = nothing
'reqtoken = FnURLDecode(reqtoken)

reqtoken = FnURLDecode(orgreqval)

'response.write reqtoken & "<br><br>"
'response.end

Dim GoMobileURL
IF application("Svr_Info")="Dev" THEN
	GoMobileURL = "http://testm.10x10.co.kr"
Else
	GoMobileURL = "https://m.10x10.co.kr"
End If

''//헤더에 키값이 정상인지 체크
'vKey_confirm = getkey_confirm(ref_accessKey, ref_secretKey, ref_Key_str)
''//키값이 안맞으면 팅겨냄
'if not(vKey_confirm) then
'	response.write "<script type='text/javascript'>alert('"& ref_Key_str &"');</script>"
'	dbget.close()	:	response.end
'end if

token = reqtoken

'/비트윈서버와 통신후 아이디 가져오기
betID = getBetweenID(token, ref_Status, ref_result_str)
'response.write betID & "/" & ref_Status & "/" & ref_result_str
'response.end

'/정상 통신인지 체크후 비정상이면 팅겨냄
if ref_Status<>"200" then
	response.write "<script type='text/javascript'>alert('"& ref_result_str &"');</script>"
	dbget.close()	:	response.end
end if

If betID <> "" Then
	Call getTenUserSn(betID, usersn)	'가져온 비트윈 ID로 usersn가져오기
End If

Const sitename = "10x10"

%>
<form name="frm" method="post" action="<%= GoMobileURL %>/apps/appcom/between/my10x10API/order/myorderlist.asp">
<input type="hidden" name="usersn" value="<%= usersn %>">
<input type="hidden" name="token" value="<%= token %>">
<input type="hidden" name="signdata" value="<%= orgreqval %>">
<input type="hidden" name="betID" value="<%= betID %>">
</form>

<script type="text/javascript">
	document.frm.submit();	
</script>

<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->