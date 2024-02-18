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
'	Description : 비트윈 상품
'	History	:  2015.05.01 김진영 생성
'#######################################################
%>

<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/between/betweenCommFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim HeadAuthorization, accessKey, secretKey
Dim offset, limit
dim ref_result_str, ref_Status, ref_accessKey, ref_secretKey, ref_Key_str, vKey_confirm
	offset	= request("offset")
	limit	= request("limit")

'response.write "수정중"
'dbget.close()	:	response.end

if Not(isNumeric(offset)) then
	response.write "<script type='text/javascript'>alert('페이지 번호는 숫자만 가능 합니다.');</script>"
	dbget.close()	:	response.end
end if
if Not(isNumeric(limit)) then
	response.write "<script type='text/javascript'>alert('페이지 크기는 숫자만 가능 합니다.');</script>"
	dbget.close()	:	response.end
end if
if (limit-offset) > 20 then
	response.write "<script type='text/javascript'>alert('20페이지 이하 단위로 끊어서 가져 가세요.');</script>"
	dbget.close()	:	response.end
end if

'//헤더에 키값이 정상인지 체크
vKey_confirm = getkey_confirm(ref_accessKey, ref_secretKey, ref_Key_str)
'//키값이 안맞으면 팅겨냄
if not(vKey_confirm) then
	response.write "<script type='text/javascript'>alert('"& ref_Key_str &"');</script>"
	dbget.close()	:	response.end
end if

If (offset = "") OR (limit = "") Then
	response.write "<script type='text/javascript'>alert('시작값 또는 종료값이 없습니다.');</script>"
	dbget.close()	:	response.end
end if

Call fnBetweenItemlistJsonFlush(offset, limit)
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
