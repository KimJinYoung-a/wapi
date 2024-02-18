<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%

function CheckVaildIP(ref)
    CheckVaildIP = false
    dim i

    ' dim VaildIP : VaildIP = Array("13.125.145.40","13.125.12.181","52.79.73.145","61.252.133.88","192.168.1.70","61.252.133.81","192.168.1.81","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    ' for i=0 to UBound(VaildIP)
    '     if (VaildIP(i)=ref) then
    '         CheckVaildIP = true
    '         exit function
    '     end if
    ' next
    dim validToken : validToken = Array("70711546f86e45b2bb3f9b5528ded10d")
    dim authtkn : authtkn = LCASE(request("authtkn"))
    for i=0 to UBound(validToken)
        if (validToken(i)=authtkn) then
            CheckVaildIP = true
            exit function
        end if
    next

end function

Dim oJson
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    Set oJson = jsObject()
    oJson("resultCode") = "S_ERR"
    oJson("resultMessage") = "Invalid auth"
    oJson.flush
    Set oJson = Nothing
    response.end
end if



dim sqlStr
sqlStr = "exec db_const.[dbo].[usp_ten_wapi_const_category_keyword_boost_synonym_truncate]"
dbget.Execute sqlStr

Set oJson = jsObject()
oJson("resultCode") = "S_OK"
oJson("resultMessage") = ""

oJson.flush
Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->