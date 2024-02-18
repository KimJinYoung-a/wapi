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

Dim oJson, oDataArr, oDataItem
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    Set oJson = jsObject()
    oJson("resultCode") = "S_ERR"
    oJson("resultMessage") = "Invalid auth"
    oJson("resultCount") = 0
    Set oJson("retDatas") = jsArray()
    oJson.flush
    Set oJson = Nothing
    response.end
end if

dim ArrRows,i
dim sqlStr


sqlStr = "db_const.dbo.[usp_ten_wapi_const_category_keyword_boost_synonym_select]"
rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF not rsget.EOF THEN
    ArrRows = rsget.getRows()
END IF
rsget.close

Set oJson = jsObject()
if IsArray(ArrRows) then
    oJson("resultCode") = "S_OK"
    oJson("resultMessage") = ""
    oJson("resultCount") = UBound(ArrRows,2)+1
    
    Set oDataArr = jsArray()			'배열구조로 선언
    For i=0 To UBound(ArrRows,2)
        Set oDataItem = jsObject()
            
        oDataItem("catecode") = ArrRows(0,i)
        oDataItem("keyword") = TRIM(ArrRows(1,i))

        set oDataArr(null) = oDataItem
        SET oDataItem = Nothing
    Next	
    
    Set oJson("retDatas") = oDataArr
    Set oDataArr = Nothing
else
    oJson("resultCode") = "S_NONE" ''데이타 없음.
    oJson("resultMessage") = ""
    oJson("resultCount") = 0
end if

oJson.flush
Set oJson = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->