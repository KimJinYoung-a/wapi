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
<!-- #include virtual="/search/searchcls.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<%

function CheckVaildIP(ref)
    CheckVaildIP = false
    dim i
    ' dim VaildIP : VaildIP = Array("13.125.145.40","13.125.12.181","52.79.73.145","61.252.133.88","192.168.1.70","61.252.133.81","192.168.1.81","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    ' 
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
    oJson("resultCount") = 0
    Set oJson("retDatas") = jsArray()
    oJson.flush
    Set oJson = Nothing
    response.end
end if

dim reBuf
dim q : q = request("string")
q = RepWord(q,"[^ㄱ-ㅎㅏ-ㅣ가-힣a-zA-Z0-9.&%+\-\_\(\)\/\\\[\]\~\s]","")


if (LEN(q)<1) then
    Set oJson = jsObject()
    oJson("resultCode") = "S_ERR"
    oJson("resultMessage") = "no param"
    oJson("resultCount") = 0
    Set oJson("retDatas") = jsArray()
    oJson.flush
    Set oJson = Nothing
    response.end 
end if

'ret_Extract = oDoc.ExtractKeyword ''형태소분석
'ret_synonym = oDoc.GetSynonymList ''동의어


dim ret_Extract
dim oDoc,i
set oDoc = new SearchItemCls
oDoc.FRectSearchTxt = q
ret_Extract = oDoc.ExtractKeyword
ret_Extract = RTrim(ret_Extract)
SET oDoc=Nothing


Dim retArr : retArr = split(ret_Extract,vbCrlf)
Dim oDataArr, oDataItem
Set oJson = jsObject()


if UBound(retArr)>-1 then
    oJson("resultCode") = "S_OK"
    oJson("resultMessage") = ""
    oJson("resultCount") = UBound(retArr)+1


    Set oDataArr = jsArray()			'배열구조로 선언
    For i=LBound(retArr) To UBound(retArr)
        Set oDataItem = jsObject()
        oDataItem("position") = i+1
        oDataItem("word") = retArr(i)
        
        
        set oDataArr(null) = oDataItem
        SET oDataItem = Nothing
    Next	

    Set oJson("retDatas") = oDataArr
    Set oDataArr = Nothing
ELSE
    oJson("resultCode") = "S_NONE"
    oJson("resultMessage") = "No result"
    oJson("resultCount") = 0
    Set oJson("retDatas") = jsArray()
END IF

oJson.flush
Set oJson = Nothing
%>
