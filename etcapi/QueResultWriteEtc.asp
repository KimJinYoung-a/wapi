<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
''UTF-8 ë¡? ?•´?•¼ ?•œê¸??´ ?•ˆê¹¨ì??ê²? ë°›ìŒ.

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("222.109.123.95","211.206.236.117","115.94.163.42","61.252.133.88","192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72","192.168.1.69")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next

    dim validToken : validToken = Array("bd2acd564c264459908cd0d744986ea0","554740c796aa47b1aae8ee9bacd2643c")
    dim authtkn : authtkn = LCASE(request("authtkn"))
    for i=0 to UBound(validToken)
        if (validToken(i)=authtkn) then
            CheckVaildIP = true
            exit function
        end if
    next

end function


dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write "S_ERR"
    dbget.Close() :     response.end
end if


dim apiAction : apiAction=requestCheckVar(request("apiAction"),32)
dim songjangno : songjangno = requestCheckVar(request("songjangno"),32)
dim songjangdiv : songjangdiv = requestCheckVar(request("songjangdiv"),10)
dim departuredt : departuredt = requestCheckVar(request("departuredt"),30)
dim arrivedt    : arrivedt = requestCheckVar(request("arrivedt"),30)

dim company_no : company_no = requestCheckVar(request("company_no"),16)
dim regstat : regstat = requestCheckVar(request("regstat"),1)
dim vatgubun : vatgubun = requestCheckVar(request("vatgubun"),2)
dim closuredate    : closuredate = requestCheckVar(request("closuredate"),30)
dim vatconvdate    : vatconvdate = requestCheckVar(request("vatconvdate"),30)
dim datastring : datastring = requestCheckVar(request("datastring"),8000)

dim sqlStr, i, j

if (apiAction="setCompanyState") then

    if ((company_no="") or (regstat="")) then
        response.write "S_ERR"
        dbget.close() : response.end
    else
        sqlStr = "db_partner.[dbo].[usp_Ten_Company_State_API_IU] '"&company_no&"','"&regstat&"','"&vatgubun&"','"&closuredate&"','"&vatconvdate&"'"
       ''response.write sqlStr
        dbget.Execute sqlStr
    end if
elseif (apiAction="setdlvtrc") then

    if ((songjangno="") or (songjangdiv="")) then
        response.write "S_ERR"
        dbget.close() : response.end
    elseif NOT isNumeric(songjangdiv) then
        response.write "S_ERR"
        dbget.close() : response.end
    else
        sqlStr = "db_order.[dbo].[usp_Ten_Delivery_Trace_Queue_API_IU] '"&songjangno&"',"&songjangdiv&",'"&departuredt&"','"&arrivedt&"'"
       ''response.write sqlStr
        dbget.Execute sqlStr
    end if

elseif (apiAction="setnvlowest") then

    if (datastring="")  then
        response.write "S_ERR"
        dbget.close() : response.end
    else
        sqlStr = "db_temp.[dbo].[usp_ten_nv_crawl_matchid_array_insert] '"&datastring&"'"
       ''response.write sqlStr
        dbget.Execute sqlStr
    end if


elseif (apiAction="zzzzzzzzzzzzzzzzzzzz")  then
    ' sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWriteWithData] "&idx&","&itemid&",'"&sellyn&"',"&salePrice&",'"&html2DB(regitemname)&"','"&ErrCode&"','"&html2DB(ErrMsg)&"'"
    ' 'response.write sqlStr
    ' dbget.Execute sqlStr
end if



response.write "S_OK"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->