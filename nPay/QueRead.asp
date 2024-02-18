<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.71","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function debugWrite(iTxt)
    if (application("Svr_Info")	= "Dev") then
        if (request("isautoscript")<>"on") then
            response.write iTxt&"<br>"
        end if
    end if
end function

''''-------------------------------------------------------------
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    response.write ref
    response.end
end if



dim apiAction   : apiAction=requestCheckVar(request("apiAction"),32)
dim topN        : topN=requestCheckVar(request("topN"),10)

if (topN="") then topN="5"

'topN="1"

dim sqlStr
dim ArrRows, i,cnt
dim bufStr

if (apiAction="CASHAMT") OR (apiAction="FIN") then
    sqlStr = "db_cs.[dbo].[sp_Ten_NPay_Que_Read] ('"&apiAction&"',"&topN&")"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
		ArrRows = rsget.getRows()
	END IF
	rsget.close
    
    if IsArray(ArrRows) then
        For i=0 To UBound(ArrRows,2)		
			bufStr = bufStr&ArrRows(0,i)&"||"&ArrRows(1,i)&"||"&ArrRows(2,i)&vBCRLF
		Next	
		bufStr = "S_OK"&vBCRLF&bufStr
    else
        bufStr ="S_NONE" ''데이타 없음.
    end if
    response.write bufStr
elseif (apiAction="TT") then

else
    response.write "ERR||0||미지정"&apiAction
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->