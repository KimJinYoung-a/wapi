<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<%
dim mallid      : mallid=requestCheckVar(request("mallid"),32)
dim apiAction   : apiAction=requestCheckVar(request("apiAction"),32)
dim topN        : topN=requestCheckVar(request("topN"),10)

if (topN="") then topN="5"
    
dim sqlStr
dim ArrRows, i,cnt
dim bufStr

sqlStr = "exec [db_etcmall].[dbo].[usp_Ten_Outmall_DiffOrder_Get] '"& mallid &"' " 
rsget.CursorLocation = adUseClient
rsget.CursorType = adOpenStatic
rsget.LockType = adLockOptimistic
rsget.Open sqlStr, dbget
IF not rsget.EOF THEN
	ArrRows = rsget.getRows()
END IF
rsget.Close

if IsArray(ArrRows) then
    For i=0 To UBound(ArrRows,2)		
		bufStr = bufStr&ArrRows(0,i)&"||"&ArrRows(1,i)&"||"&ArrRows(2,i)&"||"&ArrRows(3,i)&vBCRLF
	Next	
	bufStr = "S_OK"&vBCRLF&bufStr
else
    bufStr ="S_NONE" ''데이타 없음.
end if
response.write bufStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->