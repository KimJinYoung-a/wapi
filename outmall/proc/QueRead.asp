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

if (apiAction="CHKSTOCK") OR (apiAction="SOLDOUT") OR (apiAction="PRICE") OR (apiAction="ITEMNAME") OR (apiAction="EDIT") OR (apiAction="EDIT2") OR (apiAction="EDITBATCH") or (apiAction="CHKSTAT") or (apiAction="CONFIRM") or (apiAction="DISPVIEW") or (apiAction="KEEPSELL") or (apiAction="VIEWOPT") or (apiAction="CATEGORY") or (apiAction="EDITIMG") or (apiAction="REG") or (apiAction="REGOnSale") or (apiAction="Image") or (apiAction="EDITQTY") or (apiAction="DEL") or (apiAction="DELETE") or (apiAction="OPTSTAT") or (apiAction="IMAGE") or (apiAction="CHKOPT") or (apiAction="capEXPR") or (apiAction="REGBATCH") or (apiAction="REG2") or (apiAction="CONTENT") then
    sqlStr = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_READ] ('"&apiAction&"','"&mallid&"',"&topN&")"
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
        bufStr ="S_NONE" ''����Ÿ ����.
    end if
    response.write bufStr
elseif (apiAction="TT") then

else
    response.write "ERR||0||������"&apiAction
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->