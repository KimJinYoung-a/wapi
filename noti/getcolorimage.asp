<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    'response.write ref
    'response.end
end if


Dim topn : topN = request("topN")
if LEN(topN)<1 then topN=5
if LEN(topN)>3 then topN=1000

dim sqlStr, ArrRows
dim retData, i
dim queidx,itemid,imgurl

    sqlStr = "db_etcmall.[dbo].[usp_Ten_ColorImage_Que_Get]("&topN&")"
    rsget.Open sqlStr,dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF not rsget.EOF THEN
    	ArrRows = rsget.getRows()
    END IF
    rsget.close
    
    if IsArray(ArrRows) then
        retData = "{"
        retData = retData & """count"" : "&UBound(ArrRows,2)+1&","
        retData = retData & """datalist"" : ["
        
            For i=0 To UBound(ArrRows,2)
                queidx = ArrRows(0,i)
                itemid = ArrRows(1,i)
                imgurl = ArrRows(2,i)
                imgurl = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(itemid) + "/" + imgurl

                retData = retData & "	{"
                retData = retData & "	""queidx"" : "&queidx&","
                retData = retData & "	""imgurl"": """&imgurl&""""
                retData = retData & "	}"
                
                if i<>UBound(ArrRows,2) then
                    retData = retData & ","
                end if
                    
            Next
        
        retData = retData & "]"
        retData = retData & "}"
    else
        retData = "{"
        retData = retData & """count"" : 0,"
        retData = retData & """datalist"" : ["
        retData = retData & "]"
        retData = retData & "}"
    end if
response.write retData


%>
<!-- #include virtual="/lib/db/dbclose.asp" -->