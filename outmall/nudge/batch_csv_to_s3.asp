<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 600
%>
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
on error resume next

Const MaxPage   = 40
Const PageSize = 3000

Dim appPath : appPath = server.mappath("/outmall/nudge/") + "\"
Dim FileName: FileName = "nudge_"&replace(FormatDateTime(Now(),2),"-","")&".csv"
Dim fso, tFile

Dim sqlStr, noSellStr
Dim FTotCnt, FTotPage, FCurrPage

Dim arrResultSet

'sqlStr = "SELECT * FROM [db_datamart].[dbo].[tbl_nudge_bulkinfo] "
sqlStr = "SELECT hashID "
sqlStr = sqlStr & ",isMember "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](uregdate) as uregdate "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](appRegdate) as appRegdate "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](appLastActTime) as appLastActTime "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](firstbuytime) as firstbuytime "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](lastbuytime) as lastbuytime "
sqlStr = sqlStr & ",accbuyCnt,accbuySum,Ulevel,nextMayLevelDown "
sqlStr = sqlStr & ",[db_datamart].[dbo].[fn_dt2unixTS_UTC](expireDT) as expireDT "
sqlStr = sqlStr & "FROM [db_datamart].[dbo].[tbl_nudge_bulkinfo] "

dbDatamart_rsget.Open sqlStr,dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
If not dbDatamart_rsget.EOF Then
	'arrRecordCount = dbDatamart_rsget.RecordCount
	arrResultSet = dbDatamart_rsget.getRows()
End If
dbDatamart_rsget.Close

dim nCols, nRows, xindex, yindex
dim strBuf

nCols = ubound(arrResultSet,1)
nRows = ubound(arrResultSet,2)

Set fso = CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(appPath & FileName )

For xindex = 0 To nRows
  strBuf = ""
  For yindex=0 To nCols
     If yindex>0 Then	strBuf = strBuf & ","
     'response.write arrResultSet(yindex,xindex)
     strBuf = strBuf & arrResultSet(yindex,xindex)
  Next
  'strBuf = strBuf & vbcrlf
  tFile.WriteLine strBuf
Next

tFile.Close
Set tFile = Nothing
Set fso = Nothing
%>
<!-- #include virtual="/lib/db/dbDatamartClose.asp" -->
<%
If err.number = 0 Then
	'server.execute("upload_csv_to_s3.asp")
	response.write "OK - Maker BULK"
else
    response.write err.description
End If 
%>
