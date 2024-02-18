<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 600
'판매EP는 77번DB를 바라보고, 상품EP는 78번 DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 40
Const PageSize = 3000

Dim appPath : appPath = server.mappath("/outmall/naverEP/") + "\"
Dim FileName: FileName = "naverDailySellEP.txt"
Dim fso, tFile

Function WriteMakeNaverFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr,bufstr2,totalSum
    iRow = UBound(arrList,2)
    bufstr = "<<<pstart>>>"&VbCRLF
    For intLoop=0 to iRow
		bufstr = bufstr&arrList(1,intLoop)&"|"&arrList(2,intLoop)&"|"&arrList(3,intLoop)&VbCRLF
		totalSum = totalSum + arrList(3,intLoop)
    Next
    bufstr = bufstr&"<<<pend>>>"

	bufstr2="<<<mstart>"&VbCRLF
	bufstr2 = bufstr2&totalSum&"|"&iRow+1&"|"&date()-1&VbCRLF
	bufstr2 = bufstr2&"<<<mend>"&VbCRLF
	tFile.WriteLine bufstr2&bufstr
End function

Dim sqlStr, noSellStr
Dim FTotCnt, FTotPage, FCurrPage

sqlStr ="[db_datamart].[dbo].[sp_Ten_Naver_EPSellDataCount]"	 
dbDatamart_rsget.Open sqlStr, dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (dbDatamart_rsget.EOF OR dbDatamart_rsget.BOF) THEN
	FTotCnt = dbDatamart_rsget(0)
END IF
dbDatamart_rsget.close
'response.write FTotCnt&"<br>"

Dim i, ArrRows
IF FTotCnt > 0 THEN
	FTotPage = CLNG(FTotCnt/PageSize)
	IF FTotPage<>(FTotCnt/PageSize) THEN FTotPage=FTotPage+1
	IF (FTotPage>MaxPage) THEN FTotPage=MaxPage

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
	For i=0 to FTotPage-1
		sqlStr ="[db_datamart].[dbo].[sp_Ten_Naver_EPSellData]("&i+1&","&PageSize&")"
		dbDatamart_rsget.Open sqlStr, dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (dbDatamart_rsget.EOF OR dbDatamart_rsget.BOF) THEN
			ArrRows = dbDatamart_rsget.getRows()
		END IF
		dbDatamart_rsget.close
		CALL WriteMakeNaverFile(tFile,ArrRows)
	NExt
	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
ELSE
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )
		noSellStr="<<<mstart>>>"&VbCRLF
		noSellStr = noSellStr & "0|0|"&date()-1&VbCRLF
		noSellStr = noSellStr&"<<<mend>>>"&VbCRLF
		noSellStr = noSellStr&"<<<pstart>>>"&VbCRLF
		noSellStr = noSellStr&"<<<pend>>>"&VbCRLF
		tFile.WriteLine noSellStr
	tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbDatamartClose.asp" -->