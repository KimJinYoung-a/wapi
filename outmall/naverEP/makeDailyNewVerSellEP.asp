<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 600
'판매EP는 214번DB를 바라보고, 상품EP는 78번 DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 40
Const PageSize = 3000

Dim appPath : appPath = server.mappath("/outmall/naverEP/") + "\"
Dim FileName: FileName = "naverNewVerDailySellEP.txt"
Dim fso, tFile

Function WriteMakeNaverFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr
    iRow = UBound(arrList,2)
    bufstr = "mall_id"&vbTab&"sale_count"&vbTab&"sale_price"&vbTab&"order_count"&vbTab&"dt"&VbCRLF
    For intLoop=0 to iRow
		bufstr = bufstr&arrList(0,intLoop)&vbTab&arrList(1,intLoop)&vbTab&arrList(2,intLoop)&vbTab&arrList(3,intLoop)&vbTab&arrList(4,intLoop)&VbCRLF
    Next
    
	tFile.WriteLine bufstr

    WriteMakeNaverFile = iRow+1
End function

Dim sqlStr
Dim FTotCnt : FTotCnt = 0

Dim i, ArrRows

Set fso = CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(appPath & FileName )

    sqlStr ="[db_statistics_const].[dbo].[usp_Ten_Naver_EPSellDataNew_YYYYMMDD] "
    rsSTSget.Open sqlStr, dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    IF Not (rsSTSget.EOF OR rsSTSget.BOF) THEN
        ArrRows = rsSTSget.getRows()
    END IF
    rsSTSget.close

    if isArray(ArrRows) then
        FTotCnt = WriteMakeNaverFile(tFile,ArrRows)
    end if
tFile.Close
Set tFile = Nothing
Set fso = Nothing


response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbSTSClose.asp" -->