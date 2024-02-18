<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000  ''초단위
'상품EP는 78번 DB를 바라보고, 판매EP는 77번DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' 네이버 지식쇼핑 파일 Make / 일별
Const MaxPage   = 100
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/shodocEP/") + "\"
Dim FileName: FileName = "shodocDailyEP_temp.txt"
Dim newFileName: newFileName = "shodocDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "shodocChangedEP_temp.txt"
	newFileName = "shodocChangedEP.txt"
End If

Function WriteMakeShodocFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
	Dim intLoop,iRow
	Dim bufstr
	Dim itemid,deliverytype, deliv
	Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
	Dim itemname
	iRow = UBound(arrList,2)
	For intLoop=0 to iRow
'이하는 전시카테고리
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''배송비 /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 추가
			ArrCateNM		= ""
			CntNM			= Split(ArrCateNM,",")
			ArrCateCD		= ""
			CntCD			= Split(ArrCateCD,",")
			jaehu3depNM		= ""
		else
			ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
			CntNM			= Split(ArrCateNM,",")
			ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
			CntCD			= Split(ArrCateCD,",")
			jaehu3depNM		= Split(arrList(20,intLoop),"||")(2)
        end if

		itemname		= arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		If (deliverytype = "7") Then deliv=-1

		bufstr = "<<<begin>>>"&VbCRLF
		bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
		bufstr = bufstr&"<<<pname>>>[텐바이텐] "&itemname&"_"&jaehu3depNM&VbCRLF
		bufstr = bufstr&"<<<price>>>"&CLNG(arrList(3,intLoop))&VbCRLF
		bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=shodoc"&VbCRLF
		bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) & "/" & arrList(4,intLoop)&VbCRLF
		For lp=1 to Ubound(CntNM)+1
			If lp>4 Then Exit For
			bufstr = bufstr&"<<<cate"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr&"<<<cate"&lp&">>>"&VbCRLF
			Next
		End If

		For lp2=1 to Ubound(CntCD)+1
			If lp2>4 Then Exit For
			bufstr = bufstr&"<<<caid"&lp2&">>>"&CntCD(lp2-1)&VbCRLF
		Next
		If lp2 < 5 Then
			For lp2=lp2 to 4
				bufstr = bufstr&"<<<caid"&lp2&">>>"&VbCRLF
			Next
		End If
		bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
		bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
		bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF

		If (Now() > #08/12/2016 21:00:00# AND Now() < #08/25/2016 20:59:59#) Then		'쿠폰 유효기간 동안이라면..
			bufstr = bufstr&"<<<event>>>▶ 텐바이텐에서 네이버페이로 첫 구매시 2,000원 적립(8/12~25)"&VbCRLF
		Else
			bufstr = bufstr&"<<<event>>>▶ 구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"&VbCRLF
		End If

		IF arrList(22,intLoop) <> "" THEN
			bufstr = bufstr&"<<<coupo>>>"&Replace(arrList(22,intLoop),"&nbsp;","")&VbCRLF
		END IF
		bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
		IF (isIsChangedEP) then
			bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF                ''일단 I,U 만 //D 품절
			bufstr = bufstr&"<<<utime>>>"&arrList(10,intLoop)&VbCRLF
		End If
		bufstr = bufstr&"<<<mpric>>>"&CLNG(arrList(3,intLoop))&VbCRLF
		bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
		bufstr = bufstr&"<<<ftend>>>"
		tFile.WriteLine bufstr
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
IF(IsChangedEP) then
	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
	sqlStr = sqlStr + " (ref) values('shodoc_CH_ST')"
	dbCTget.execute sqlStr
else
	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
	sqlStr = sqlStr + " (ref) values('shodoc_DY_ST')"
	dbCTget.execute sqlStr
end if


if (IsChangedEP) then
	sqlStr ="[db_outmall].[dbo].[sp_Ten_Shodoc_EPDataCount](1)"
else
	sqlStr ="[db_outmall].[dbo].[sp_Ten_Shodoc_EPDataCount]"
end if
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close

'response.write FTotCnt&"<br>"

Dim i, ArrRows
Dim iLastItemid : iLastItemid=9999999

IF FTotCnt > 0 THEN
	FTotPage = CLNG(FTotCnt/PageSize)
	IF FTotPage<>(FTotCnt/PageSize) THEn FTotPage=FTotPage+1
	IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set tFile = fso.CreateTextFile(appPath & FileName )

	For i=0 to FTotPage-1
		ArrRows = ""
		if (IsChangedEP) then
			sqlStr ="[db_outmall].[dbo].[sp_Ten_Shodoc_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
		else
			sqlStr ="[db_outmall].[dbo].[sp_Ten_Shodoc_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
		end if

		rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
			ArrRows = rsCTget.getRows()
		END IF
		rsCTget.close

		if isArray(ArrRows) then
			CALL WriteMakeShodocFile(tFile,ArrRows, IsChangedEP, iLastItemid)
		end if

		''작성시간 체크
		IF(IsChangedEP) then
			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
			sqlStr = sqlStr + " (ref) values('shodoc_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
			dbCTget.execute sqlStr
		else
			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
			sqlStr = sqlStr + " (ref) values('shodoc_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
			dbCTget.execute sqlStr
		end if
    NExt
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''작성시간 체크
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('shodoc_CH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('shodoc_DY_ED')"
    dbCTget.execute sqlStr
end if

'2013-12-10 15:40 김진영 추가 TEMP파일을 원본 파일로 복사
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->