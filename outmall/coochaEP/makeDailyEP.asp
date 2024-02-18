<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000 ''초단위
'상품EP는 78번 DB를 바라보고, 판매EP는 77번DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' 네이버 지식쇼핑 파일 Make / 일별
Const MaxPage   = 100   ''maxpage 100..2015-05-07 김진영 변경
Const PageSize = 5000  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/coochaEP/") + "\"
Dim FileName: FileName = "coochaDailyEP_temp.txt"
Dim newFileName: newFileName = "coochaDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "coochaChangedEP_temp.txt"
	newFileName = "coochaChangedEP.txt"
End If

Function WriteMakeCoochaFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr
    Dim itemid,deliverytype, deliv
    Dim ArrCateNM, ArrCateCD, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname
    iRow = UBound(arrList,2)

    For intLoop=0 to iRow
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''배송비 /2000, 2500, 0

		itemname		= "[텐바이텐]"&arrList(2,intLoop)
		itemname		= Replace(itemname,"&nbsp;","")
		itemname		= Replace(itemname,"&nbsp","")
		If (deliverytype = "7") Then deliv=-1

		If isNULL(arrList(20,intLoop)) Then  ''2013/12/07 추가
			ArrCateNM		= ""
			CntNM			= Split(ArrCateNM,",")
			ArrCateCD		= ""
			CntCD			= Split(ArrCateCD,",")
		Else
			ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
			CntNM			= Split(ArrCateNM,",")
			ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
			CntCD			= Split(ArrCateCD,",")
        End If
		
		bufstr = "<<<begin>>>"&VbCRLF
		bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
		bufstr = bufstr&"<<<dealid>>>"&itemid&VbCRLF									'*상품이속한딜ID	| 상품ID로 하지고 협의
		bufstr = bufstr&"<<<mainpid>>>1"&VbCRLF											'*딜의메인상품		| 1:대표상품, 0:대표상품아님	| 1로 협의봄
		bufstr = bufstr&"<<<pname>>>"&itemname&VbCRLF
		bufstr = bufstr&"<<<price>>>"&CLNG(arrList(3,intLoop))&VbCRLF
		bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&VbCRLF
		bufstr = bufstr&"<<<ccwurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=coocha"&VbCRLF
		bufstr = bufstr&"<<<ccmurl>>>"&"http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=coocha"&VbCRLF
		bufstr = bufstr&"<<<cmwurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=coomoa"&VbCRLF
		bufstr = bufstr&"<<<cmmurl>>>"&"http://m.10x10.co.kr/category/category_itemPrd.asp?itemid="&itemid&"&rdsite=coomoa"&VbCRLF
		bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
		bufstr = bufstr&"<<<buycount>>>"&arrList(12,intLoop)&VbCRLF
		For lp=1 to Ubound(CntNM)+1
			If lp>4 Then Exit For
			bufstr = bufstr&"<<<cate"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
'			tmpLastDeptNM = CntNM(lp-1)
		Next
		If lp < 5 Then
			For lp=lp to 4
				bufstr = bufstr&"<<<cate"&lp&">>>"&VbCRLF
			Next
		End If
		bufstr = bufstr&"<<<model>>>"&VbCRLF											'*없을경우 NULL
		bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
		bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
		bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF

		IF arrList(22,intLoop) <> "" THEN
			bufstr = bufstr&"<<<coupo>>>"&Replace(arrList(22,intLoop),"&nbsp;","")&VbCRLF
		END IF
		bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
		IF (isIsChangedEP) then
			bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
			bufstr = bufstr&"<<<utime>>>"&arrList(10,intLoop)&VbCRLF
		End If
		bufstr = bufstr&"<<<mpric>>>"&VbCRLF											'<<<price>>>와 같을경우 NULL 가능
		bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
		bufstr = bufstr&"<<<ftend>>>"
		tFile.WriteLine bufstr
		bufstr = ""
		iLastItemid = itemid
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('coocha_CH_ST')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('coocha_DY_ST')"
    dbCTget.execute sqlStr
end if


if (IsChangedEP) then
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPDataCount](1)"
else
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPDataCount]"
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
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
        else
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Coocha_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
        end if

        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeCoochaFile(tFile,ArrRows, IsChangedEP, iLastItemid)
        end if

        ''작성시간 체크
        IF(IsChangedEP) then
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('coocha_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        else
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('coocha_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
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
    sqlStr = sqlStr + " (ref) values('coocha_CH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('coocha_DY_ED')"
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