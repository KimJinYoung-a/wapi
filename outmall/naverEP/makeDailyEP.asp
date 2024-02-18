<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 1000  ''�ʴ���
'��ǰEP�� 78�� DB�� �ٶ󺸰�, �Ǹ�EP�� 77��DB�� �ٶ󺻴�
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'' ���̹� ���ļ��� ���� Make / �Ϻ�
Const MaxPage   = 100   ''maxpage ���� 40->50���� 2013-12-13����, 50->60���� 2014-09-23 ������ ����, 60->70���� 2014-10-08 ���� ,70->100 ���� 2016-06-29
Const PageSize = 5000  ''3000->5000

Dim appPath : appPath = server.mappath("/outmall/naverEP/") + "\"
Dim FileName: FileName = "naverDailyEP_temp.txt"
Dim newFileName: newFileName = "naverDailyEP.txt"
Dim fso, tFile

Dim IsChangedEP : IsChangedEP = (request("epType")="chg")
If (IsChangedEP) Then
	FileName = "naverChangedEP_temp.txt"
	newFileName = "naverChangedEP.txt"
End If

Function WriteMakeNaverFile(tFile, arrList, isIsChangedEP,byref iLastItemid )
    Dim intLoop,iRow
    Dim bufstr
    Dim itemid,deliverytype, deliv
    Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
    Dim tmpLastDeptNM, itemname
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
'���ϴ� ����ī�װ�
		itemid			= arrList(1,intLoop)
		deliverytype	= arrList(8,intLoop)
		deliv 			= arrList(19,intLoop)  ''��ۺ� /2000, 2500, 0

		IF isNULL(arrList(20,intLoop)) then  ''2013/12/07 �߰�
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

		'2014-01-15 ������ �߰�..Ư���귣��� 3depth���� �������� �Ѱɷ� �ְ� �ʹ��Ͽ� �߰���
'		IF (isIsChangedEP) then			'���EP
'			If (arrList(24,intLoop) <> "") AND (arrList(21,intLoop) = "U") Then			'�߰��� depth���� �ְ� ��������(U)
'				jaehu3depNM = jaehu3depNM &"_"& db2html(arrList(24,intLoop))
'			End If
'		Else							'��üEP
'			If (arrList(24,intLoop) <> "") Then
'				jaehu3depNM = jaehu3depNM &"_"& db2html(arrList(24,intLoop))
'			End If
'		End If

		'�켱 ���� Depth3ItemName > Depth3MakerName > ����ī�װ���
		If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
			IF (isIsChangedEP) then			'���EP
				If arrList(21,intLoop) = "U" Then	'��������(U)
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			Else
				If (arrList(24,intLoop) <> "") OR (arrList(25,intLoop) <> "") Then
					If (arrList(25,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(25,intLoop))
					ElseIf (arrList(24,intLoop) <> "") Then
						jaehu3depNM = db2html(arrList(24,intLoop))
					End If
				End If
			End If
		End If

		bufstr = "<<<begin>>>"&VbCRLF
		bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
		bufstr = bufstr&"<<<pname>>>[�ٹ�����] "&itemname&"_"&jaehu3depNM&VbCRLF
		bufstr = bufstr&"<<<price>>>"&CLNG(arrList(3,intLoop))&VbCRLF
		bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=nvshop_sp"&VbCRLF
		bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
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

		If (Now() > #06/11/2017 21:00:00# AND Now() < #06/25/2017 20:59:59#) Then		'���� ��ȿ�Ⱓ �����̶��..
			bufstr = bufstr&"<<<event>>>�ڳ��̹� ���ļ��� ���� 3,000�� ����������"&VbCRLF
'			bufstr = bufstr&"<<<event>>>�� �ٹ����ٿ��� ���̹����̷� ù ���Ž� 2,000�� ����(8/12~25)"&VbCRLF
		Else
			bufstr = bufstr&"<<<event>>>�� ���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"&VbCRLF
'			bufstr = bufstr&"<<<event>>>�� 13�ֳ� ����� ���� �� ���ϸ��� ����! �ִ� 30% ����!"&VbCRLF
		End If

		IF arrList(22,intLoop) <> "" THEN
			bufstr = bufstr&"<<<coupo>>>"&Replace(arrList(22,intLoop),"&nbsp;","")&VbCRLF
		END IF
		bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
		IF (isIsChangedEP) then
			bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF                ''�ϴ� I,U �� //D ǰ��
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

''�ۼ��ð� üũ
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_CH_ST')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_DY_ST')"
    dbCTget.execute sqlStr
end if


if (IsChangedEP) then
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataCount](1)"
else
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPDataCount]"
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
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPData]("&i+1&","&PageSize&",1,"&iLastItemid&")"
        else
            sqlStr ="[db_outmall].[dbo].[sp_Ten_Naver_EPData]("&i+1&","&PageSize&",0,"&iLastItemid&")"
        end if

        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
        	ArrRows = rsCTget.getRows()
        END IF
        rsCTget.close

        if isArray(ArrRows) then
            CALL WriteMakeNaverFile(tFile,ArrRows, IsChangedEP, iLastItemid)
        end if

        ''�ۼ��ð� üũ
        IF(IsChangedEP) then
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('nvshop_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        else
            sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
            sqlStr = sqlStr + " (ref) values('nvshop_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
            dbCTget.execute sqlStr
        end if
    NExt

    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

''�ۼ��ð� üũ
IF(IsChangedEP) then
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_CH_ED')"
    dbCTget.execute sqlStr
else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog"
    sqlStr = sqlStr + " (ref) values('nvshop_DY_ED')"
    dbCTget.execute sqlStr
end if

'2013-12-10 15:40 ������ �߰� TEMP������ ���� ���Ϸ� ����
Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	Newfso.CopyFile appPath & FileName ,appPath & newFileName
Set Newfso = nothing
response.write FTotCnt&"�� ���� ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->