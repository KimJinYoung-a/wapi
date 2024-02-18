<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1000  ''초단위
'상품EP는 78번 DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 70
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/daumEP/") + "\"
Dim FileName: FileName = "daumDailyEP_temp.txt"
Dim newFileName: newFileName = "daumDailyEP.txt"
Dim fso, tFile
Dim EPTYPE
EPTYPE = requestCheckVar(Request("epType"),7)

If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"		'요약EP
			FileName = "daumChangedEP_temp.txt"
			newFileName = "daumChangedEP.txt"
		Case "new"		'신규EP
			FileName = "daumNewEP_temp.txt"
			newFileName = "daumNewEP.txt"
		Case "best100"	'베스트100EP
			FileName = "daumBest100EP_temp.txt"
			newFileName = "daumBest100EP.txt"
		Case "review"	'상품평EP
			FileName = "daumReviewEP_temp.txt"
			newFileName = "daumReviewEP.txt"
	End Select
End If

Function FormatDate(ddate, formatstring)
	dim s
	Select Case formatstring
		Case "0000-00-00 00:00:00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "0000-00-00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000"
			s = CStr(year(ddate)) &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000000000"
			s = CStr(year(ddate))  &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")  &_
				Num2Str(hour(ddate),2,"0","R")  &_
				Num2Str(minute(ddate),2,"0","R") &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R")
		Case "0000.00.00-00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & "-" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00.00 00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000/00/00"
			s = CStr(year(ddate)) & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00/00"
			s = Num2Str(year(ddate),2,"0","R") & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00.00"
			s = Num2Str(year(ddate),2,"0","R") & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00"
			s = Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00"
			s = Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case Else
			s = CStr(ddate)
	End Select

	FormatDate = s
End Function

Function printUserId(strID,lng,chr)
	dim le, te
	le = len(strID)
	if(le<lng) Then
		printUserId = String(lng, le)
		Exit Function
	end if
	te = left(strID,le-lng) & String(lng, chr)
	printUserId = te
End Function

Function WriteMakeDaumFile(tFile, arrList, isEPTYPE, byref iLastItemid)
	Dim intLoop,iRow
	Dim bufstr
	Dim itemid,deliverytype, deliv
	Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
	Dim tmpLastDeptNM, itemname
	iRow = UBound(arrList,2)
	If (isEPTYPE = "" OR isEPTYPE = "chg" OR isEPTYPE = "new") Then
	    For intLoop=0 to iRow
			itemid			= arrList(1,intLoop)
			If arrList(19,intLoop) = 0 AND arrList(8,intLoop) <> "7" Then
				deliv = 0
			ElseIf arrList(8,intLoop) = "7" Then
				deliv = -1
			Else
				deliv = 1
			End If

			IF isNULL(arrList(20,intLoop)) then
			    ArrCateNM		= ""
	    		CntNM			= Split(ArrCateNM,",")
	    		ArrCateCD		= ""
	    		CntCD			= Split(ArrCateCD,",")
	    		jaehu3depNM		= ""
			Else
	    		ArrCateNM		= Split(arrList(20,intLoop),"||")(0)
	    		CntNM			= Split(ArrCateNM,",")
	    		ArrCateCD		= Split(arrList(20,intLoop),"||")(1)
	    		CntCD			= Split(ArrCateCD,",")
	    		jaehu3depNM		= Split(arrList(20,intLoop),"||")(2)
	        End If

			itemname = arrList(2,intLoop)
			itemname = Replace(itemname,"&nbsp;","")
			itemname = Replace(itemname,"&nbsp","")
			If (arrList(24,intLoop) <> "") Then
				jaehu3depNM = jaehu3depNM &"_"& db2html(arrList(24,intLoop))
			End If
			bufstr = "<<<begin>>>"&VbCRLF
			bufstr = bufstr&"<<<pid>>>"&itemid&VbCRLF
			bufstr = bufstr&"<<<price>>>"&CLNG(arrList(3,intLoop))&VbCRLF
			bufstr = bufstr&"<<<pname>>>[텐바이텐] "&itemname&"_"&jaehu3depNM&VbCRLF
			If isEPTYPE <> "chg" Then
				bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=daumshop"&VbCRLF
				bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
				For lp2=1 to Ubound(CntCD)+1
					If lp2>4 Then Exit For
					bufstr = bufstr&"<<<cate"&lp2&">>>"&CntCD(lp2-1)&VbCRLF
				Next
				If lp2 < 5 Then
					For lp2=lp2 to 4
						bufstr = bufstr&"<<<cate"&lp2&">>>"&VbCRLF
					Next
				End If

				For lp=1 to Ubound(CntNM)+1
					If lp>4 Then Exit For
					bufstr = bufstr&"<<<catename"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
				Next
				If lp < 5 Then
					For lp=lp to 4
						bufstr = bufstr&"<<<catename"&lp&">>>"&VbCRLF
					Next
				End If
				bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
				bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
				bufstr = bufstr&"<<<sales>>>"&arrList(12,intLoop)&VbCRLF
				IF arrList(22,intLoop) <> "" THEN
					bufstr = bufstr&"<<<coupon>>>"&Replace(arrList(22,intLoop),"&nbsp;","")&VbCRLF
				Else
					bufstr = bufstr&"<<<coupon>>>"&VbCRLF
				END IF

				If (Now() > #02/21/2016 21:00:00# AND Now() < #02/28/2016 20:59:59#) Then		'쿠폰 유효기간 동안이라면..
					bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
				End If
				bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
				bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
				If deliv = 1 Then
					bufstr = bufstr&"<<<deliv2>>>"&arrList(19,intLoop)&VbCRLF
				End If
				bufstr = bufstr&"<<<review>>>"&arrList(15,intLoop)&VbCRLF

				If (Now() > #02/21/2016 21:00:00# AND Now() < #02/28/2016 20:59:59#) Then		'쿠폰 유효기간 동안이라면..
					bufstr = bufstr&"<<<event>>>★다음 쇼핑하우 전용 3,000원 할인쿠폰★"&VbCRLF
				Else
					bufstr = bufstr&"<<<event>>>▶ 구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"&VbCRLF
				End If
			End If
			bufstr = bufstr&"<<<end>>>"
			tFile.WriteLine bufstr
			iLastItemid = itemid
	    Next
	ElseIf (isEPTYPE = "review") Then
		For intLoop=0 to iRow
			bufstr = "<<<begin>>>"&VbCRLF
			bufstr = bufstr&"<<<mallid>>>tenbyten10x10"&VbCRLF
			bufstr = bufstr&"<<<pid>>>"&arrList(1,intLoop)&VbCRLF
			bufstr = bufstr&"<<<reviewid>>>"&arrList(2,intLoop)&VbCRLF
			If Len(db2html(arrList(3,intLoop))) > 20 Then
				bufstr = bufstr&"<<<title>>>"&Trim(LEFT(db2html(arrList(3,intLoop)),10))&"..."&VbCRLF
			Else
				bufstr = bufstr&"<<<title>>>"&db2html(arrList(3,intLoop))&VbCRLF
			End If
			bufstr = bufstr&"<<<content>>>"&db2html(arrList(3,intLoop))&VbCRLF
			bufstr = bufstr&"<<<writer>>>"&printUserId(arrList(4,intLoop), 2, "*")&VbCRLF
			bufstr = bufstr&"<<<cdate>>>"&FormatDate(arrList(5,intLoop),"00000000000000")&VbCRLF
			bufstr = bufstr&"<<<point>>>"&CInt(arrList(6,intLoop)) + 1&VbCRLF
			bufstr = bufstr&"<<<end>>>"
			tFile.WriteLine bufstr
			iLastItemid = arrList(2,intLoop)
		Next
	ElseIf (isEPTYPE = "best100") Then
		For intLoop=0 to iRow
			bufstr = "<<<begin>>>"&VbCRLF
			bufstr = bufstr&"<<<pid>>>"&arrList(1,intLoop)&VbCRLF
			bufstr = bufstr&"<<<cate1>>>"&arrList(2,intLoop)&VbCRLF
			bufstr = bufstr&"<<<catename1>>>"&arrList(3,intLoop)&VbCRLF
			If CInt(arrList(0,intLoop)) <= 100 Then
				bufstr = bufstr&"<<<toprank>>>"&arrList(0,intLoop)&VbCRLF
			Else
				bufstr = bufstr&"<<<toprank>>>"&VbCRLF
			End If
			bufstr = bufstr&"<<<cate1rank>>>"&arrList(4,intLoop)&VbCRLF
			bufstr = bufstr&"<<<end>>>"
			tFile.WriteLine bufstr
		Next
	End If
End Function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''작성시간 체크
If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_ST')"
		Case "new"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NW_ST')"
		Case "best100"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_ST')"
		Case "review"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_ST')"
	End Select
    dbCTget.execute sqlStr
Else
	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_DY_ST')"
	dbCTget.execute sqlStr
End If

If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPDataCount]('C')"
		Case "new"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPDataCount]('N')"
		Case "best100"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPDataCount]('B')"
		Case "review"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPDataCount]('R')"
	End Select
Else
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPDataCount]"
End If
rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
	FTotCnt = rsCTget(0)
END IF
rsCTget.close

Dim i, ArrRows
Dim iLastItemid : iLastItemid=99999999

IF FTotCnt > 0 THEN
    FTotPage = CLNG(FTotCnt/PageSize)
    IF FTotPage<>(FTotCnt/PageSize) THEN FTotPage=FTotPage+1
    IF (FTotPage>MaxPage) THEN FTotPage=MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(appPath & FileName )
		    For i=0 to FTotPage-1
		        ArrRows = ""
		        If EPTYPE <> "" Then
					Select Case EPTYPE
						Case "chg"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPData]("&i+1&","&PageSize&",'C',"&iLastItemid&")"
						Case "new"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPData]("&i+1&","&PageSize&",'N',"&iLastItemid&")"
						Case "best100"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPData]("&i+1&","&PageSize&",'B',"&iLastItemid&")"
						Case "review"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPData]("&i+1&","&PageSize&",'R',"&iLastItemid&")"
					End Select
		        Else
		            sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_EPData]("&i+1&","&PageSize&",'A',"&iLastItemid&")"
		        End If

		        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		        	ArrRows = rsCTget.getRows()
		        END IF
		        rsCTget.close

		        If isArray(ArrRows) then
		            CALL WriteMakeDaumFile(tFile, ArrRows, EPTYPE, iLastItemid)
		        End If

		        ''작성시간 체크
				If EPTYPE <> "" Then
					Select Case EPTYPE
						Case "chg"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "new"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NW_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "best100"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "review"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_"&(i+1)*PageSize&"_"&iLastItemid&"')"
					End Select
					dbCTget.execute sqlStr
				Else
					sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_DY_"&(i+1)*PageSize&"_"&iLastItemid&"')"
					dbCTget.execute sqlStr
				End If
		    Next
	    tFile.Close
		Set tFile = Nothing
	Set fso = Nothing
END IF

''작성시간 체크
If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_ED')"
		Case "new"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NW_ED')"
		Case "best100"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_ED')"
		Case "review"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_ED')"
	End Select
    dbCTget.execute sqlStr
Else
    sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_DY_ED')"
    dbCTget.execute sqlStr
End If

Dim Newfso
Set Newfso = Server.CreateObject("Scripting.FileSystemObject")
	If Newfso.FileExists(appPath & FileName) Then
		Newfso.CopyFile appPath & FileName ,appPath & newFileName
	End If
Set Newfso = nothing
response.write FTotCnt&"건 생성 ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->