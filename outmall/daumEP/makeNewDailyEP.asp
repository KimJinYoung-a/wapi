<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''초단위
'상품EP는 78번 DB를 바라본다
%>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Const MaxPage   = 800           ''2016-06-29 70->100
Const PageSize = 5000

Dim appPath : appPath = server.mappath("/outmall/daumEP/") + "\"
Dim FileName: FileName = "daumAllEP_temp.txt"
Dim newFileName: newFileName = "daumAllEP.txt"
Dim fso, tFile
Dim EPTYPE
EPTYPE = requestCheckVar(Request("epType"),10)

If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"			'요약EP
			FileName = "daumUpdateEP_temp.txt"
			newFileName = "daumUpdateEP.txt"
		Case "best100"		'베스트100EP
			FileName = "daumBest100EP_temp.txt"
			newFileName = "daumBest100EP.txt"
		Case "review"		'상품평EP
			FileName = "daumReviewEP_temp.txt"
			newFileName = "daumReviewEP.txt"
		Case "Newreview"	'신버전 상품평EP
			FileName = "daumNewReviewEP_temp.txt"
			newFileName = "daumNewReviewEP.txt"
		Case "Newreview2"	'신버전 상품평EP (전체전달용)
			FileName = "daumAllReviewEP_temp.txt"
			newFileName = "daumAllReviewEP.txt"
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

Function WriteMakeDaumFile(tFile, arrList, isEPTYPE, byref iLastItemid, iFTotCnt)
	Dim intLoop,iRow
	Dim bufstr
	Dim itemid,deliverytype, deliv
	Dim ArrCateNM, ArrCateCD, jaehu3depNM, CntNM, CntCD, lp, lp2
	Dim tmpLastDeptNM, itemname
	Dim isCouponYN : isCouponYN = "N"
	Dim iCouponEvtText
	Dim cpnAssignPrice, cpnVal ''2018/11/30 추가

	iRow = UBound(arrList,2)

	If (Now() > #05/22/2016 21:00:00# AND Now() < #05/29/2016 20:59:59#) Then
		isCouponYN = "Y"
		iCouponEvtText = "30,000원 이상 구매 시 3,000원 할인 쿠폰 증정 & 마일리지 적립"
	End If

	If (Now() > #10/01/2018 00:00:00# AND Now() < #10/01/2018 21:59:59#) Then
		isCouponYN = "Y"
		iCouponEvtText = "10/1일 오늘 단 하루만! 최대 3만원 할인쿠폰"
	End If

	If (isEPTYPE = "" OR isEPTYPE = "chg") Then
		If (isEPTYPE = "") Then
			'If iFTotCnt >= 350000 Then
			'	bufstr = "<<<tocnt>>>350000"
			'Else
				bufstr = "<<<tocnt>>>"&iFTotCnt
			'End If
			tFile.WriteLine bufstr
		End If

	    For intLoop=0 to iRow
			itemid			= arrList(1,intLoop)
			deliverytype	= arrList(8,intLoop)
			deliv 			= arrList(19,intLoop)  ''배송비 /2000, 2500, 0

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

	    		'2뎁쓰면 2뎁쓰명이 나오게 수정..2017-10-17 김진영
	    		If Ubound(CntNM) = 1 then
					jaehu3depNM = Split(ArrCateNM, ",")(1)
		    	End If
	        End If

			itemname = arrList(2,intLoop)
			itemname = Replace(itemname,"&nbsp;","")
			itemname = Replace(itemname,"&nbsp","")
			If (deliverytype = "7") Then deliv=-1
			If (arrList(24,intLoop) <> "") Then
				jaehu3depNM = jaehu3depNM &"_"& db2html(arrList(24,intLoop))
			End If

			cpnAssignPrice = CLNG(arrList(3,intLoop))
			if (CLNG(arrList(27,intLoop))>0) and (cpnAssignPrice>CLNG(arrList(27,intLoop))) then cpnAssignPrice=CLNG(arrList(27,intLoop))
			cpnVal = arrList(22,intLoop)


			If isEPTYPE = "" Then	'전체EP
				bufstr = ""
				bufstr = bufstr&"<<<begin>>>"&VbCRLF
				bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
				If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then
					bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF   ''소비자가
				End If
				bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF		  ''판매가(쿠폰적용가)
				bufstr = bufstr&"<<<pname>>>[텐바이텐] "&itemname&"_"&jaehu3depNM&VbCRLF
				bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=daumshop"&VbCRLF
				bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
'				bufstr = bufstr&"<<<upimg>>>Y"&VbCRLF
				For lp2=1 to Ubound(CntCD)+1
					If lp2>4 Then Exit For
					bufstr = bufstr&"<<<caid"&lp2&">>>"&CntCD(lp2-1)&VbCRLF
				Next
				For lp=1 to Ubound(CntNM)+1
					If lp>4 Then Exit For
					bufstr = bufstr&"<<<cate"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
				Next
				bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
				bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
				IF cpnVal <> "" THEN
					bufstr = bufstr&"<<<coupo>>>"&Replace(cpnVal,"&nbsp;","")&VbCRLF
					bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
				ELSE
					If isCouponYN = "Y" Then		'쿠폰 유효기간 동안이라면..
						bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
					End If
				End IF

				bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
				bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
				If arrList(15,intLoop) > 0 Then
					bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
				End If
				If (cpnVal <> "" and isCouponYN = "Y") Then		'쿠폰 유효기간 동안이라면..
					bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
				Else
					bufstr = bufstr&"<<<event>>>구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"&VbCRLF
				End If
				bufstr = bufstr&"<<<ftend>>>"
				tFile.WriteLine bufstr
				iLastItemid = itemid
			ElseIf isEPTYPE = "chg" Then				'요약EP
				If arrList(21,intLoop) = "I" Then		'요약 신규EP
					bufstr = ""
					bufstr = bufstr&"<<<begin>>>"&VbCRLF
					bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
					If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then
						bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF
					End If
					bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF
					bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
					bufstr = bufstr&"<<<utime>>>"&FormatDate(arrList(10,intLoop),"00000000000000")&VbCRLF
					bufstr = bufstr&"<<<pname>>>[텐바이텐] "&itemname&"_"&jaehu3depNM&VbCRLF
					bufstr = bufstr&"<<<pgurl>>>"&"http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid&"&rdsite=daumshop"&VbCRLF
					bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
					'bufstr = bufstr&"<<<upimg>>>Y"&VbCRLF
					For lp2=1 to Ubound(CntCD)+1
						If lp2>4 Then Exit For
						bufstr = bufstr&"<<<caid"&lp2&">>>"&CntCD(lp2-1)&VbCRLF
					Next
					For lp=1 to Ubound(CntNM)+1
						If lp>4 Then Exit For
						bufstr = bufstr&"<<<cate"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
					Next
					bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
					bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
					IF cpnVal <> "" THEN
						bufstr = bufstr&"<<<coupo>>>"&Replace(cpnVal,"&nbsp;","")&VbCRLF
						bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
					ELSE
						If isCouponYN = "Y" Then		'쿠폰 유효기간 동안이라면..
							bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
						End If
					end if

					bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
					bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
					If arrList(15,intLoop) > 0 Then
						bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
					End If
					If (cpnVal <> "" and isCouponYN = "Y") Then		'쿠폰 유효기간 동안이라면..
						bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
					Else
						bufstr = bufstr&"<<<event>>>구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"&VbCRLF
					End If
					bufstr = bufstr&"<<<ftend>>>"
					tFile.WriteLine bufstr
					iLastItemid = itemid
				ElseIf arrList(21,intLoop) = "U" Then	'요약 수정EP
					bufstr = ""
					bufstr = bufstr&"<<<begin>>>"&VbCRLF
					bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
					If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then  ''추가함(2018/11/30)
						bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF
					End If
					bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF
					bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
					bufstr = bufstr&"<<<utime>>>"&FormatDate(arrList(10,intLoop),"00000000000000")&VbCRLF
					bufstr = bufstr&"<<<pname>>>[텐바이텐] "&itemname&"_"&jaehu3depNM&VbCRLF
					'bufstr = bufstr&"<<<igurl>>>"&"http://webimage.10x10.co.kr/image/basic/" & GetImageSubFolderByItemid(itemid) + "/" & arrList(4,intLoop)&VbCRLF
					For lp2=1 to Ubound(CntCD)+1
						If lp2>4 Then Exit For
						bufstr = bufstr&"<<<caid"&lp2&">>>"&CntCD(lp2-1)&VbCRLF
					Next
					For lp=1 to Ubound(CntNM)+1
						If lp>4 Then Exit For
						bufstr = bufstr&"<<<cate"&lp&">>>"&Replace(CntNM(lp-1),"&nbsp;","")&VbCRLF
					Next
					bufstr = bufstr&"<<<brand>>>"&Replace(arrList(14,intLoop),"&nbsp;","")&VbCRLF
					bufstr = bufstr&"<<<maker>>>"&Replace(arrList(6,intLoop),"&nbsp;","")&VbCRLF
					IF cpnVal <> "" THEN
						bufstr = bufstr&"<<<coupo>>>"&Replace(cpnVal,"&nbsp;","")&VbCRLF
						bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
					ELSE
						If isCouponYN = "Y" Then		'쿠폰 유효기간 동안이라면..
							bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
						End If
					END IF

					bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
					bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
					If arrList(15,intLoop) > 0 Then
						bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
					End If
					If (cpnVal <> "") and (isCouponYN = "Y") Then		'쿠폰 유효기간 동안이라면..
						bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
					Else
						bufstr = bufstr&"<<<event>>>구매 시 마일리지 적립 & 신규회원 가입 시 보너스쿠폰 증정!"&VbCRLF
					End If
					bufstr = bufstr&"<<<ftend>>>"
					tFile.WriteLine bufstr
					iLastItemid = itemid
				ElseIf arrList(21,intLoop) = "D" Then	'요약 삭제EP
					bufstr = ""
					bufstr = bufstr&"<<<begin>>>"&VbCRLF
					bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
					bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
					bufstr = bufstr&"<<<utime>>>"&FormatDate(arrList(10,intLoop),"00000000000000")&VbCRLF
					bufstr = bufstr&"<<<ftend>>>"
					tFile.WriteLine bufstr
					iLastItemid = itemid
				End If
			End If
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
	ElseIf (isEPTYPE = "Newreview") Then
		If iFTotCnt >= 350000 Then
			bufstr = "<<<tocnt>>>350000"
		Else
			bufstr = "<<<tocnt>>>"&iFTotCnt
		End If
		tFile.WriteLine bufstr
		For intLoop=0 to iRow
			bufstr = ""
			bufstr = bufstr&"<<<begin>>>"&VbCRLF
			bufstr = bufstr&"<<<mapid>>>"&arrList(1,intLoop)&VbCRLF
			bufstr = bufstr&"<<<reviewid>>>"&arrList(2,intLoop)&VbCRLF
			bufstr = bufstr&"<<<status>>>S"&VbCRLF
			If Len(db2html(arrList(3,intLoop))) > 20 Then
				bufstr = bufstr&"<<<title>>>"&Trim(LEFT(db2html(arrList(3,intLoop)),10))&"..."&VbCRLF
			Else
				bufstr = bufstr&"<<<title>>>"&db2html(arrList(3,intLoop))&VbCRLF
			End If
			bufstr = bufstr&"<<<content>>>"&db2html(arrList(3,intLoop))&VbCRLF
			bufstr = bufstr&"<<<writer>>>"&printUserId(arrList(4,intLoop), 2, "*")&VbCRLF
			bufstr = bufstr&"<<<cdate>>>"&FormatDate(arrList(5,intLoop),"00000000000000")&VbCRLF
'			bufstr = bufstr&"<<<rating>>>"&CInt(arrList(6,intLoop)) + 1&"/5"&VbCRLF
			bufstr = bufstr&"<<<rating>>>"&CInt(arrList(6,intLoop))&"/5"&VbCRLF	'2020-08-21 김진영 수정..+1 뺌
			bufstr = bufstr&"<<<ftend>>>"
			tFile.WriteLine bufstr
			iLastItemid = arrList(2,intLoop)
		Next
	ElseIf (isEPTYPE = "Newreview2") Then
		bufstr = "<<<tocnt>>>"&iFTotCnt
		tFile.WriteLine bufstr
		For intLoop=0 to iRow
			bufstr = ""
			bufstr = bufstr&"<<<begin>>>"&VbCRLF
			bufstr = bufstr&"<<<mapid>>>"&arrList(1,intLoop)&VbCRLF
			bufstr = bufstr&"<<<reviewid>>>"&arrList(2,intLoop)&VbCRLF
			bufstr = bufstr&"<<<status>>>S"&VbCRLF
			If Len(db2html(arrList(3,intLoop))) > 20 Then
				bufstr = bufstr&"<<<title>>>"&Trim(LEFT(db2html(arrList(3,intLoop)),10))&"..."&VbCRLF
			Else
				bufstr = bufstr&"<<<title>>>"&db2html(arrList(3,intLoop))&VbCRLF
			End If
			bufstr = bufstr&"<<<content>>>"&db2html(arrList(3,intLoop))&VbCRLF
			bufstr = bufstr&"<<<writer>>>"&printUserId(arrList(4,intLoop), 2, "*")&VbCRLF
			bufstr = bufstr&"<<<cdate>>>"&FormatDate(arrList(5,intLoop),"00000000000000")&VbCRLF
			bufstr = bufstr&"<<<rating>>>"&CInt(arrList(6,intLoop))&"/5"&VbCRLF	'2020-08-21 김진영 수정..+1 뺌
			bufstr = bufstr&"<<<ftend>>>"
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
		Case "chg"			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_ST')"
		Case "best100"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_ST')"
		Case "review"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_ST')"
		Case "Newreview"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV_ST')"
		Case "Newreview2"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV2_ST')"
	End Select
    dbCTget.execute sqlStr
Else
	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_DY_ST')"
	dbCTget.execute sqlStr
End If

If EPTYPE <> "" Then
	Select Case EPTYPE
		Case "chg"			sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]('C')"
		Case "best100"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]('B')"
		Case "review"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]('R')"
		Case "Newreview"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]('M')"
		Case "Newreview2"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]('K')"
	End Select
Else
    sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPDataCount]"
End If
dbCTget.CommandTimeout = 120 ''2019/01/16 추가
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
    IF (FTotPage>MaxPage) THEN
		FTotPage=MaxPage
		FTotCnt=MaxPage*PageSize
	ENd IF

    Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(appPath & FileName )
		    For i=0 to FTotPage-1
		        ArrRows = ""
		        If EPTYPE <> "" Then
					Select Case EPTYPE
						Case "chg"			sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'C',"&iLastItemid&")"
						Case "best100"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'B',"&iLastItemid&")"
						Case "review"		sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'R',"&iLastItemid&")"
						Case "Newreview"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'M',"&iLastItemid&")"
						Case "Newreview2"	sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'K',"&iLastItemid&")"
					End Select
		        Else
		            sqlStr ="[db_outmall].[dbo].[sp_Ten_Daum_NewEPData]("&i+1&","&PageSize&",'A',"&iLastItemid&")"
		        End If
				dbCTget.CommandTimeout = 120 ''2019/01/16 추가
		        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		        	ArrRows = rsCTget.getRows()
		        END IF
		        rsCTget.close

		        If isArray(ArrRows) then
		            CALL WriteMakeDaumFile(tFile, ArrRows, EPTYPE, iLastItemid, FTotCnt)
		        End If

		        ''작성시간 체크
				If EPTYPE <> "" Then
					Select Case EPTYPE
						Case "chg"			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "best100"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "review"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "Newreview"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV_"&(i+1)*PageSize&"_"&iLastItemid&"')"
						Case "Newreview2"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV2_"&(i+1)*PageSize&"_"&iLastItemid&"')"
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
		Case "chg"			sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_CH_ED')"
		Case "best100"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_BT_ED')"
		Case "review"		sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_RV_ED')"
		Case "Newreview"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV_ED')"
		Case "Newreview2"	sqlStr = "insert into [db_outmall].[dbo].tbl_nate_scraplog (ref) values('daumshop_NRV2_ED')"
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