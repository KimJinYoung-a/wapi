<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200  ''�ʴ���
'��ǰEP�� 78�� DB�� �ٶ󺻴�
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
		Case "chg"			'���EP
			FileName = "daumUpdateEP_temp.txt"
			newFileName = "daumUpdateEP.txt"
		Case "best100"		'����Ʈ100EP
			FileName = "daumBest100EP_temp.txt"
			newFileName = "daumBest100EP.txt"
		Case "review"		'��ǰ��EP
			FileName = "daumReviewEP_temp.txt"
			newFileName = "daumReviewEP.txt"
		Case "Newreview"	'�Ź��� ��ǰ��EP
			FileName = "daumNewReviewEP_temp.txt"
			newFileName = "daumNewReviewEP.txt"
		Case "Newreview2"	'�Ź��� ��ǰ��EP (��ü���޿�)
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
	Dim cpnAssignPrice, cpnVal ''2018/11/30 �߰�

	iRow = UBound(arrList,2)

	If (Now() > #05/22/2016 21:00:00# AND Now() < #05/29/2016 20:59:59#) Then
		isCouponYN = "Y"
		iCouponEvtText = "30,000�� �̻� ���� �� 3,000�� ���� ���� ���� & ���ϸ��� ����"
	End If

	If (Now() > #10/01/2018 00:00:00# AND Now() < #10/01/2018 21:59:59#) Then
		isCouponYN = "Y"
		iCouponEvtText = "10/1�� ���� �� �Ϸ縸! �ִ� 3���� ��������"
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
			deliv 			= arrList(19,intLoop)  ''��ۺ� /2000, 2500, 0

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

	    		'2������ 2�������� ������ ����..2017-10-17 ������
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


			If isEPTYPE = "" Then	'��üEP
				bufstr = ""
				bufstr = bufstr&"<<<begin>>>"&VbCRLF
				bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
				If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then
					bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF   ''�Һ��ڰ�
				End If
				bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF		  ''�ǸŰ�(�������밡)
				bufstr = bufstr&"<<<pname>>>[�ٹ�����] "&itemname&"_"&jaehu3depNM&VbCRLF
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
					If isCouponYN = "Y" Then		'���� ��ȿ�Ⱓ �����̶��..
						bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
					End If
				End IF

				bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
				bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
				If arrList(15,intLoop) > 0 Then
					bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
				End If
				If (cpnVal <> "" and isCouponYN = "Y") Then		'���� ��ȿ�Ⱓ �����̶��..
					bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
				Else
					bufstr = bufstr&"<<<event>>>���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"&VbCRLF
				End If
				bufstr = bufstr&"<<<ftend>>>"
				tFile.WriteLine bufstr
				iLastItemid = itemid
			ElseIf isEPTYPE = "chg" Then				'���EP
				If arrList(21,intLoop) = "I" Then		'��� �ű�EP
					bufstr = ""
					bufstr = bufstr&"<<<begin>>>"&VbCRLF
					bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
					If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then
						bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF
					End If
					bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF
					bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
					bufstr = bufstr&"<<<utime>>>"&FormatDate(arrList(10,intLoop),"00000000000000")&VbCRLF
					bufstr = bufstr&"<<<pname>>>[�ٹ�����] "&itemname&"_"&jaehu3depNM&VbCRLF
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
						If isCouponYN = "Y" Then		'���� ��ȿ�Ⱓ �����̶��..
							bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
						End If
					end if

					bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
					bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
					If arrList(15,intLoop) > 0 Then
						bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
					End If
					If (cpnVal <> "" and isCouponYN = "Y") Then		'���� ��ȿ�Ⱓ �����̶��..
						bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
					Else
						bufstr = bufstr&"<<<event>>>���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"&VbCRLF
					End If
					bufstr = bufstr&"<<<ftend>>>"
					tFile.WriteLine bufstr
					iLastItemid = itemid
				ElseIf arrList(21,intLoop) = "U" Then	'��� ����EP
					bufstr = ""
					bufstr = bufstr&"<<<begin>>>"&VbCRLF
					bufstr = bufstr&"<<<mapid>>>"&itemid&VbCRLF
					If CLNG(cpnAssignPrice) < CLNG(arrList(25,intLoop)) Then  ''�߰���(2018/11/30)
						bufstr = bufstr&"<<<lprice>>>"&CLNG(arrList(25,intLoop))&VbCRLF
					End If
					bufstr = bufstr&"<<<price>>>"&cpnAssignPrice&VbCRLF
					bufstr = bufstr&"<<<class>>>"&arrList(21,intLoop)&VbCRLF
					bufstr = bufstr&"<<<utime>>>"&FormatDate(arrList(10,intLoop),"00000000000000")&VbCRLF
					bufstr = bufstr&"<<<pname>>>[�ٹ�����] "&itemname&"_"&jaehu3depNM&VbCRLF
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
						If isCouponYN = "Y" Then		'���� ��ȿ�Ⱓ �����̶��..
							bufstr = bufstr&"<<<cpdown>>>Y"&VbCRLF
						End If
					END IF

					bufstr = bufstr&"<<<point>>>"&arrList(11,intLoop)&VbCRLF
					bufstr = bufstr&"<<<deliv>>>"&deliv&VbCRLF
					If arrList(15,intLoop) > 0 Then
						bufstr = bufstr&"<<<revct>>>"&arrList(15,intLoop)&VbCRLF
					End If
					If (cpnVal <> "") and (isCouponYN = "Y") Then		'���� ��ȿ�Ⱓ �����̶��..
						bufstr = bufstr&"<<<event>>>"&iCouponEvtText&VbCRLF
					Else
						bufstr = bufstr&"<<<event>>>���� �� ���ϸ��� ���� & �ű�ȸ�� ���� �� ���ʽ����� ����!"&VbCRLF
					End If
					bufstr = bufstr&"<<<ftend>>>"
					tFile.WriteLine bufstr
					iLastItemid = itemid
				ElseIf arrList(21,intLoop) = "D" Then	'��� ����EP
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
			bufstr = bufstr&"<<<rating>>>"&CInt(arrList(6,intLoop))&"/5"&VbCRLF	'2020-08-21 ������ ����..+1 ��
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
			bufstr = bufstr&"<<<rating>>>"&CInt(arrList(6,intLoop))&"/5"&VbCRLF	'2020-08-21 ������ ����..+1 ��
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

''�ۼ��ð� üũ
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
dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
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
				dbCTget.CommandTimeout = 120 ''2019/01/16 �߰�
		        rsCTget.Open sqlStr, dbCTget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		        IF Not (rsCTget.EOF OR rsCTget.BOF) THEN
		        	ArrRows = rsCTget.getRows()
		        END IF
		        rsCTget.close

		        If isArray(ArrRows) then
		            CALL WriteMakeDaumFile(tFile, ArrRows, EPTYPE, iLastItemid, FTotCnt)
		        End If

		        ''�ۼ��ð� üũ
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

''�ۼ��ð� üũ
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
response.write FTotCnt&"�� ���� ["&FileName&"]"
%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->