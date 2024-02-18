<%
CONST CMAXMARGIN = 18
CONST CMALLNAME = "11st1010"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST APIURL = "http://api.11st.co.kr/rest"
CONST APISSLURL = "https://api.11st.co.kr/rest"
CONST APIkey = "a2319e071dbc304243ee60abd07e9664"
CONST CDEFALUT_STOCK = 99999

Class C11stItem
	Public Fitemid
	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public Fitemname
	Public FitemDiv
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public ForgSuplyCash
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FisUsing
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public Fkeywords
	Public Fvatinclude
	Public ForderComment
	Public FoptionCnt
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Ficon2Image
	Public Fsourcearea
	Public Fmakername
	Public FBrandName
	Public FBrandNameKor
	Public FItemsize
	Public FItemsource
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FSt11StatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public Fsocname_kor
	Public FDepthCode
	Public Fcdmkey
	Public Fcddkey
	Public FSt11GoodNo
	Public FSt11price
	Public FSt11SellYn

	Public FSafeDiv
	Public FIsNeed
	Public FDepth1Code
	Public FAdultType
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
		sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

				If (optLimit <> 0) Then
					limitYCnt =  limitYCnt + 1
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If limitYCnt = 0 Then
			IsMayLimitSoldout = "Y"
		Else
			IsMayLimitSoldout = "N"
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin
		Dim ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT m.mustPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as m "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE m.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and m.itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= m.startDate and getdate() <= m.endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice			= rsget("mustPrice")
			outmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
		sqlStr = sqlStr & " FROM db_partner.dbo.tbl_partner "
		sqlStr = sqlStr & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : ����
		sqlStr = sqlStr & " and id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ownItemCnt = rsget("CNT")
		End If
		rsget.Close

		If specialPrice <> "" Then
			tmpPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			tmpPrice = Forgprice
		Else
			If outmallstandardMargin = "" Then
				outmallstandardMargin	= FOutmallstandardMargin
			End If

			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If (GetTenTenMargin < outmallstandardMargin) Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'�ִ� ���� ����
	Public Function getLimit11stEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 1000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimit11stEa = ret
	End Function

	'// 11st �Ǹſ��� ��ȯ
	Public Function get11stSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				get11stSellYn = "Y"
			Else
				get11stSellYn = "N"
			End If
		Else
			get11stSellYn = "N"
		End If
	End Function

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST��ǰ] "&FItemName
		Else
			buf = "[�ٹ�����] "&FItemName
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"&","��")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
        buf = LeftB(buf, 140)
        getItemNameFormat = buf
    end function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

	Public Function checkTenItemOptionValid()
		Dim strSql, chkRst, chkMultiOpt
		Dim cntType, cntOpt
		chkRst = true
		chkMultiOpt = false

		If FoptionCnt > 0 Then
			'// ���߿ɼ�Ȯ��
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				chkMultiOpt = true
				cntType = rsget.RecordCount
			End If
			rsget.Close

			If chkMultiOpt Then
				'// ���߿ɼ� �϶�
				strSql = "Select optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly

				If Not(rsget.EOF or rsget.BOF) Then
					Do until rsget.EOF
						cntOpt = ubound(split(db2Html(rsget("optionname")), ",")) + 1
						If cntType <> cntOpt then
							chkRst = false
						End If
						rsget.MoveNext
					Loop
				Else
					chkRst = false
				End If
				rsget.Close
			Else
				'// ���Ͽɼ��� ��
				strSql = "Select optionTypeName, optionname "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where itemid=" & FItemid
				strSql = strSql & " 	and isUsing='Y' and optsellyn='Y' "
				strSql = strSql & " 	and optaddprice=0 "
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>="&CMAXLIMITSELL&")) "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If (rsget.EOF or rsget.BOF) Then
					chkRst = false
				End If
				rsget.Close
			End If
		End If
		'//��� ��ȯ
		checkTenItemOptionValid = chkRst
	End Function

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
' o11st.FOneItem.FLimityn
		i = 0
		sqlStr = ""
		sqlStr = sqlStr & "SELECT Count(*) as cnt FROM db_item.dbo.tbl_item_option where itemid = '"&iitemid&"' and optaddprice > 0 "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			cnt = rsget("cnt")
		rsget.Close

		If cnt = 0 Then
			getiszeroWonSoldOut = "N"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " SELECT itemid, itemoption, optlimitno, optlimitsold "
			sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option  "
			sqlStr = sqlStr & " where itemid = '"&iitemid&"'  "
			sqlStr = sqlStr & " and optaddprice = 0 "
			sqlStr = sqlStr & " and isusing = 'Y' "
			sqlStr = sqlStr & " and optsellyn = 'Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				Do until rsget.EOF
					goptlimitno		= rsget("optlimitno")
					goptlimitsold	= rsget("optlimitsold")
					If goptlimitno - goptlimitsold > CMAXLIMITSELL Then
						i = i + 1
					End If
					rsget.MoveNext
				Loop

				If i = 0 Then		'0�� �ɼ��� ��� 5�� ���ϸ� ǰ��
					getiszeroWonSoldOut = "Y"
				Else
					getiszeroWonSoldOut = "N"
				End If
			Else
				getiszeroWonSoldOut = "Y"
			End If
			rsget.Close
		End If
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function get11stContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_11st.jpg'></p><br>"

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

		If Fitemsize <> "" Then
			strRst = strRst & "- ������ : " & Fitemsize & "<br>"
		End if

		If Fitemsource <> "" Then
			strRst = strRst & "- ��� : " &  Fitemsource & "<br>"
		End If
		strRst = strRst & Replace(Replace(FItemContent,"",""),"","")

		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ ><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ ><br>")

		'#��� ���ǻ���
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_11st.jpg"">")
		get11stContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = ""
			strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_11st.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_11st.jpg"">")
			get11stContParamToReg = strRst
		End If
		rsget.Close
	End Function

	'// �˻���
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
		If trim(Fkeywords) = "" Then Exit Function
		Fkeywords  = replace(Fkeywords,"%", "")
		Fkeywords  = replace(Fkeywords,"/", ",")
		Fkeywords  = replace(Fkeywords,chr(13), "")
		Fkeywords  = replace(Fkeywords,chr(10), "")
		Fkeywords  = replace(Fkeywords,chr(9), "")
		Fkeywords  = replace(Fkeywords,chr(32), "")

		arrRst = Split(Fkeywords,",")
		If Ubound(arrRst) = 0 then
			arrRst2 = split(arrRst(0),";")
			If Ubound(arrRst2) > 0 then
				arrRst = split(Fkeywords,";")
			End If
		End If

		If Ubound(arrRst)+1 >= 5 then
			getItemKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			getItemKeyword = keyword1
		End If
	End Function

    public function isImageChanged()
        Dim ibuf : ibuf = getBasicImage
        if InStr(ibuf,"-")<1 then
            isImageChanged = FALSE
            Exit function
        end if
        isImageChanged = ibuf <> FregImageName
    end function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	Public Function get11stAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		strRst = strRst & "	<prdImage01>"&FbasicImage&"</prdImage01>"					'#��ǥ �̹��� URL | �̹����� 11���� ������ �ٿ�ε��Ͽ� 300 x 300 ������� ������¡ �ѵ� 11���� �̹��������� ���� �˴ϴ�. �̹��� Ȯ���ڴ� gif, jpg, jpeg, png �� ��밡���մϴ�. �̹��� url ȣ��� "Content-Type" �� ���ǰ� �Ǿ����� ������ �̹��� �ٿ�ε尡 �̷�� �����ʽ��ϴ�.
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "	<prdImage0"&i+1&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"</prdImage0"&i+1&">"					'�߰� �̹��� 1 URL
				End If
				rsget.MoveNext
				If i>=3 Then Exit For
			Next
		End If
		rsget.Close
'		strRst = strRst & "	<prdImage05/>"					'����̹��� | �˻� ��� �������� ī�װ� ����Ʈ ���������� ����Ǵ� �̹����Դϴ�.
'		strRst = strRst & "	<prdImage09/>"					'ī����̹��� | ��ŷ��/��ȹ���� ī��� ���� �̹����Դϴ�.
'		strRst = strRst & "	<prdImage01Src/>"				'�̹��� ����Ʈ �ڵ� | ����Ʈ �ڵ�� ��ȯ�Ͽ� �����ž� �մϴ�.
		get11stAddImageParam = strRst
	End Function

	Public Function get11stSafeParam()
		Dim strRst, certTypeCd, strSql, arrRows, notarrRows, nlp, newDiv, newCertNo
		If FSafeDiv = "2" Then

			strSql = ""
			strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				arrRows = rsget.getRows()
			Else
				notarrRows = "Y"
			End If
			rsget.Close

			If notarrRows = "" Then		'���ȹ� ����� �����Ͷ�� ����� �ű�
				If FsafetyYn = "Y" Then
					For nLp =0 To UBound(arrRows,2)
				    	newDiv = ""
						Select Case arrRows(1,nLp)
							Case "10"		newDiv = "102"		'�����ǰ > ��������
							Case "20"		newDiv = "104"		'�����ǰ > ����Ȯ�� �Ű�
							Case "30"		newDiv = "127"		'�����ǰ > ������ ���ռ� Ȯ��
							Case "40"		newDiv = "101"		'��Ȱ��ǰ > ��������
							Case "50"		newDiv = "103"		'��Ȱ��ǰ > ��������Ȯ��
							Case "60"		newDiv = "124"		'��Ȱ��ǰ > ����ǰ��ǥ��
							Case "70"		newDiv = "128"		'�����ǰ > ��������
							Case "80"		newDiv = "129"		'�����ǰ > ����Ȯ��
							Case "90"		newDiv = "130"		'�����ǰ > ������ ���ռ� Ȯ��
						End Select

						newCertNo = arrRows(0,nLp)
						If newCertNo = "x" Then
							newCertNo = ""
						End If

						strRst = strRst & "	<ProductCert>"
						strRst = strRst & "		<certTypeCd>"&newDiv&"</certTypeCd>"
						strRst = strRst & "		<certKey><![CDATA["&newCertNo&"]]></certKey>"				'������ȣ
						strRst = strRst & "	</ProductCert>"
					Next
				Else
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>132</certTypeCd>"								'#132 : [�����ǰ/��Ȱ��ǰ] ��ǰ�󼼼��� ����
					strRst = strRst & "		<certKey/>"
					strRst = strRst & "	</ProductCert>"

				End If
			Else
				If FsafetyYn = "Y" AND FSafetyNum <> "" Then
					Select Case FsafetyDiv
						Case "10"	certTypeCd = "101"													'[����ǰ] ��������
						Case "20"	certTypeCd = "102"													'[�����ǰ] ��������
						Case "30"	certTypeCd = "124"													'[����ǰ] ����/ǰ��ǥ��
						Case "40"	certTypeCd = "103"													'[����ǰ] ��������Ȯ��
						Case "50"	certTypeCd = "123"													'[����ǰ] ��̺�ȣ����
					End Select
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>"&certTypeCd&"</certTypeCd>"
					strRst = strRst & "		<certKey><![CDATA["&FSafetyNum&"]]></certKey>"				'������ȣ
					strRst = strRst & "	</ProductCert>"
				Else
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>132</certTypeCd>"							'#132 : [�����ǰ/��Ȱ��ǰ] ��ǰ�󼼼��� ����
					strRst = strRst & "		<certKey/>"
					strRst = strRst & "	</ProductCert>"
				End If
			End If
		Else
			'strRst = strRst & "	<ProductCert>"
			'strRst = strRst & "		<certTypeCd>131</certTypeCd>"								'#131 : �ش����(����� �ƴ� ���)..����ȵ� 131�ڵ� �����
			'strRst = strRst & "		<certKey/>"
			'strRst = strRst & "	</ProductCert>"
		End If
		get11stSafeParam = strRst
	End Function

	Public Function get11stSafeNewParam()
		Dim strRst, certTypeCd, strSql, arrRows, notarrRows, nlp, newDiv, newCertNo, i, crtfGrpObjClfCd
		Dim crtfGrpTypCd, certNum, safetyDiv, certKey
		If FSafeDiv = "2" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 1 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			strSql = strSql & " ORDER BY regdate DESC "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				certNum = rsget("certNum")
				safetyDiv = rsget("safetyDiv")
				If certNum = "x" Then
					certNum = ""
				End If
			Else
				certNum = ""
				safetyDiv = ""
			End If
			rsget.Close

			strRst = ""
			For i = 1 to 4
				Select Case i
					Case "1"			crtfGrpTypCd = "01"				'01: �����ǰ/��Ȱ��ǰ KC����
					Case "2"			crtfGrpTypCd = "02"				'02: �����ǰ KC����
					Case "3"			crtfGrpTypCd = "03"				'03 : �����ű����� KC����
					Case "4"			crtfGrpTypCd = "04"				'04 : ��Ȱȭ�� �� �������ǰ
				End Select

				crtfGrpObjClfCd = ""
				'crtfGrpObjClfCd(01) : KC�������, crtfGrpObjClfCd(03 or 05) : 03 KC�������ƴ� / 05 ��Ȱȭ�� �� �������ǰ ��� �ƴ�
				Select Case crtfGrpTypCd
					Case "01"
						If (safetyDiv = "10") OR (safetyDiv = "20") OR (safetyDiv = "40") OR (safetyDiv = "50") Then
							crtfGrpObjClfCd = "01"
						End If
					Case "02"
						If (safetyDiv = "70") OR (safetyDiv = "80") Then
							crtfGrpObjClfCd = "01"
						End If
					Case "04"
						crtfGrpObjClfCd = "05"
					Case Else
						crtfGrpObjClfCd = "03"
				End Select

				'���� ��ȭ���� �ƴϰ� ������ȣ�� ������ 03 ó��
				If crtfGrpTypCd <> "04" AND certNum = "" Then
					crtfGrpObjClfCd = "03"
				End If

				If crtfGrpObjClfCd = "" Then
					crtfGrpObjClfCd = "03"
				End If

				Select Case safetyDiv
					Case "10"		newDiv = "102"		'�����ǰ > ��������
					Case "20"		newDiv = "104"		'�����ǰ > ����Ȯ�� �Ű�
					Case "30"		newDiv = "127"		'�����ǰ > ������ ���ռ� Ȯ��
					Case "40"		newDiv = "101"		'��Ȱ��ǰ > ��������
					Case "50"		newDiv = "103"		'��Ȱ��ǰ > ��������Ȯ��
					Case "60"		newDiv = "124"		'��Ȱ��ǰ > ����ǰ��ǥ��
					Case "70"		newDiv = "128"		'�����ǰ > ��������
					Case "80"		newDiv = "129"		'�����ǰ > ����Ȯ��
					Case "90"		newDiv = "130"		'�����ǰ > ������ ���ռ� Ȯ��
				End Select

				If crtfGrpTypCd = "01" AND ((safetyDiv = "10") OR (safetyDiv = "20") OR (safetyDiv = "40") OR (safetyDiv = "50")) Then
					certTypeCd = newDiv
					certKey = certNum
				ElseIf crtfGrpTypCd = "02" AND ((safetyDiv = "70") OR (safetyDiv = "80")) Then
					certTypeCd = newDiv
					certKey = certNum
				Else
					certTypeCd = ""
					certKey = ""
				End If

				strRst = strRst & "	<ProductCertGroup>"												'���������׷�
				strRst = strRst & "		<crtfGrpTypCd>"&crtfGrpTypCd&"</crtfGrpTypCd>"				'���������׷��ȣ | ���������׷��ȣ�� �������� �ʴ� ��ǰ ������ ���, �ش� ���� �Է����� �ʽ��ϴ�. �����ǰ/��Ȱ��ǰ, �����ǰ, �����ű�����, ��Ȱȭ�� �׹� �������ǰ�� ���� �������� �Է��� �ʼ��� ī�װ��� ��� 01, 02, 03, 04�� ���������� ��� �Է����ּ���. �� 01 : �����ǰ/��Ȱ��ǰ KC���� �� 02 : �����ǰ KC���� �� 03 : �����ű����� KC���� �� 04 : ��Ȱȭ�� �� �������ǰ
				strRst = strRst & "		<crtfGrpObjClfCd>"&crtfGrpObjClfCd&"</crtfGrpObjClfCd>"		'KC������󿩺� | ���������׷��ȣ�� 01, 02, 03, 04�� ��� ������󿩺� ���� �ʼ� �Է��ؾ� �մϴ�. (���������׷��ȣ 01 : ������󿩺� 01, 02, 03 �� 1 ��� ���� / ���������׷��ȣ 02 : ������󿩺� 01, 03 �� 1 ��� ���� / ���������׷��ȣ 03 : ������󿩺� 01, 03 �� 1 ��� ���� / ���������׷��ȣ 04 : ������󿩺� 04, 05 �� 1 ��� ����) �� 01 : KC������� �� 02 : KC������� �� 03 : KC������� �ƴ� �� 04 : ��Ȱȭ�� �� �������ǰ ��� �� 05 : ��Ȱȭ�� �� �������ǰ ��� �ƴ�
	'			strRst = strRst & "		<crtfGrpExptTypCd></crtfGrpExptTypCd>"						'KC�������� | KC������󿩺ΰ� 02�� ��� KC�������� ���� �ʼ� �Է��ؾ� �մϴ�. �� 02 : ���Ŵ��������� �� 03 : ������Ը������
				strRst = strRst & "		<ProductCert>"												'�������� | ���������� �ִ� 100�� ������ �����մϴ�.
				If certTypeCd <> "" Then
					strRst = strRst & "			<certTypeCd>"&certTypeCd&"</certTypeCd>"			'��������
				Else
					strRst = strRst & "			<certTypeCd></certTypeCd>"							'��������
				End If

				If certKey <> "" Then
					strRst = strRst & "			<certKey><![CDATA["&certKey&"]]></certKey>"			'������ȣ
				Else
					strRst = strRst & "			<certKey></certKey>"								'������ȣ
				End If
				strRst = strRst & "		</ProductCert>"
				strRst = strRst & "	</ProductCertGroup>"
			Next
		End If
		get11stSafeNewParam = strRst
	End Function

	Public Function get11stOptionParam()
		Dim strSql, strRst1, strRst2, strRst3, i, optNm, optDc, chkMultiOpt, optLimit, arrList1, arrList2, optCode, j, optionname, optaddprice, tmpoName
		Dim voptsellyn, visusing
		strRst1 = ""
		strRst2 = ""
		strRst3 = ""
		chkMultiOpt = false
		If FoptionCnt > 0 Then
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql,dbget,1
		    if not rsget.Eof then
		    	chkMultiOpt = true
		        arrList1 = rsget.getRows()
		    end if
		    rsget.close
		End If
		strRst1 = strRst1 & "	<optSelectYn>Y</optSelectYn>"									'������ �ɼ� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
		strRst1 = strRst1 & "	<txtColCnt>1</txtColCnt>"										'������ | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ɼ��� ����Ͻ� ��� 1 �������� �ּž� �մϴ�.
'		If chkMultiOpt = True Then
'		strRst1 = strRst1 & "	<optionAllQty>"&getLimit11stEa&"</optionAllQty>"				'��Ƽ�ɼ� �ϰ������� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� ��� ���� ������ api ������ �Ұ��մϴ�. �ϰ������� ����.
'		strRst1 = strRst1 & "	<optionAllAddPrc>0</optionAllAddPrc>"							'��Ƽ�ɼ� �ɼǰ� 0�� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� �ɼǰ� ������ api ������ �Ұ��մϴ�. 0���� �Է� ����
'		strRst1 = strRst1 & "	<optionAllAddWght/>"											'��Ƽ�ɼ� �ϰ��ɼ��߰����� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "��ǰ�� �ɼǰ� ���� ��� ����"�� �����Ͻ� ��� ��ϼ� �ɼ��� ����˴ϴ�. "��Ƽ�ɼ�" ����� �ƴ� "�̱ۿɼ�" ��� �� ���� Element�� �������ּž� �մϴ�. ��Ƽ�ɼ��� �ɼǺ� �ɼǹ��� ������ api ������ �Ұ��մϴ�. �ϰ������� ����.
'		End If
		strRst1 = strRst1 & "	<prdExposeClfCd>00</prdExposeClfCd>"							'��ǰ�� �ɼǰ� ���� ��� ���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 00 : ��ϼ�, 01 : �ɼǰ� �����ټ�, 02 : �ɼǰ� ������ ����, 03 : �ɼǰ��� ���� ��, 04 : �ɼǰ��� ���� ��
		strRst1 = strRst1 & "	<optUpdateYn>Y</optUpdateYn>"
'		strRst1 = strRst1 & "	<optMixYn/>"													'��ü�ɼ� ���տ��� |"�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������ ��ü �ɼǰ��� ���յǾ� ��Ƽ�ɼ����� ���, N : �ɼ� ����Key�� �����ϴ�(���� ��) �����θ� ��Ƽ �ɼ� ���
		If chkMultiOpt = True Then
			strSql = "select typeseq, optionTypeName, optionKindName, optaddPrice from db_item.[dbo].[tbl_item_option_Multiple] where itemid = " & FItemid & " ORDER BY Typeseq, kindSeq "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				arrList2 = rsget.getRows()
			End If
			rsget.close

			For i = 0 To Ubound(arrList1, 2)
				strRst2 = strRst2 & "	<ProductRootOption>"											'ProductRootOption |"�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
				strRst2 = strRst2 & "		<colTitle>"&arrList1(2,i)&"</colTitle>"						'�ɼǸ� | 40Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�.
				For j = 0 to Ubound(arrList2, 2)
					If arrList1(1,i) = arrList2(0,j) then

						arrList2(2,j) = replace(arrList2(2,j), "&", "+")			'2017-06-05 ������..���� �븮 ��û���� &->+�� ����

						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<colOptPrice>0</colOptPrice>"					'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "			<colValue0>"&arrList2(2,j)&"</colValue0>"		'�ɼǰ� | 50Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�. �� ��ǰ�ȿ��� �ɼǰ��� �ߺ��� �ɼ� �����ϴ�.
						strRst2 = strRst2 & "		</ProductOption>"
					end if
				Next
				strRst2 = strRst2 & "	</ProductRootOption>"
			Next

			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, isUsing, optsellyn "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
'			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "		'2017-05-17 ������ ����
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				strRst2 = strRst2 & "	<ProductOptionExt>"												'ProductOptionExt | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    tmpoName	= db2html(rsget("optionname"))
				    visUsing= rsget("isUsing")
				    voptsellyn= rsget("optsellyn")

					optionname = ""
					For i = 0 To Ubound(arrList1, 2)
						If Ubound(Split(tmpoName, ",")) > 0 Then				'2017-06-15 ������ �߰�
							optionname = optionname & arrList1(2,i) &":"&Split(tmpoName, ",")(i) &"��"
						End If
					Next

					If Right(optionname,1) = "��" Then
						optionname = Left(optionname, Len(optionname) - 1)
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
			    	If voptsellyn <> "Y" Then optLimit = 0	'2017-05-17 ������ ����
			    	If visUsing <> "Y" Then optLimit = 0	'2017-05-17 ������ ����

			'		If optLimit > 0 Then		'2017-05-17 ������ ����
						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<useYn>"&chkiif(optLimit>0,"Y","N")&"</useYn>"
						strRst2 = strRst2 & "			<colOptPrice>"&optaddprice&"</colOptPrice>"				'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "			<colCount>"&optLimit&"</colCount>"			'�ɼ������� | ��Ƽ�ɼ��� ���� �ϰ������� �ǹǷ� �Է��Ͻø� �ȵ˴ϴ�. �ɼǻ���(useYn)�� N�� ���� 0 �Է� �����մϴ�.
						strRst2 = strRst2 & "			<colSellerStockCd>"&optCode&"</colSellerStockCd>"		'��������ȣ | ������ ����ϴ� ����ȣ
						strRst2 = strRst2 & "			<optionMappingKey>"&optionname&"</optionMappingKey>"	'�ɼǸ���Key | ��Ƽ�ɼ��� ���յ� �ɼ��� �����ϱ� ���� Key(��: �ɼǸ�1:�ɼǰ�1�ӿɼǸ�2:�ɼǰ�2)
						strRst2 = strRst2 & "		</ProductOption>"
			'		End If
					rsget.MoveNext
				Loop
				strRst2 = strRst2 & "	</ProductOptionExt>"
			End If
			rsget.Close
		Else
			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				If db2Html(rsget("optionTypeName"))<>"" Then
					optNm = Replace(db2Html(rsget("optionTypeName")),":","")
				Else
					optNm = "�ɼ�"
				End If
				strRst1 = strRst1 & "	<colTitle>"&optNm&"</colTitle>"
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    optionname	= db2html(rsget("optionname"))
				    optionname = replace(optionname, "&", "+")			'2017-06-05 ������..���� �븮 ��û���� &->+�� ����
				    optionname = replace(optionname, ",", "+")
				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					If optLimit > 0 Then
						strRst2 = strRst2 & "	<ProductOption>"								'ProductOption | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
						strRst2 = strRst2 & "		<useYn>Y</useYn>"							'�ɼǻ��� | ��Ƽ�ɼ��� ���� �������� �ʴ� ����Դϴ�. Y : �����, N : ǰ��
						strRst2 = strRst2 & "		<colOptPrice>"&optaddprice&"</colOptPrice>"	'�ɼǰ� | �⺻ �ǸŰ��� +100%/-50%���� �����Ͻ� �� �ֽ��ϴ�. �ɼǰ����� 0���� ��ǰ�� �ݵ�� 1�� �̻� �־�� �մϴ�.
						strRst2 = strRst2 & "		<colValue0>"&optionname&"</colValue0>"		'�ɼǰ� | 50Byte ������ �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�. �� ��ǰ�ȿ��� �ɼǰ��� �ߺ��� �ɼ� �����ϴ�
						strRst2 = strRst2 & "		<colCount>"&optLimit&"</colCount>"			'�ɼ������� | ��Ƽ�ɼ��� ���� �ϰ������� �ǹǷ� �Է��Ͻø� �ȵ˴ϴ�. �ɼǻ���(useYn)�� N�� ���� 0 �Է� �����մϴ�.
						strRst2 = strRst2 & "		<colSellerStockCd>"&optCode&"</colSellerStockCd>"	'��������ȣ | ������ ����ϴ� ����ȣ
						strRst2 = strRst2 & "	</ProductOption>"
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If

		If FitemDiv = "06" Then
			strRst3 = strRst3 & "	<ProductCustOption>"										'�ɼǵ�� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �������ۼ��� �ɼ��� ��� �ִ� 5������ ��� ����
			strRst3 = strRst3 & "		<colOptName>�ؽ�Ʈ�� �Է��ϼ���</colOptName>"				'������ �ۼ��� �ɼǸ� | �ɼǸ� �ִ� �������� �ѱ�10��/����(����)20��)���� �Է°����ϸ� Ư�� ����[',",%,&,<,>,#,��]�� �Է��� �� �����ϴ�.
			strRst3 = strRst3 & "		<colOptUseYn>Y</colOptUseYn>"							'�ɼ� ��� ���� | Y : �����, N : ������
			strRst3 = strRst3 & "	</ProductCustOption>"
		End If
		get11stOptionParam = strRst1 & strRst2 & strRst3
'rw get11stOptionParam
'response.end
	End Function

	'�⺻���� ��� XML
	Public Function get11stItemRegParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
		strRst = strRst & "<Product>"
'		strRst = strRst & "	<abrdBuyPlace/>"												'�ؿܻ�ǰ�ڵ�
'		strRst = strRst & "	<abrdSizetableDispYn/>"											'�ؿܻ����� ����ǥ ���⿩��
'		strRst = strRst & "	<selMnbdNckNm><![CDATA[�ٹ�����]]></selMnbdNckNm>"				'�г��� | Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���. �г����� �Է����� ������ ��ǥ �г����� �ڵ����� ��ϵ˴ϴ�. @ �ٹ����� �Է��ϸ� �����߻�
		strRst = strRst & "	<selMthdCd>01</selMthdCd>"										'#�ǸŹ�� | 01 : �������Ǹ�, 02 : ������, 03 : ������, 04 : �����Ǹ�, 05 : �߰��Ǹ�
		strRst = strRst & "	<dispCtgrNo>"&FDepthCode&"</dispCtgrNo>"						'#ī�װ���ȣ | ������ ī�װ��� �Է°����մϴ�. ��ī�װ��� �Է��ϼž� �ϸ� ��ī�װ��� ���� ��� ��ī�װ��� �Է��ϼž� �մϴ�. ī�װ���ȣ ��ȸ ���񽺸� �̿��Ͽ� �ǽð� ��ȸ�� �����մϴ�. ī�װ� ������ ��ī�װ� (Ȥ�� ��ī�װ�) ������ �����մϴ�. ��ī�װ�, ��ī�װ��� �����ϰ��� �� ��� ��ǰ�� ���� ������ּ���.
		strRst = strRst & "	<prdTypCd>01</prdTypCd>"										'#���� ��ǰ �ڵ� | ���� ī�װ� ���� �� Ÿ���� �����Ͽ� ����� �� �ֽ��ϴ�. ���� ���޻� ��ǰ�� ���� ��ǰ���ó�� API�� �̿��� �ֽñ� �ٶ��ϴ�. 01 : �Ϲݹ�ۻ�ǰ, 13 : ���޻� �����ǰ
'		strRst = strRst & "	<hsCode/>"														'?#H.S Code | ���ѹα� ����û�� �Ű�Ǵ� HSCode �Դϴ�. �ؿܼ��� ���չ�� ��ǰ��, ���¸�����ۻ�ǰ�� ��쿡�� �ʼ� �����մϴ�. SO Ȥ�� PO�� ��ǰ��� ī�װ����� �⺻ HSCode�� Ȯ�� �Ͻ� �� �ֽ��ϴ�. ��ǰ�� ���ݿ� �´� HSCode�� �����ϼž� �ϸ�, �߸� ����Ͽ� ������ �߻��� ��� �����в��� �ذ��ϼž� �մϴ�. �Ʒ� ÷�������� ���� ��Ź �帳�ϴ�.
		strRst = strRst & "	<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"				'#��ǰ�� | Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���. ���ڼ��� 50Byte�� ���� �� �����Դϴ�. �ѱ� 25��, ����/���� 50�� �̳��� �Է��� �����մϴ�. �Է��� �Ұ��� Ư�����ڰ� ���Ե� ���, �ش� ���ڴ� ��ǰ���� �ڵ� �̳���ó�� �˴ϴ�. [�ڼ�������]
'		strRst = strRst & "	<prdNmEng/>"													'���� ��ǰ��
'		strRst = strRst & "	<engDispYn/>"													'11���� ���� ���� | Y : ����, N : �����
'		strRst = strRst & "	<advrtStmt/>"													'��ǰȫ������ | Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���. ���ڼ��� 40Byte �� ���ѵ˴ϴ�. �ѱ� 20��, ����/���� 40�� �̳��� �Է��� �ֽʽÿ�.
		strRst = strRst & "	<brand><![CDATA[�ٹ�����]]></brand>"								'�귣�� | �귣�带 ��Ȯ�� �Է��ϸ� �ش� ��ǰ�� �˻� ������ �� �������ϴ�. �귣��� �ؽ�Ʈ ���·θ� �Է� �귣�� ���� ���񽺿� ���ø� ���ؼ��� �귣����� ��Ȯ�� �Է��� �ּž� �մϴ�. Ư�� ���縵�� �������ּ���.
		strRst = strRst & "	<rmaterialTypCd>05</rmaterialTypCd>"							'#����� ���� �ڵ� | 01 : ��깰, 02 : ���깰, 03 : ����ǰ, 04 : ������ �ǹ� ǥ�ô�� �ƴ�, 05 : ��ǰ�� �������� �󼼼��� ����
		strRst = strRst & "	<orgnTypCd>03</orgnTypCd>"										'������ �ڵ� | 01 : ����. ���������� �ڵ带 ���� �Է��ؾ� �մϴ�, 02 : �ؿ�. �ؿܿ����� �ڵ带 ���� �Է��ؾ� �մϴ�, 03 : ��Ÿ. ���������� �Է��ؾ��մϴ�
'		strRst = strRst & "	<orgnTypDtlsCd/>"												'������ ���� �ڵ� | ������ �ڵ尡 "����", "�ؿ�"�� ��� ������ ���� �ڵ� ���� �Է��ϼž� �մϴ�.
		strRst = strRst & "	<orgnNmVal><![CDATA["&GetSourcearea&"]]></orgnNmVal>"			'�������� | ������ �ڵ尡 "��Ÿ"�� ��� ���������� �Է��ϼž� �մϴ�.
'		strRst = strRst & "	<ProductRmaterial>"												'����� ���� | ����� ������ ����ǰ(03)�� ��� ����� ������ �Է��ϼž� �մϴ�. ��ǰ�� �ִ� 10��, ��ǰ�� ����� ���� ������ �ִ� 5������ ��� �����մϴ�.
'		strRst = strRst & "		<rmaterialNm/>"												'����� ��ǰ��
'		strRst = strRst & "		<ingredNm/>"												'�����
'		strRst = strRst & "		<orgnCountry/>"												'������
'		strRst = strRst & "		<content/>"													'�Է�
'		strRst = strRst & "	</ProductRmaterial>"
		strRst = strRst & "	<beefTraceStat>03</beefTraceStat>"								'#��깰 �̷¹�ȣ | 01 : �̷¹�ȣ ǥ�ô�� ��ǰ, 02 : �̷¹�ȣ ǥ�ô�� �ƴ�, 03 : �󼼼��� ����
'		strRst = strRst & "	<beefTraceNo/>"													'�̷¹�ȣ ǥ�ô�� ��ǰ | 01 : �̷¹�ȣ ǥ�ô����ǰ ���ý� �̷¹�ȣ ǥ�ô�� ��ǰ(xxxx)�� �� ������ �Է��մϴ�. Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���. ���ڼ��� 20Byte �� ���ѵ˴ϴ�. �ѱ� 10��, ����/���� 20�� �̳��� �Է��� �ֽʽÿ�.
		strRst = strRst & "	<sellerPrdCd>"&FItemid&"</sellerPrdCd>"							'�Ǹ��� ��ǰ�ڵ� | �ߺ��� �����ϸ� �� �ڵ尪���� 11���� ��ǰ ��ȸ ���� �����մϴ�. �ʼ����� �ƴϸ� ���� �����մϴ�.
		strRst = strRst & "	<suplDtyfrPrdClfCd>"&CHKIIF(FVatInclude="N","02","01")&"</suplDtyfrPrdClfCd>"	'#�ΰ���/�鼼��ǰ�ڵ� | �鼼��ǰ ���ý�, ����/������ å���� �Ǹ��ڴԲ� �ֽ��ϴ�. 01 : ������ǰ, 02 : �鼼��ǰ, 03 : ������ǰ
'		strRst = strRst & "	<forAbrdBuyClf></forAbrdBuyClf>"								'#�ؿܱ��Ŵ����ǰ ���� | SellerOffice ���� �ÿ� �۷ι������� ������ ��쿡�� ����� �� �ֽ��ϴ�. �Ϲ� ������ ��� �������ּ���. �ؿܰ��� �۷ι������� �Ϲ��ǸŻ�ǰ(01)�θ� ��ϰ����մϴ�. ���ǻ����� �����ø� 11���� ���MD�� ����� �ּ���.01 : �Ϲ��ǸŻ�ǰ, 02 : �ؿ��ǸŴ����ǰ
	If FItemdiv = "06" OR FItemdiv = "16" Then
		strRst = strRst & "	<prdStatCd>10</prdStatCd>"										'#��ǰ���� | �ֹ����ۻ�ǰ���� ����Ͻø� �������� ���/��ǰ/��ȯ�� �Ұ����Ͽ� Ŭ������ �߻��� �� ������ �����ϰ� ������ֽñ� �ٶ��ϴ�. Open API�� �ֹ����ۻ�ǰ ��� ��, �Ǹ��ڰ� �� ���뿡 ���� ������ �� ������ ������ ���ֵ˴ϴ�. 01 : ����ǰ, 02 : �߰��ǰ (�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 03 : ����ǰ 04 : ���ۻ�ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 05 : ����(����)��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 07 : ��ͼ���ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 08 : ��ǰ��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 09 : ��ũ��ġ��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 10 : �ֹ����ۻ�ǰ
	Else
		strRst = strRst & "	<prdStatCd>01</prdStatCd>"										'#��ǰ���� | �ֹ����ۻ�ǰ���� ����Ͻø� �������� ���/��ǰ/��ȯ�� �Ұ����Ͽ� Ŭ������ �߻��� �� ������ �����ϰ� ������ֽñ� �ٶ��ϴ�. Open API�� �ֹ����ۻ�ǰ ��� ��, �Ǹ��ڰ� �� ���뿡 ���� ������ �� ������ ������ ���ֵ˴ϴ�. 01 : ����ǰ, 02 : �߰��ǰ (�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 03 : ����ǰ 04 : ���ۻ�ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 05 : ����(����)��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 07 : ��ͼ���ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 08 : ��ǰ��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 09 : ��ũ��ġ��ǰ(�ǸŹ���� "�߰��Ǹ�"�� ��츸 ���ð����մϴ�.) 10 : �ֹ����ۻ�ǰ
	End If
'		strRst = strRst & "	<useMon/>"														'��밳���� | �ǸŹ���� �߰��Ǹ��� ��� �ݵ�� �Է��� �ּž� �մϴ�.
'		strRst = strRst & "	<paidSelPrc/>"													'���Դ�� �ǸŰ� | �ǸŹ���� �߰��Ǹ��� ��� �ݵ�� �Է��� �ּž� �մϴ�.
'		strRst = strRst & "	<exteriorSpecialNote/>"											'�ܰ�/��ɻ� Ư�̻��� | �ǸŹ���� �߰��Ǹ��� ��� �ݵ�� �Է��� �ּž� �մϴ�.
		strRst = strRst & "	<minorSelCnYn>"&Chkiif(IsAdultItem() = "Y", "N", "Y")&"</minorSelCnYn>"		'#�̼����� ���Ű��� | �̼����� ���źҰ��� �����Ͻø�, �̼����� ȸ������ ��ǰ�̹����� ������� ������ '19��'���� ǥ�õ˴ϴ�. ���źҰ� ��ǰ�� ���Ű������� ǥ���� ���, �Ǹű��� ó�� �� �� �ֽ��ϴ�. Y : ����, N : �Ұ���
		strRst = strRst & get11stAddImageParam()
		strRst = strRst & "	<htmlDetail><![CDATA["&get11stContParamToReg&"]]></htmlDetail>"	'�󼼼��� | iframe ����� ���������� �������� �ʽ��ϴ�. html �� �Է��Ͻ� ��� <![CDATA[ ]]> �� ���� �ּ���. �ܺη��� ��ũ�� ���ѵǸ� �ڼ��� ������ �󼼼��� ����ǥ�� ������ �ּ���. html�� �Է��ϴ� ��� �Ϻ� ��ũ��Ʈ �� ��Ÿ�� �±״� ���ѵǴ� SellerOffice ��ǰ��Ͽ��� �󼼼��� html �̸����⸦ �ݵ�� �׽�Ʈ�� �ּ���. html guide�� �ؼ��Ͽ� �Է��ϸ�, ���� ���� �ɼ� ã�Ⱑ �������ϴ�. Guide�� �̿��Ͽ� �󼼼����� ����� ������.
'		strRst = strRst & get11stSafeParam()		''2022-07-04 ���� ��
		strRst = strRst & get11stSafeNewParam()		''2022-07-04 ���� ��
'		strRst = strRst & "	<ProductMedical>"												'�Ƿ��� ǰ���㰡
'		strRst = strRst & "		<MedicalKey/>"												'�Ƿ��� ǰ���㰡��ȣ
'		strRst = strRst & "		<MedicalRetail/>"											'�Ƿ��� �Ǹž��Ű� ��� �� ��ȣ
'		strRst = strRst & "		<MedicalAd/>"												'�Ƿ������������ǹ�ȣ
'		strRst = strRst & "	</ProductMedical>"
		strRst = strRst & "	<reviewDispYn>Y</reviewDispYn>"									'��ǰ����/�ı� ���ÿ���
		strRst = strRst & "	<reviewOptDispYn>Y</reviewOptDispYn>"							'��ǰ����/�ı� �ɼ� ���⿩��
		strRst = strRst & "	<selTermUseYn>N</selTermUseYn>"									'#�ǸűⰣ | Y : ������. �ǸűⰣ ������ �����մϴ�., N : ��������(�������Ǹ��� ��츸). ��� �����ǸŰ� �̷�� ���ϴ�.
'		strRst = strRst & "	<selPrdClfCd/>"													'�ǸűⰣ�ڵ�/����Ⱓ�ڵ� | 0:100 : �ǸűⰣ �����Է�. �ǸŹ�� - "�������Ǹ�" �� ��츸 ��밡��, 3:101 : 3��, 5:102 : 5��, 7:103 : 7��, 15:104 : 15��, 30:105 : 30��(1����), 60:106 : 60��(2����), 90:107 : 90��(3����), 120:108 : 120��(4����), 3:401 : 3��, 5:402 : 5��, 7:403 : 7��, 15:404 : 15��, 30:405 : 30��(1����), 60:406 : 60��(2����), 90:407 : 90��(3����), 0:400 : ����Ⱓ �����Է�
'		strRst = strRst & "	<aplBgnDy/>"													'�Ǹ� ������/���� ������
'		strRst = strRst & "	<aplEndDy/>"													'�Ǹ� ������/���� ������ | �ǸűⰣ/����Ⱓ ���� �Է��� ��츸 �Է�. �������� �ڵ� ���. �������ּ���
		strRst = strRst & "	<setFpSelTermYn>N</setFpSelTermYn>"								'������ �ǸűⰣ ���� | Y : ������, N : ��������..�����Ǹ��� ���� ���� ����
'		strRst = strRst & "	<selPrdClfFpCd/>"												'�ǸűⰣ�ڵ� | ������ �ǸűⰣ ���� - Y�� ��츸 ��밡��. 0:100 : ��������, 3:101 : 3��, 5:102 : 5��, 7:103 : 7��, 15:104 : 15��, 30:105 : 30��(1����), 60:106 : 60��(2����), 90:107 : 90��(3����), 120:108 : 120��(4����)
'		strRst = strRst & "	<wrhsPlnDy/>"													'�԰����� | �԰������� �Ǹ������ϰ� ���� ��, Ȥ�� �� ���ķ� ������ �ּž� �ϸ�, �ֹ�ó�� ��, �ִ� 15�Ͽ� ���ؼ� 1ȸ ������ �� �ֽ��ϴ�. �Է��Ͻ� �԰������� ��ǰ�� �������� �ȳ��Ǹ�, �԰����� ���� ��, �ſ������� �����ǿ���, ������ �ֽʽÿ�.
'		strRst = strRst & "	<contractCd/>"													'�����ڵ� | �޴��� ī�װ��� ��ǰ ��Ͻ� �ʼ��� �����ϼž� �մϴ�. 01 : �Ϲ� ���� �ܸ���, 02 : ����� ���� �ܸ���
'		strRst = strRst & "	<chargeCd/>"													'����� �ڵ� | ����� ���� �ܸ����� ��� �ݵ�� �Է��ϼž� �մϴ�.
'		strRst = strRst & "	<periodCd/>"													'�����Ⱓ �ڵ� | 01 : ������, 02 : 12����, 03 : 18����, 04 : 24����, 05 : 30����, 06 : 36����
'		strRst = strRst & "	<phonePrc/>"													'�ܸ��� ��� ���� | ,���� ���ڸ� �Է��ϼ���. 60,000(X) 60000(O)
		If FDepth1Code = "2967" Then		'����ī�װ�
		strRst = strRst & "	<maktPrc>"&ForgPrice&"</maktPrc>"								'���� | ī�װ��� ������ ��� �ݵ�� �Է��� �ּž� �մϴ�.(����, DVD/��緹������) ���� ������ ���� ������ �ؼ��Ͽ� ����ϼž� �մϴ�. ���������� ������ �����ϰų� ������ ������ ����� ��� ���� å���� �Ǹ��ڿ��� ������, �Ǹ����� ���� �������� ���� �� �ֽ��ϴ�. ���� ������: 18���� �̸��� �Ű������� ��� ������ 10% �̳� ����, �������� �ǸŰ��� �ִ� 10% �̳� ���� �����մϴ�.(���ϸ���, OKĳ���� �� �ջ� ����)
		End If
		strRst = strRst & "	<selPrc>"&MustPrice&"</selPrc>"									'#�ǸŰ� | �ǸŰ��� 10�� ������, �ִ� 10�� �� �̸����� �Է� �����մϴ�. �ǸŰ� ���� ���� ��, �ִ� 50% �λ�/80% ���ϱ��� �����Ͻ� �� �ֽ��ϴ�. �����̿��� ī�װ�/�ǸŰ��� ���� �ٸ��� ����� �� �ֽ��ϴ�.
'		strRst = strRst & "	<cuponcheck/>"													'�⺻������� �������� | "�⺻�������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. S : ������ ����(��ǰ����) �� ��ǰ�����ÿ��� �Է°����մϴ�. ������ ���� ������ �Ͼ�� �ʽ��ϴ�. Y : ������, N : ��������, S : ������ ����(��ǰ����)
'		strRst = strRst & "	<dscAmtPercnt/>"												'���μ�ġ | "�⺻�������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ǸŰ�����(xxxx)�� �� ��ġ�� �Է��մϴ�.
'		strRst = strRst & "	<cupnDscMthdCd/>"												'���δ��� �ڵ� | "�⺻�������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : ��, 02 : %
'		strRst = strRst & "	<cupnUseLmtDyYn/>"												'���� ����Ⱓ �������� | "�⺻�������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "���� ����Ⱓ ����"�� �Ͻ� ��쿡�� Element�� �Է��� �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<cupnIssEndDy/>"												'��������Ⱓ ������ | "�⺻�������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "���� ����Ⱓ ����"�� �Ͻ� ��쿡�� Element�� �Է��� �ּ���.
'		strRst = strRst & "	<ocbYN/>"														'OKĳ���� ���� �������� | "OKĳ���� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<ocbValue/>"													'������ġ | "OKĳ���� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ǸŰ�����(xxxx) �� �� ��ġ�� �Է��մϴ�.
'		strRst = strRst & "	<ocbWyCd/>"														'�������� �ڵ� | "OKĳ���� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : %, 02 : ��
'		strRst = strRst & "	<mileageYN/>"													'���ϸ��� ���� �������� | "���ϸ��� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<mileageValue/>"												'������ġ | "���ϸ��� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ǸŰ�����(xxxx) �� �� ��ġ�� �Է��մϴ�.
'		strRst = strRst & "	<mileageWyCd/>"													'�������� �ڵ� | "���ϸ��� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : %, 02 : ��
'		strRst = strRst & "	<intFreeYN/>"													'������ �Һ� ���� �������� | "������ �Һ� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<intfreeMonClfCd/>"												'������| "������ �Һ� ����"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 05 : 2����, 01 : 3����, 06 : 4����, 07 : 5����, 02 : 6����, 08 : 7����, 09 : 8����, 03 : 9����, 10 : 10����, 11 : 11����, 04 : 12����
'		strRst = strRst & "	<pluYN/>"														'������������ ���� ���� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<pluDscCd/>"													'������������ ���� ���� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : ��������, 02 : �ݾױ���
'		strRst = strRst & "	<pluDscBasis/>"													'������������ ���� �ݾ� �� ���� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
'		strRst = strRst & "	<pluDscAmtPercnt/>"												'������������ �ݾ�/�� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
'		strRst = strRst & "	<pluDscMthdCd/>"												'������������ �����ڵ� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : %, 02 : ��
'		strRst = strRst & "	<pluUseLmtDyYn/>"												'������������ ����Ⱓ ���� | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<pluIssStartDy/>"												'������������ ����Ⱓ ������ | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "���� ����Ⱓ ����"�� �Ͻ� ��쿡�� Element�� �Է��� �ּ���.
'		strRst = strRst & "	<pluIssEndDy/>"													'������������ ����Ⱓ ������ | "���� ��������"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. "���� ����Ⱓ ����"�� �Ͻ� ��쿡�� Element�� �Է��� �ּ���.
'		strRst = strRst & "	<hopeShpYn/>"													'����Ŀ� ���� ���� ���� | "����Ŀ�"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. Y : ������, N : ��������
'		strRst = strRst & "	<hopeShpPnt/>"													'������ġ | "����Ŀ�"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
'		strRst = strRst & "	<hopeShpWyCd/>"													'�������� �ڵ� | "����Ŀ�"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : %, 02 : ��
	If (FOptionCnt > 0) OR (FItemdiv = "06") Then
		strRst = strRst & get11stOptionParam()
	End If
'		strRst = strRst & "	<useOptCalc/>"													'������ɼ� �������� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. ��� : ����� �ɼ� ���� �Է� �� ���, ���Է� ������ ���� : Y : ���, N : ���� ����� �ɼ��� ����/��������/�л�����, ħ��/Ŀư/ī��Ʈ, Ȩ/���׸���/DIY ī�װ������� ��� �����մϴ�. ����� �ɼ��� ������ �ɼ��� �ּ� 1�� �̻� �Բ� ����ؾ� ��� �����մϴ�. ����� �ɼ��� �ۼ��� �ɼǰ� ���ÿ� ����� �� �����ϴ�. ������ �ɼ��� ����ϸ� ����� �ɼ��� ����� �� �����ϴ�. �Ǹ��ּҰ�, �Ǹ��ִ밪, �ܰ����ذ�, �ǸŴ���-���ڴ� ���ڷ� �Է��ϼ���. �ʱ� ���߽ÿ��� �ɼǵ�ϰ� SellerOffice ��ǰ��ϰ� �ݵ�� �����Ͻø鼭 �������ּž� �մϴ�.
'		strRst = strRst & "	<optCalcTranType/>"												'����� �ɼ� Ÿ�Լ��� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. reg : ���, upd : ����
'		strRst = strRst & "	<optTypCd/>"													'������ɼǱ��а� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
'		strRst = strRst & "	<optItem1Nm/>"													'ù��° ����� �ɼǸ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ִ� 20byte, �ʰ� ������ ����
'		strRst = strRst & "	<optItem1MinValue/>"											'ù��° ����� �ɼ� �Ǹ��ּҰ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 1~1,000,000
'		strRst = strRst & "	<optItem1MaxValue/>"											'ù��° ����� �ɼ� �Ǹ��ִ밪 | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 1~1,000,000
'		strRst = strRst & "	<optItem2Nm/>"													'�ι�° ����� �ɼǸ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �ִ� 20byte, �ʰ� ������ ����
'		strRst = strRst & "	<optItem2MinValue/>"											'�ι�° ����� �ɼ� �Ǹ��ּҰ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 1~1,000,000
'		strRst = strRst & "	<optItem2MaxValue/>"											'�ι�° ����� �ɼ� �Ǹ��ִ밪 | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 1~1,000,000
'		strRst = strRst & "	<optUnitPrc/>"													'?�ܰ����ذ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 0.001~1,000,000
'		strRst = strRst & "	<optUnitCd/>"													'?���� �����ڵ� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. 01 : mm, 02 : cm, 03 : m
'		strRst = strRst & "	<optSelUnit/>"													'?�ǸŴ���-���� | "�ɼǵ��"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���. �Է� ���� ���� 1~1,000,000
'		strRst = strRst & "	<ProductComponent>"												'?�߰�������ǰ | "�߰�������ǰ���"�� ���� ���� �����ÿ��� Element�� ��� ������ �ּ���.
'		strRst = strRst & "		<addPrdGrpNm/>"
'		strRst = strRst & "		<compPrdNm/>"
'		strRst = strRst & "		<sellerAddPrdCd/>"
'		strRst = strRst & "		<addCompPrc/>"
'		strRst = strRst & "		<compPrdQty/>"
'		strRst = strRst & "		<compPrdVatCd/>"
'		strRst = strRst & "		<addUseYn/>"
'		strRst = strRst & "		<addPrdWght/>"
'		strRst = strRst & "	</ProductComponent>"
		strRst = strRst & "	<prdSelQty>"&getLimit11stEa&"</prdSelQty>"						'������ | ��� ������ �ݵ�� �Է��ϼž� �ϸ� �ɼ��� ���� ��� �Է°��� ������� �ɼǼ����� �������� �ڵ���� �Ǿ� �ݿ��˴ϴ�. ���� 0���� �Է��� �� �����ϴ�. ��ǰ �Ǹ� �ߴ��� ���Ͻø� �Ǹ����� ó���Ͻñ� �ٶ��ϴ�.
'		strRst = strRst & "	<selMinLimitTypCd/>"											'�ּұ��ż��� �����ڵ� | "�ּұ��ż���" ���񽺸� �̿����� �����Ŵٸ� Element�� ������ �ּ���. �ڵ� "���� ����(00)"���� �����˴ϴ�. ���� ���� �ش� ���� �Ⱓ�� �Ѵ�(30��)�Դϴ�. | 00 : ���� ����, 01 : 1ȸ ����
'		strRst = strRst & "	<selMinLimitQty/>"												'�ּұ��ż��� ���� | "�ּұ��ż���" ���񽺸� �̿����� �����Ŵٸ� Element�� ������ �ּ���. �ڵ� "���� ����(00)"���� �����˴ϴ�.
		strRst = strRst & "	<selLimitTypCd>01</selLimitTypCd>"								'�ִ뱸�ż��� �����ڵ� | "�ִ뱸�ż���" ���񽺸� �̿����� �����Ŵٸ� "�����ڵ�"�� ������� Element�� ������ �ּ���. �ڵ� "���� ����(00)"���� �����˴ϴ�. 00 : ���� ����, 01 : 1ȸ ����, 02 : �Ⱓ ����
		strRst = strRst & "	<selLimitQty>"& FOrderMaxNum &"</selLimitQty>"					'�ִ뱸�ż��� ���� | "�ִ뱸�ż���" ���񽺸� �̿����� �����Ŵٸ� "�����ڵ�"�� ������� Element�� ������ �ּ���. �ڵ� "���� ����(00)"���� �����˴ϴ�.
'		strRst = strRst & "	<townSelLmtDy/>"												'�ִ뱸�ż��� �籸�űⰣ | "�ִ뱸�ż���" ���񽺸� �̿����� �����Ŵٸ� "�����ڵ�"�� ������� Element�� ������ �ּ���. �ڵ� "���� ����(00)"���� �����˴ϴ�.
		strRst = strRst & "	<useGiftYn>N</useGiftYn>"										'����ǰ ���� ��뿩�� | Y : �����, N : ������
'		strRst = strRst & "	<ProductGift>"
'		strRst = strRst & "		<giftInfo/>"												'����ǰ ����
'		strRst = strRst & "		<giftNm/>"
'		strRst = strRst & "		<aplBgnDt/>"
'		strRst = strRst & "		<aplEndDt/>"
'		strRst = strRst & "	</ProductGift>"
		strRst = strRst & "	<gblDlvYn>N</gblDlvYn>"											'#�������� �̿뿩�� | �������� ���� ���, �⺻ '�̿����(N)���� ���õǸ�, �Ʒ��� ���� ������ ��� �����Ǿ�� �����մϴ�. 1. ����ȸ�������� �������� �̿뿩�ΰ� "���� �Ǵ� �̿�"���� �Ǿ��ְ� 2. ����Ϸ��� ��ǰī�װ��� �����Թ�� �̿뿩�ΰ� "�̿�(Y)"���� �Ǿ��ְ� ī�װ��� �������� ���ɿ���Ȯ�� 3. ��ǰ�ɼ��� "������"�� �ƴϾ�� �ϰ� 4. ��ǰ�� ��۹���� '�ù�' �Ǵ� '����(����/���)'�� �Ǿ��ְ� 5. ��ǰ�� ��ۺ����� '����' �Ǵ� ��������� '����������' Ȥ�� '������ �ʼ�'�̾�� �ϰ� 6. ��������� ���� ������ ������ �ݵ�� �����ؾ� �ϰ� 7. ��ǰ���Դ� �ݵ�� �Է��ؾ� �ϰ� 8. ��ǰ�� ������� "�����ּ�" �ΰ�츸 '��������'�� ����. Y : �̿�, N : �̿����
'		strRst = strRst & "	<gblHsCode/>"													'�������� HSCode | �������� "�̿�" ��ǰ�ΰ�� �ʼ��� �Է��ϼž� �մϴ�.
		strRst = strRst & "	<dlvCnAreaCd>02</dlvCnAreaCd>"									'#��۰������� �ڵ� | 01 : ����, 02 : *����(���� �����갣���� ����), 03 : ����, 04 : ��õ, 05 : ����, 06 : �뱸, 07 : ����, 08 : �λ�, 09 : ���, 10 : ���, 11 : ����, 12 : �泲, 13 : ���, 14 : �泲, 15 : ���, 16 : ����, 17 : ����, 18 : ����, 19 : ����/���, 20 : ����/���/����, 21 : ���/�泲, 22 : ���/�泲, 23 : ����/����, 24 : �λ�/���, 25 : ����/���/���ֵ����갣 ��������, 26 : �Ϻ������Ұ�
		strRst = strRst & "	<dlvWyCd>01</dlvWyCd>"											'#��۹�� | "�������� ��ǰ" �ΰ�� '�ù� �Ǵ� ����(����/���)'�� �Է°����մϴ�. 01 : �ù�, 02 : ����(����/���), 03 : ��������(ȭ�����), 04 : ������, 05 : ����ʿ����
		'2019-05-23 10:465 dlvSendCloseTmpltNo �߰�
		strRst = strRst & "	<dlvSendCloseTmpltNo>570949</dlvSendCloseTmpltNo>"				'#�߼۸��� ���ø���ȣ | �߼۸��� ���ø���ȣ (���ù߼�, �Ϲݹ߼�) 1�� ��� �����ϸ� �����Է� �����Դϴ�. �⺻������ ��۹���� �ù��� ��ǰ�� ���Ͽ� ��ȿ�ϸ� �ؿ����� ��ǰ, �����ǸŻ�ǰ, �ֹ����ۻ�ǰ, ������Ź��� ��ǰ�� �ݿ� ��󿡼� ���ܵ˴ϴ�.
		strRst = strRst & "	<dlvCstInstBasiCd>07</dlvCstInstBasiCd>"						'#��ۺ� ���� | 01 : ����, 02 : ���� ��ۺ�, 03 : ��ǰ ���Ǻ� ����, 04 : ������ ����, 05 : 1���� ��ۺ�, 07 : �Ǹ��� ���Ǻ� ��ۺ� 2010.08.20 06->07 �� ����, 08 : ����� ���Ǻ� ��ۺ� 2010.10.08 �߰�, 09 : 11���� ���� ����� ��ۺ�, 10 : 11�����ؿܹ�����Ǻι�ۺ� (11���� �ؿ� ����� ����ϴ� ���)
'		strRst = strRst & "	<dlvCst1/>"														'��ۺ� | ��ǰ ���Ǻ� ����(03), ���� ��ۺ�(02)
'		strRst = strRst & "	<dlvCst4/>"														'��ۺ� | 1���� ��ۺ�(05)
'		strRst = strRst & "	<dlvCst3/>"														'��ۺ� | ������ ����(04) "������ ����"�� �����߰��� ���� ��ۺ� �ִ� 10������ ���� ����
'		strRst = strRst & "	<dlvCstInfoCd/>"												'��ۺ� | ���� ��ۺ�(02) ��ۺ� �߰� �ȳ� �������� : ��������(N), �������Ұ�(02) 01 : (��ǰ������), 02 : (��ǰ�� ���� ����), 03 : (������ ���� ����), 04 : (��ǰ/������ ����), 06 : (����/��� ����, �̿� �߰����)
'		strRst = strRst & "	<PrdFrDlvBasiAmt/>"												'��ǰ���Ǻ� ���� ��ǰ���رݾ� | ��ǰ���Ǻ� ����(03)
'		strRst = strRst & "	<dlvCnt1/>"														'������ ���� ���� ~�̻� ���� | ������ ����(04) "������ ����"�� �����߰��� ���� ���� ������ �ִ� 10������ ���� ����
'		strRst = strRst & "	<dlvCnt2/>"														'������ ���� ���� ~���� ���� | ������ ����(04) "������ ����"�� �����߰��� ���� ���� ������ �ִ� 9������ ���� ����
		strRst = strRst & "	<bndlDlvCnYn>Y</bndlDlvCnYn>"									'#������� ���� | Y : ����, N : �Ұ�
		strRst = strRst & "	<dlvCstPayTypCd>03</dlvCstPayTypCd>"							'#������� | 01 : ����������, 02 : �������Ұ�, 03 : �������ʼ�
		strRst = strRst & "	<jejuDlvCst>3000</jejuDlvCst>"									'#���� | ���� �߰� ��ۺ�
		strRst = strRst & "	<islandDlvCst>3000</islandDlvCst>"								'#�����갣 | �����갣 �߰� ��ۺ�
		strRst = strRst & "	<addrSeqOut>2</addrSeqOut>"										'#����� �ּ� �ڵ� | �켱 SellerOffice ��ǰ��Ͽ��� ����� �ּҰ� ����� �Ǿ��־�� �մϴ�. ��ϵ� ����� �ּҸ� Api ��ȸ ���񽺸� ���� �ּ� �������ڵ带 ��ȸ�մϴ�. ����� �ּ���ȸ���� ��ȸ�� �������ڵ带 �Է��Ͻø� �˴ϴ�. ���� ����� �ּҸ� �����Ͻ� ��� �⺻�ּҷ� �ڵ� ������ �˴ϴ�. �Ͽ� ��ǰ������ �Ͻǰ�� ��������� �⺻�ּҷ� �缳���� �˴ϴ�. ��ǰ������ �⺻�ּ� �������� ���� �̽������� ���̱� ���� ����� �ڵ� �Է��� �����մϴ�. ���������� �Ƿ��� ������� �������߸� �Ѵ�.
'		strRst = strRst & "	<outsideYnOut/>"												'����� �ּ� �ؿ� ���� | "����� �ּ� �ؿ� ����"�� ����� �ּҰ� �ؿ��� ��쿡�� �Է��Ͻð� ������ ���� ������ �ּ���. Y : ����� �ؿ�, N : ����� ����
'		strRst = strRst & "	<addrSeqOutMemNo/>"												'���� ID ȸ�� ��ȣ | ������� "���� ID ȸ�� ��ȣ (�������)"�� ���� ����� ����ϴ� ��쿡�� �Է��� �ּ���.
		strRst = strRst & "	<addrSeqIn>3</addrSeqIn>"										'#��ǰ/��ȯ�� �ּ� �ڵ� | �켱 SellerOffice ��ǰ��Ͽ��� ��ǰ/��ȯ�� �ּҰ� ����� �Ǿ��־�� �մϴ�. ��ϵ� ��ǰ/��ȯ�� �ּҸ� Api ��ȸ ���񽺸� ���� �ּ� �������ڵ带 ��ȸ�մϴ�. ��ǰ/��ȯ�� �ּ���ȸ���� ��ȸ�� �������ڵ带 �Է��Ͻø� �˴ϴ�. ���� ��ǰ/��ȯ�� �ּҸ� �����Ͻ� ��� �⺻�ּҷ� �ڵ� ������ �˴ϴ�. �Ͽ� ��ǰ������ �Ͻ� ��� ��������� �⺻�ּҷ� �缳���� �˴ϴ�. ��ǰ������ �⺻�ּ� �������� ���� �̽������� ���̱� ���� ����� �ڵ� �Է��� �����մϴ�.
'		strRst = strRst & "	<outsideYnIn/>"													'��ǰ/��ȯ�� �ּ� �ؿ� ���� | "��ǰ/��ȯ�� �ּ� �ؿ� ����"�� ��ǰ/��ȯ�� �ּҰ� �ؿ��� ��쿡�� �Է��Ͻð� ������ ���� ������ �ּ���. Y : ��ǰ/��ȯ�� �ؿ�, N : ��ǰ/��ȯ�� ����
'		strRst = strRst & "	<addrSeqInMemNo/>"												'���� ID ȸ�� ��ȣ | ��ǰ���� "���� ID ȸ�� ��ȣ (��ǰ����)"�� ���� ��ǰ�� ����ϴ� ��쿡�� �Է��� �ּ���.
'		strRst = strRst & "	<abrdCnDlvCst/>"												'�ؿ���� ��ۺ� | ��ۺ�� 10�������� �Է��ϼž� �մϴ�. 3000(O), 2999(X), 2,900(X) ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�.

''''''''��ȯ/��ǰ�� �ݾ� ����..2020-01-20 ������
'		strRst = strRst & "	<rtngdDlvCst>2500</rtngdDlvCst>"								'#��ǰ ��ۺ� | ��ۺ�� 10�������� �Է��ϼž� �մϴ�. 3000(O), 2999(X), 2,900(X)
'		strRst = strRst & "	<exchDlvCst>5000</exchDlvCst>"									'#��ȯ ��ۺ�(�պ�) | ��ۺ�� 10�������� �Է��ϼž� �մϴ�. 3000(O), 2999(X), 2,900(X)
		strRst = strRst & "	<rtngdDlvCst>3000</rtngdDlvCst>"								'#��ǰ ��ۺ� | ��ۺ�� 10�������� �Է��ϼž� �մϴ�. 3000(O), 2999(X), 2,900(X)
		strRst = strRst & "	<exchDlvCst>6000</exchDlvCst>"									'#��ȯ ��ۺ�(�պ�) | ��ۺ�� 10�������� �Է��ϼž� �մϴ�. 3000(O), 2999(X), 2,900(X)
''''''''��ȯ/��ǰ�� �ݾ� ����..2020-01-20 ������ ��

		strRst = strRst & "	<rtngdDlvCd>01</rtngdDlvCd>"									'�ʱ��ۺ� ����� �ΰ���� | �ʱ��ۺ� ���� �� �ΰ���� �����ڵ带 �Է����� ���� ��� ����ǰ��ۺ� ��ȯ��ۺ񺸴� ũ�ų� ���� ��쿡�� ��02�� ��, ����ǰ��ۺ� ��ȯ��ۺ񺸴� ���� ��쿡�� ��01���պ� �ڵ尡 �ڵ� ��ϵ˴ϴ�. 01 : �պ�(��x2), 02 : ��
		strRst = strRst & "	<asDetail><![CDATA[�ٹ����� ���ູ���� 1644-6035]]></asDetail>"	'#A/S �ȳ� | �ݵ�� �Է��ϼž� �ϸ� �Է��� ������ �����ø� . �̶� �Է����ּž� �մϴ�. ������ �ȵ˴ϴ�. Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���.
		strRst = strRst & "	<rtngExchDetail><![CDATA[�ٹ����� ���ູ���� 1644-6035]]></rtngExchDetail>"	'#��ǰ/��ȯ �ȳ� | ��ǰ�� �������� �ȳ��Ǵ� ��������, ��ǰ/��ȯ ���Ǹ� ���̽� �� �ֽ��ϴ�. �ݵ�� �Է��ϼž� �ϸ� �Է��� ������ �����ø� . �̶� �Է����ּž� �մϴ�. ������ �ȵ˴ϴ�. Ư������ ���� ���ԵǾ� ���� ��� <![CDATA[ ]]> �� ���� �ּ���.
		strRst = strRst & "	<dlvClf>02</dlvClf>"											'#��� ��ü | ������� ���� ������ �˴ϴ�. ��ǰ�� ������� �Ǹ��� ����� �� ���: ��ü ���, 11���� ���� ID�� ������� ���: 11���� ���, 11���� �ؿ� ���� ������� ���: 11���� �ؿܹ��, 01 : 11���� ��� (���� ID�� ������� ����ϴ� ���), 02 : ��ü��� (������ ����� ó���ϴ� ���), 03 : 11���� �ؿ� ��� (11���� �ؿ� ���� ������� ����ϴ� ���) �������� �ʴ� ��� default�� 02(��ü���)���� ó���˴ϴ�.
'		strRst = strRst & "	<abrdInCd/>"													'#11���� �ؿ� �԰� ���� | ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�. ��ǰ�� ������� 11���� ���� �Ⱦ� ���� ������ ���: 11���� ���� �Ⱦ�, �Ǹ��� ���� �߼��� ���: �Ǹ��ڹ߼�, ���� ������ ���: ���� ���� 01 : 11���� ���� �Ⱦ�, 02 : �Ǹ��ڹ߼�, 03 : ���� ����
'		strRst = strRst & "	<prdWght/>"														'#��ǰ ���� | g ������ �Է� ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �Ǵ� �������� ��ǰ" �ΰ�� �ʼ��Դϴ�.
'		strRst = strRst & "	<ntShortNm/>"													'#����������(�����) | ��ǥ��ǰ�� ������������ �����Ͻø� �˴ϴ�. ������ ��ۻ�ǰ�� ��� �ʼ��Դϴ�. �Ʒ� ÷�������� ���� ��Ź �帳�ϴ�.
'		strRst = strRst & "	<globalOutAddrSeq/>"											'#�Ǹ��� �ؿ� ����� �ּ� | ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�. �켱 SellerOffice ��ǰ��Ͽ��� ����� �ּ�(�ؿ�)�� ����� �Ǿ��־�� �մϴ�. ��ϵ� ����� �ּҸ� Api ��ȸ ���񽺸� ���� �ּ� �������ڵ带 ��ȸ�մϴ�. ����� �ּ���ȸ���� ��ȸ�� �������ڵ带 �Է��Ͻø� �˴ϴ�. ��ǰ������ �⺻�ּ� �������� ���� �̽������� ���̱� ���� ����� �ڵ� �Է��� �����մϴ�.
'		strRst = strRst & "	<mbAddrLocation05/>"											'#�Ǹ��� �ؿ� ����� ���� ���� | ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�. �ؿ� �ڵ�� �Է��Ͻñ� �ٶ��ϴ�. 01 : ����, 02 : �ؿ�
'		strRst = strRst & "	<globalInAddrSeq/>"												'#�Ǹ��� ��ǰ/��ȯ�� �ּ� | ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�. �켱 SellerOffice ��ǰ��Ͽ��� ��ǰ/��ȯ�� �ּ�(�ؿ�)�� ����� �Ǿ��־�� �մϴ�. ��ϵ� ��ǰ/��ȯ�� �ּҸ� Api ��ȸ ���񽺸� ���� �ּ� �������ڵ带 ��ȸ�մϴ�. ��ǰ/��ȯ�� �ּ���ȸ���� ��ȸ�� �������ڵ带 �Է��Ͻø� �˴ϴ�. ��ǰ������ �⺻�ּ� �������� ���� �̽������� ���̱� ���� ��ǰ/��ȯ�� �ڵ� �Է��� �����մϴ�.
'		strRst = strRst & "	<mbAddrLocation06>01</mbAddrLocation06>"						'#�Ǹ��� ��ǰ/��ȯ�� ���� ���� | ��� ��ü(dlvClf) �ڵ尡 03(11���� �ؿ� ���)�� ��� �ʼ��Դϴ�. 01 : ����, 02 : �ؿ�
'		strRst = strRst & "	<mnfcDy/>"														'��������
'		strRst = strRst & "	<eftvDy/>"														'��ȿ����
		strRst = strRst & get11stItemInfoCdParameter
		strRst = strRst & "	<company><![CDATA["&CStr(FMakerName)&"]]></company>"			'������ | ������� �ؽ�Ʈ ���·θ� �Է��ϸ� �����簡 ���� �� "����"���� �Է��մϴ�.
'		strRst = strRst & "	<modelNm/>"														'�𵨸� | �𵨸��� �ؽ�Ʈ ���·θ� �Է��ϸ� �𵨸��� ���� �� "����"���� �Է��մϴ�. (���� �𵨸� : �⺻ ���� ���� SQBAB9401)
'		strRst = strRst & "	<modelCd/>"														'���ڵ� | ����Ͻ� ��ǰ ���� ������ �ĺ�������, ����+����, ����, ���� ������ ���յ� �𵨹�ȣ�� �Է����ֽʽÿ�. (���� ���ڵ� : SQBAB9401)
'		strRst = strRst & "	<mnfcDy/>"														'����/������� | ī�װ��� ����/����/DVD�� ��� �Է��� �ּž� �մϴ�. �������ڸ� ������ �����Ұ��, ������ȭ������� ���� ó�� ���� �� �ֽ��ϴ�.
'		strRst = strRst & "	<mainTitle/>"													'���� | ī�װ��� ������ ��� �����Ͽ� �Է��Ͻ� �� �ֽ��ϴ�. �������� �ѱ�50��, ����/���� 100�� �̳��� �Է��� �ֽʽÿ�.
'		strRst = strRst & "	<artist/>"														'��Ƽ��Ʈ/����(���) | ī�װ��� ����/DVD�� ��� �ݵ�� �Է��� �ּž� �մϴ�.(����,����(TAPE),DVD/����) ��Ƽ��Ʈ/����(���)���� �ѱ� 50��,����/���� 100�� �̳��� �Է��� �ֽʽÿ�.
'		strRst = strRst & "	<mudvdLabel/>"													'���� �� | ī�װ��� ����/����/DVD>����, ����[TAPE]�� ��� �����Ͽ� �Է��Ͻ� �� �ֽ��ϴ�. ���̺��� �ѱ� 100��,����/���� 200�� �̳��� �Է��� �ֽʽÿ�.
'		strRst = strRst & "	<maker/>"														'������ | ī�װ��� ����/����/DVD>����, ����[TAPE],DVD �� ��� �ݵ�� �Է� �ϼž� �մϴ�.(����,����(TAPE),DVD/����) ���̺��� �ѱ� 100��,����/���� 200�� �̳��� �Է��� �ֽʽÿ�.
'		strRst = strRst & "	<albumNm/>"														'�ٹ��� | ī�װ��� ����/����/DVD>����, ����[TAPE] �� ��� �ݵ�� �Է� �ϼž� �մϴ�.(����,����(TAPE)) �ٹ����� �ѱ� 100��,����/���� 200�� �̳��� �Է��� �ֽʽÿ�
'		strRst = strRst & "	<dvdTitle/>"													'DVD Ÿ��Ʋ | ī�װ��� ����/����/DVD>DVD �� ��� �ݵ�� �Է� �ϼž� �մϴ�.(DVD/����) DVD Ÿ��Ʋ�� �ѱ� 100��,����/���� 200�� �̳��� �Է��� �ֽʽÿ�.
'		strRst = strRst & "	<bcktExYn/>"													'��ٱ��� ��� ���� | ��ٱ��� ��� ������ Y/N���� �Է� �˴ϴ�.
		strRst = strRst & "	<prcCmpExpYn>Y</prcCmpExpYn>"									'���ݺ� ����Ʈ ��� ���� | ���ݺ񱳻���Ʈ ����� ���û����̸� "�����"�� �����մϴ�. Y : �����, N : ��Ͼ���
		strRst = strRst & "</Product>"
		get11stItemRegParameter = strRst
'response.write strRst
'response.end
	End Function

	'��ǰ�����������
    Public Function get11stItemInfoCdParameter()
		Dim strSql, buf
		Dim mallinfoCd, infoContent, mallinfodiv, vType
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00001') THEN '������ ����ǥ��' "
		strSql = strSql & " 	 WHEN (M.infoCd='00002') THEN '�������� ����' "
		strSql = strSql & " 	 WHEN (M.infoCd='10000') THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' "
		strSql = strSql & " 	 WHEN (M.infoCd='21011') AND Len(isNull(F.infocontent, '')) < 2 THEN I.itemname "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035' "
		strSql = strSql & " 	WHEN LEN( isNull(F.infocontent, '')) < 2 THEN '��ǰ �� ����' " & vbcrlf
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = '11st' and IC.itemid='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = CInt(rsget("mallinfodiv"))
			vType = 891010 + mallinfodiv
			If mallinfodiv = "47" Then
				vType = "1149547"
			ElseIf mallinfodiv = "48" Then
				vType = "1149546"
			End If

			buf = buf & "	<ProductNotification>"												'��ǰ�����������
			buf = buf & "		<type>"&vType&"</type>"											'#�����ڵ�
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
				End If
				buf = buf & "			<item>"													'#�׸����� | ������ �ش��ϴ� �׸�����
				buf = buf & "				<code><![CDATA["&mallinfoCd&"]]></code>"			'�׸��ڵ�
				buf = buf & "				<name><![CDATA["&infoContent&"]]></name>"			'�׸� | ��¥�Է� ����� YYYY/MM/DD (��/��/��) �������� �Է��ؾ� �մϴ�.
				buf = buf & "			</item>"
				rsget.MoveNext
			Loop
			buf = buf & "	</ProductNotification>"
		End If
		rsget.Close
		get11stItemInfoCdParameter = buf
    End Function
End Class

Class C11st
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectItemID

	Private Sub Class_Initialize()
		redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

	'// �̵�� ��ǰ ���(��Ͽ�)
	Public Sub get11stNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "	SELECT itemid FROM ("
            addSql = addSql & "     SELECT itemid"
            addSql = addSql & " 	,count(*) as optCNT"
			addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	FROM db_item.dbo.tbl_item_option"
            addSql = addSql & " 	WHERE itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	GROUP BY itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, isNULL(R.st11StatCD,-9) as st11StatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, tm.safeDiv, tm.isNeed, tm.depth1Code, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_11st_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		ELSE
		    strSql = strSql & " and (i.deliveryType<>9)"
	    END IF
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'�ù�(�Ϲ�)
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
'		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and ( "
		strSql = strSql & "			convert(varchar(6), (i.cate_large + i.cate_mid)) not in ( "
		strSql = strSql & "				SELECT convert(varchar(6), cdl+cdm)  "
		strSql = strSql & "				FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category]  "	'2023-06-23 ������ / ������� ī�װ��� Ư���귣��� �Ǹŵǵ���
		strSql = strSql & "				WHERE mallgubun='"&CMALLNAME&"' "
		strSql = strSql & "		) or i.makerid in ( "
		strSql = strSql & "			'heidi2022', "
		strSql = strSql & "			'luna2022', "
		strSql = strSql & "			'uand2051', "
		strSql = strSql & "			'wpc001', "
		strSql = strSql & "			'lifeshop0510', "
		strSql = strSql & "			'bijou2023', "
		strSql = strSql & "			'blesscompany', "
		strSql = strSql & "			'JINNYSTAR01', "
		strSql = strSql & "			'sportsconnection', "
		strSql = strSql & "			'ithinkso', "
		strSql = strSql & "			'greenh03', "
		strSql = strSql & "			'kingkongoutlet', "
		strSql = strSql & "			'shoemiz', "
		strSql = strSql & "			'goldn', "
		strSql = strSql & "			'funnyfun', "
		strSql = strSql & "			'doran1020', "
		strSql = strSql & "			'gabangpop1010', "
		strSql = strSql & "			'osjarak' "
		strSql = strSql & "		) "
		strSql = strSql & "	)  "
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_11st_regItem WHERE st11StatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.	'11st��ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub get11stEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, m.st11GoodNo, m.st11price, m.st11SellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, isNULL(m.st11StatCD,-9) as st11StatCD, tm.safeDiv, tm.isNeed, tm.depth1Code, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000)) "
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : �Ϲ�, 06 : �ֹ�����(����), 16 : �ֹ�����, 07 : ��������
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or (( "
		strSql = strSql & "			convert(varchar(6), (i.cate_large + i.cate_mid)) in ( "
		strSql = strSql & "				SELECT convert(varchar(6), cdl+cdm)  "
		strSql = strSql & "				FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category]  "	'2023-06-23 ������ / ������� ī�װ��� Ư���귣��� �Ǹŵǵ���
		strSql = strSql & "				WHERE mallgubun='11st1010') "
		strSql = strSql & "			) and ( "
		strSql = strSql & "				i.makerid not in ( "
		strSql = strSql & "					'heidi2022', "
		strSql = strSql & "					'luna2022', "
		strSql = strSql & "					'uand2051', "
		strSql = strSql & "					'wpc001', "
		strSql = strSql & "					'lifeshop0510', "
		strSql = strSql & "					'bijou2023', "
		strSql = strSql & "					'blesscompany', "
		strSql = strSql & "					'JINNYSTAR01', "
		strSql = strSql & "					'sportsconnection', "
		strSql = strSql & "					'ithinkso', "
		strSql = strSql & "					'greenh03', "
		strSql = strSql & "					'kingkongoutlet', "
		strSql = strSql & "					'shoemiz', "
		strSql = strSql & "					'goldn', "
		strSql = strSql & "					'funnyfun', "
		strSql = strSql & "					'doran1020', "
		strSql = strSql & "					'gabangpop1010', "
		strSql = strSql & "					'osjarak' "
		strSql = strSql & "				) "
		strSql = strSql & "			) "
		strSql = strSql & "		) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.st11GoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FtenCateLarge		= rsget("cate_large")
				FOneItem.FtenCateMid		= rsget("cate_mid")
				FOneItem.FtenCateSmall		= rsget("cate_small")
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FsmallImage		= rsget("smallImage")
				FOneItem.Fmakerid			= rsget("makerid")
				FOneItem.Fregdate			= rsget("regdate")
				FOneItem.FlastUpdate		= rsget("lastUpdate")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FsaleYn			= rsget("sailyn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.Fkeywords			= rsget("keywords")
				FOneItem.ForderComment		= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSt11GoodNo		= rsget("st11GoodNo")
				FOneItem.FSt11price			= rsget("st11price")
				FOneItem.FSt11SellYn		= rsget("st11SellYn")

				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv       	= rsget("infoDiv")
	            FOneItem.Fsafetyyn      	= rsget("safetyyn")
	            FOneItem.FsafetyDiv     	= rsget("safetyDiv")
	            FOneItem.FsafetyNum     	= rsget("safetyNum")
	            FOneItem.FmaySoldOut    	= rsget("maySoldOut")
	            FOneItem.Fregitemname   	= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
                FOneItem.FDepthCode			= rsget("depthCode")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

End Class

'11���� ��ǰ�ڵ� ���
Function get11stGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 st11goodno FROM db_etcmall.dbo.tbl_11st_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		get11stGoodno = rsget("st11goodno")
	End If
	rsget.Close
End Function

'11���� ��ǰ�ڵ�/��ǰ�� ���
Function get11stGoodno2(iitemid, ist11goodno, byref MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, saleyn, tmpPrice, vdeliverytype, ispecialPrice, outmallstandardMargin
	Dim GetTenTenMargin, st11goodno, ownItemCnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.buycash, i.orgprice, i.sailyn, r.st11goodno, i.deliverytype, isnull(mi.mustPrice, 0) as specialPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as mi "
	strSql = strSql & " 	on i.itemid = mi.itemid "
	strSql = strSql & " 	and mi.mallgubun = '11st1010' "
	strSql = strSql & " 	and (GETDATE() >= mi.startDate and GETDATE() <= mi.endDate ) "
	strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		saleyn		= rsget("sailyn")
		st11goodno	= rsget("st11goodno")
		vdeliverytype = rsget("deliverytype")
		ispecialPrice = rsget("specialPrice")
		outmallstandardMargin = rsget("outmallstandardMargin")
	Else
		get11stGoodno2 = ""
		Exit Function
		response.end
	End If
	rsget.close

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as CNT "
	strSql = strSql & " FROM db_item.dbo.tbl_item i "
	strSql = strSql & " JOIN db_partner.dbo.tbl_partner p on i.makerid = p.id "
	strSql = strSql & " WHERE p.purchaseType in (3, 5, 6) "		'3 : PB, 5 : ODM, 6 : ����
	strSql = strSql & " and i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ownItemCnt = rsget("CNT")
	End If
	rsget.Close

	If ispecialPrice <> "0" Then
		tmpPrice = ispecialPrice
	ElseIf ownItemCnt > 0 Then
		tmpPrice = orgprice
	Else
		GetTenTenMargin = CLng((10000 - buycash / sellcash * 100 * 100) / 100)
		If (GetTenTenMargin < outmallstandardMargin) Then
			tmpPrice = orgprice
		Else
			tmpPrice = sellcash
		End If
	End If
	MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	ist11goodno = st11goodno
End Function

'11���� ��ǰ�ڵ�, �ɼ� �� ���
Function get11stGoodno3(iitemid, ist11goodno, byref opCnt)
	Dim strSql, st11goodno, optioncnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.optioncnt, r.st11goodno "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		opCnt		= rsget("optioncnt")
		ist11goodno	= rsget("st11goodno")
	Else
		get11stGoodno3 = ""
		Exit Function
		response.end
	End If
	rsget.close
End Function

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Public Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

function replaceRst(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, """", "&quot;")
	'v = replace(v, "&", "&amp;")

	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function

function replaceMsg(v)
	if IsNull(v) then
		replaceMsg = ""
		Exit function
	end if
	v = Replace(v, vbcrlf,"")
	v = Replace(v, vbCr,"")
	v = Replace(v, vbLf,"")
    replaceMsg = v
end function
%>
