<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "hmall1010"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST HMALLMARGIN = 11
CONST CDEFALUT_STOCK = 9999

Class CHmallItem
	Public FItemid
	Public Fitemname
	Public FsmallImage
	Public Fmakerid
	Public Fregdate
	Public FlastUpdate
	Public ForgPrice
	Public FSellCash
	Public FBuyCash
	Public FsellYn
	Public FsaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FHmallRegdate
	Public FHmallLastUpdate
	Public FHmallGoodNo
	Public FHmallPrice
	Public FoctyCnryGbcd
	Public FoctyCnryNm
	Public FitemLCsfCd
	Public FitemMCsfCd
	Public FitemSCsfCd
	Public FitemCsfGbcd
	Public Fitemsize	
	Public Fitemsource
	Public FHmallSellYn
	Public FMrgnRate
	Public FregUserid
	Public FHmallStatCd
	Public FCateMapCnt
	Public Fdeliverytype
	Public Fdefaultdeliverytype
	Public FdefaultfreeBeasongLimit
	Public FoptionCnt
	Public FregedOptCnt
	Public FrctSellCNT
	Public FaccFailCNT
	Public FlastErrStr
	Public FinfoDiv
	Public FoptAddPrcCnt
	Public FoptAddPrcRegType
	Public FitemDiv
	Public FMetaOption
	Public FMallinfoDiv
	Public FOutboundShippingPlaceCode
	Public FProductId
	Public ForgSuplyCash
	Public Fisusing
	Public Fkeywords
	Public Fvatinclude
	Public FordMakeYn
	Public ForderComment
	Public FAdultType
	Public FbasicImage
	Public FbasicimageNm
	Public FmainImage
	Public FmainImage2
	Public Ficon2Image
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FCateKey
	Public FDepth1Name
	Public FDepth2Name
	Public FDepth3Name
	Public FDepth4Name
	Public FDepth5Name
	Public FDepth6Name

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName

	Public FId
	Public FSocname_kor
	Public FDeliverPhone
	Public FSocname
	Public FDeliver_name
	Public FReturn_zipcode
	Public FReturn_address
	Public FReturn_address2
	Public FDivname
	Public FMaeipdiv
	Public FJeju
	Public FNotJeju
	Public FDefaultSongjangDiv

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt
		limitYCnt = 0
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
				If (FLimitYN <> "Y") Then optLimit = 999

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

	Public Function getSafetyParam()
		Dim strSql, isCertYn, safeCertTypeGbcd, safetyDiv, gbnFlag
		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, t.certNum "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " WHERE i.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			isCertYn	= "Y"
			safetyDiv	= rsget("safetyDiv")
		Else
			isCertYn	= "N"
			safeCertTypeGbcd = "N"
		End If
		rsget.Close

		If isCertYn = "Y" Then
			Select Case safetyDiv
				Case "10", "40", "70"		safeCertTypeGbcd = "01"
				Case "20", "50", "80"		safeCertTypeGbcd = "02"
				Case "30", "60", "90"		safeCertTypeGbcd = "03"
			End Select

			Select Case safetyDiv
				Case "10", "20", "30"		gbnFlag = "elec"
				Case "40", "50", "60"		gbnFlag = "life"
				Case "70", "80", "90"		gbnFlag = "child"
			End Select
		End If
		getSafetyParam = isCertYn&"|_|"&safeCertTypeGbcd&"|_|"&gbnFlag
	End Function

	Public Function IsAllOptionChange
		Dim sqlStr, tenOptCnt, regedHmallOptCnt, addPriceCnt
		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			tenOptCnt = rsget("cnt")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and optaddprice > 0 "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			addPriceCnt = rsget("cnt")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_etcmall.[dbo].[tbl_hmall_regedOption]  "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and outmallOptName <> '���Ͽɼ�' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			regedHmallOptCnt = rsget("cnt")
		rsget.Close

		If tenOptCnt > 0 AND regedHmallOptCnt = 0 Then			'��ǰ -> �ɼ�
			IsAllOptionChange = "Y"
		ElseIf tenOptCnt = 0 AND regedHmallOptCnt > 0 Then		'�ɼ� -> ��ǰ
			IsAllOptionChange = "Y"
		ElseIf addPriceCnt > 0 Then								'�ɼ��߰��ݾ� ���ٰ� ���� ���
			IsAllOptionChange = "Y"
		Else
			IsAllOptionChange = "N"
		End If
	End Function

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

	'���� ���� �귣�� ���� üũ
	Public Function fnCheckMakerid()
		Dim strSql, cntMakerId

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_partner.dbo.tbl_partner_group G "
		strSql = strSql & " JOIN db_partner.dbo.tbl_partner as P on G.groupid = p.groupid "
		strSql = strSql & " WHERE G.jungsan_gubun = '���̰���' "
		strSql = strSql & " and p.id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			cntMakerId = rsget("cnt")
		End If
		rsget.Close

		If cntMakerId > 0 Then
			fnCheckMakerid = True
		Else
			fnCheckMakerid = False
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

	'// hmall �Ǹſ��� ��ȯ
	Public Function gethmallSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				gethmallSellYn = "Y"
			Else
				gethmallSellYn = "N"
			End If
		Else
			gethmallSellYn = "N"
		End If
	End Function

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[���� "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[��ü����]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[��ü]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<br><font color=blue>����:"&getLimitEa&"</font>"
	    End if
	End Function

	Function getLimitEa()
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
		getLimitEa = ret
	End Function

	Function getOptionLimitEa(ino, isold)
		dim ret : ret = (ino - isold - 5)
		if (ret < 1) then ret=0
		If (ret >= 1000) Then ret = 999
		getOptionLimitEa = ret
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
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
			MustPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			MustPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If (GetTenTenMargin < CMAXMARGIN) Then
				MustPrice = CStr(GetRaiseValue(Forgprice/10)*10)
			Else
				If (FSellCash < Round(FHmallPrice * 0.55, 0)) Then
					MustPrice = CStr(GetRaiseValue(Round(FHmallPrice * 0.55, 0)/10)*10)
				Else
					MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
				End If
			End If
		End If
	End Function

	Public Function getBasicImage()
		If IsNULL(FbasicImageNm) or (FbasicImageNm="") Then Exit function
		getBasicImage = FbasicImageNm
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

	Public Function getHmallContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '����','����' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_hmall.jpg'></p><br>"

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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_hmall.jpg"">")
		getHmallContParamToReg = strRst

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
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_hmall.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_hmall.jpg"">")
			getHmallContParamToReg = strRst
		End If
		rsget.Close
	End Function

	Public Function getAttrInfo()
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		chkMultiOpt = false

		If FoptionCnt = 0 Then
			strRst = ""
			strRst = strRst & "					<uitmCombYn>N</uitmCombYn>"									'��ǰ�Ӽ����տ���
			strRst = strRst & "					<uitm1AttrTypeNm><![CDATA[���Ͽɼ�]]></uitm1AttrTypeNm>"	'��ǰ�Ӽ�1�Ӽ�������
			strRst = strRst & "					<uitm2AttrTypeNm></uitm2AttrTypeNm>"						'��ǰ�Ӽ�2�Ӽ�������
			strRst = strRst & "					<uitm3AttrTypeNm></uitm3AttrTypeNm>"						'��ǰ�Ӽ�3�Ӽ�������
			strRst = strRst & "					<uitm4AttrTypeNm></uitm4AttrTypeNm>"						'��ǰ�Ӽ�4�Ӽ�������
			strRst = strRst & "					<uitmChocPossYn>N</uitmChocPossYn>"							'��ǰ�Ӽ����ð��ɿ���
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			If chkMultiOpt = True Then
				optionTypeName1 = ""
				optionTypeName2 = ""
				optionTypeName3 = ""
				optionTypeName4 = ""
				strSql = ""
				strSql = strSql & " SELECT typeseq, optionTypeName From db_item.[dbo].[tbl_item_option_Multiple] "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " GROUP BY typeseq, optionTypeName "
				strSql = strSql & " ORDER BY Typeseq "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.Eof Then
					Do until rsget.EOF
						Select Case rsget("typeseq")
							Case "1"		optionTypeName1 = rsget("optionTypeName")
							Case "2"		optionTypeName2 = rsget("optionTypeName")
							Case "3"		optionTypeName3 = rsget("optionTypeName")
							Case "4"		optionTypeName4 = rsget("optionTypeName")
						End Select
						rsget.MoveNext
					Loop
				End If
				rsget.close

				strRst = ""
				strRst = strRst & "					<uitmCombYn>Y</uitmCombYn>"									'��ǰ�Ӽ����տ���
				strRst = strRst & "					<uitm1AttrTypeNm><![CDATA["&optionTypeName1&"]]></uitm1AttrTypeNm>"	'��ǰ�Ӽ�1�Ӽ�������
				strRst = strRst & "					<uitm2AttrTypeNm><![CDATA["&optionTypeName2&"]]></uitm2AttrTypeNm>"						'��ǰ�Ӽ�2�Ӽ�������
				strRst = strRst & "					<uitm3AttrTypeNm><![CDATA["&optionTypeName3&"]]></uitm3AttrTypeNm>"						'��ǰ�Ӽ�3�Ӽ�������
				strRst = strRst & "					<uitm4AttrTypeNm><![CDATA["&optionTypeName4&"]]></uitm4AttrTypeNm>"						'��ǰ�Ӽ�4�Ӽ�������
				strRst = strRst & "					<uitmChocPossYn>Y</uitmChocPossYn>"							'��ǰ�Ӽ����ð��ɿ���
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 optionTypeName "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " GROUP BY optionTypeName "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					strRst = ""
					strRst = strRst & "					<uitmCombYn>Y</uitmCombYn>"									'��ǰ�Ӽ����տ���
					strRst = strRst & "					<uitm1AttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitm1AttrTypeNm>"	'��ǰ�Ӽ�1�Ӽ�������
					strRst = strRst & "					<uitm2AttrTypeNm></uitm2AttrTypeNm>"						'��ǰ�Ӽ�2�Ӽ�������
					strRst = strRst & "					<uitm3AttrTypeNm></uitm3AttrTypeNm>"						'��ǰ�Ӽ�3�Ӽ�������
					strRst = strRst & "					<uitm4AttrTypeNm></uitm4AttrTypeNm>"						'��ǰ�Ӽ�4�Ӽ�������
					strRst = strRst & "					<uitmChocPossYn>Y</uitmChocPossYn>"							'��ǰ�Ӽ����ð��ɿ���
				End If
				rsget.Close
			End If
		End If
		getAttrInfo = strRst
	End Function

	Public Function getOptSellUitmDtl
		Dim strSql, strRst, i, chkMultiOpt, j, chkqty
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, commaCount
		chkqty = ""
		chkMultiOpt = false

		If FoptionCnt = 0 Then
			buf = ""
			buf = buf & "	<Dataset id=""dsSellUitmDtl"">"											'#�ǸŻ�ǰ�Ӽ�����
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>INSERT</rowType>"
			buf = buf & "				<chk>1</chk>"												'�Ӽ����翩�� | ��ǰ�ΰ�츸 ��� , Default :1  (��ǰ����� ����� �ȵɶ��� ����ϼ���, �Ϲ������� chk���̵� ��ǰ ����� ���������� �˴ϴ�.)
			buf = buf & "				<bsitmCd></bsitmCd>"										'���ػ�ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���), null ��
			buf = buf & "				<slitmCd></slitmCd>"										'��ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , null ��
			buf = buf & "				<uitmCd></uitmCd>"											'��ǰ�Ӽ��ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , �Ӽ��ڵ�(uitmTmpCd)�� �Է�
			buf = buf & "				<uitmTmpCd>0</uitmTmpCd>"									'#�Ӽ��ڵ� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<uitm1AttrNm><![CDATA[���Ͽɼ�]]></uitm1AttrNm>"			'#��ǰ�Ӽ�1�Ӽ������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<uitm2AttrNm><![CDATA[���Ͽɼ�]]></uitm2AttrNm>"			'#��ǰ�Ӽ�2�Ӽ������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#��ǰ�Ӽ�3�Ӽ������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#��ǰ�Ӽ�4�Ӽ������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#�ǸŽ������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#�Ǹ��������� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<uitmTotNm><![CDATA[���Ͽɼ�]]></uitmTotNm>"				'#��ǰ�Ӽ���ü�� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<addQty>0</addQty>"											'#�߰����� |	��ǰ/�����ΰ�� ���
			buf = buf & "				<maxSellPossQty>"&getLimitEa()&"</maxSellPossQty>"			'#�ִ��ǸŰ��ɼ��� | ��ǰ/�����ΰ�� ���
			buf = buf & "				<sellGbcd>00</sellGbcd>"									'#�Ǹű����ڵ� | 00 ���� <- (����ΰ�� �⺻������ ���ÿ��), 11 �Ͻ��ߴ�, 19 �����ߴ�
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			buf = ""
			buf = buf & "	<Dataset id=""dsSellUitmDtl"">"											'#�ǸŻ�ǰ�Ӽ�����
			buf = buf & "		<rows>"

			If chkMultiOpt = True Then
				j = 0

				strSql = ""
				strSql = strSql & " SELECT optionname, optlimitno, optlimitsold "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE itemid = " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						commaCount = Ubound(Split(("optionname"), ","))
						buf = buf & "			<row>"
						buf = buf & "				<rowType>INSERT</rowType>"
						buf = buf & "				<chk>1</chk>"												'�Ӽ����翩�� | ��ǰ�ΰ�츸 ��� , Default :1  (��ǰ����� ����� �ȵɶ��� ����ϼ���, �Ϲ������� chk���̵� ��ǰ ����� ���������� �˴ϴ�.)
						buf = buf & "				<bsitmCd></bsitmCd>"										'���ػ�ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���), null ��
						buf = buf & "				<slitmCd></slitmCd>"										'��ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , null ��
						buf = buf & "				<uitmCd></uitmCd>"											'��ǰ�Ӽ��ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , �Ӽ��ڵ�(uitmTmpCd)�� �Է�
						buf = buf & "				<uitmTmpCd>"&j&"</uitmTmpCd>"								'#�Ӽ��ڵ� | ��ǰ/�����ΰ�� ���

						buf = buf & "				<uitm1AttrNm><![CDATA["&Split(rsget("optionname"), ",")(0)&"]]></uitm1AttrNm>"			'#��ǰ�Ӽ�1�Ӽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitm2AttrNm><![CDATA["&Split(rsget("optionname"), ",")(1)&"]]></uitm2AttrNm>"			'#��ǰ�Ӽ�2�Ӽ������� | ��ǰ/�����ΰ�� ���
						If Ubound(Split(rsget("optionname"), ",")) = 2 Then
							buf = buf & "				<uitm3AttrNm><![CDATA["&Split(rsget("optionname"), ",")(2)&"]]></uitm3AttrNm>"								'#��ǰ�Ӽ�3�Ӽ������� | ��ǰ/�����ΰ�� ���
						Else
							buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#��ǰ�Ӽ�3�Ӽ������� | ��ǰ/�����ΰ�� ���
						End If

						If Ubound(Split(rsget("optionname"), ",")) = 3 Then
							buf = buf & "				<uitm4AttrNm><![CDATA["&Split(rsget("optionname"), ",")(3)&"]]></uitm4AttrNm>"								'#��ǰ�Ӽ�3�Ӽ������� | ��ǰ/�����ΰ�� ���
						Else
							buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#��ǰ�Ӽ�4�Ӽ������� | ��ǰ/�����ΰ�� ���
						End If
						buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#�ǸŽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#�Ǹ��������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitmTotNm><![CDATA["& Replace(rsget("optionname"), ",", "/") &"]]></uitmTotNm>"				'#��ǰ�Ӽ���ü�� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<addQty>0</addQty>"											'#�߰����� |	��ǰ/�����ΰ�� ���
						If FLimityn = "Y" Then
							If rsget("optlimitno") - rsget("optlimitsold") - 5 < 0 Then
								chkqty = 0
							Else
								chkqty = rsget("optlimitno") - rsget("optlimitsold") - 5
							End If
							buf = buf & "				<maxSellPossQty>"&chkqty&"</maxSellPossQty>"			'#�ִ��ǸŰ��ɼ��� | ��ǰ/�����ΰ�� ���
						Else
							buf = buf & "				<maxSellPossQty>9999</maxSellPossQty>"						'#�ִ��ǸŰ��ɼ��� | ��ǰ/�����ΰ�� ���
						End If
						buf = buf & "				<sellGbcd>00</sellGbcd>"									'#�Ǹű����ڵ� | 00 ���� <- (����ΰ�� �⺻������ ���ÿ��), 11 �Ͻ��ߴ�, 19 �����ߴ�
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				rsget.close
			Else
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						buf = buf & "			<row>"
						buf = buf & "				<rowType>INSERT</rowType>"
						buf = buf & "				<chk>1</chk>"												'�Ӽ����翩�� | ��ǰ�ΰ�츸 ��� , Default :1  (��ǰ����� ����� �ȵɶ��� ����ϼ���, �Ϲ������� chk���̵� ��ǰ ����� ���������� �˴ϴ�.)
						buf = buf & "				<bsitmCd></bsitmCd>"										'���ػ�ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���), null ��
						buf = buf & "				<slitmCd></slitmCd>"										'��ǰ�ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , null ��
						buf = buf & "				<uitmCd></uitmCd>"											'��ǰ�Ӽ��ڵ� | ��ǰ�ΰ�츸 ���(��ǰ����� ����� �ȵɶ��� ����ϼ���) , �Ӽ��ڵ�(uitmTmpCd)�� �Է�
						buf = buf & "				<uitmTmpCd>"&j&"</uitmTmpCd>"								'#�Ӽ��ڵ� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitm1AttrNm><![CDATA["&rsget("typename")&"]]></uitm1AttrNm>"			'#��ǰ�Ӽ�1�Ӽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitm2AttrNm><![CDATA["&rsget("kindname")&"]]></uitm2AttrNm>"			'#��ǰ�Ӽ�2�Ӽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#��ǰ�Ӽ�3�Ӽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#��ǰ�Ӽ�4�Ӽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#�ǸŽ������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#�Ǹ��������� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<uitmTotNm><![CDATA["&rsget("typename")&"/"&rsget("kindname")&"]]></uitmTotNm>"				'#��ǰ�Ӽ���ü�� | ��ǰ/�����ΰ�� ���
						buf = buf & "				<addQty>0</addQty>"											'#�߰����� |	��ǰ/�����ΰ�� ���
						If FLimityn = "Y" Then
							If rsget("limitno") - rsget("limitsold") - 5 < 0 Then
								chkqty = 0
							Else
								chkqty = rsget("limitno") - rsget("limitsold") - 5
							End If
							buf = buf & "				<maxSellPossQty>"&chkqty&"</maxSellPossQty>"			'#�ִ��ǸŰ��ɼ��� | ��ǰ/�����ΰ�� ���
						Else
							buf = buf & "				<maxSellPossQty>9999</maxSellPossQty>"						'#�ִ��ǸŰ��ɼ��� | ��ǰ/�����ΰ�� ���
						End If
						buf = buf & "				<sellGbcd>00</sellGbcd>"									'#�Ǹű����ڵ� | 00 ���� <- (����ΰ�� �⺻������ ���ÿ��), 11 �Ͻ��ߴ�, 19 �����ߴ�
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				rsget.close
			End If
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptSellUitmDtl = buf
	End Function

	Public Function getOptTypeMst(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false
		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'��ǰ�Ӽ��Ӽ���������
			buf = buf & "				<uitmAttrTypeNm><![CDATA[���Ͽɼ�]]></uitmAttrTypeNm>"							'��ǰ�Ӽ��Ӽ�������
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			If chkMultiOpt = True Then
				buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
				buf = buf & "		<rows>"
				strSql = ""
				strSql = strSql & " SELECT typeseq, optionTypeName From db_item.[dbo].[tbl_item_option_Multiple] "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " GROUP BY typeseq, optionTypeName "
				strSql = strSql & " ORDER BY Typeseq "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.Eof Then
					Do until rsget.EOF
						Select Case rsget("typeseq")
							Case "1"		optionTypeName1 = rsget("optionTypeName")
							Case "2"		optionTypeName2 = rsget("optionTypeName")
							Case "3"		optionTypeName3 = rsget("optionTypeName")
							Case "4"		optionTypeName4 = rsget("optionTypeName")
						End Select
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmAttrTypeSeq>"&rsget("typeseq")&"</uitmAttrTypeSeq>"	'��ǰ�Ӽ��Ӽ���������
						buf = buf & "				<uitmAttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitmAttrTypeNm>"							'��ǰ�Ӽ��Ӽ�������
						buf = buf & "			</row>"
						rsget.MoveNext
					Loop
				End If
				rsget.close
				buf = buf & "		</rows>"
				buf = buf & "	</Dataset>"
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 optionTypeName "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " GROUP BY optionTypeName "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
					buf = buf & "		<rows>"
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'��ǰ�Ӽ��Ӽ���������
					buf = buf & "				<uitmAttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitmAttrTypeNm>"							'��ǰ�Ӽ��Ӽ�������
					buf = buf & "			</row>"
					buf = buf & "		</rows>"
					buf = buf & "	</Dataset>"
				End If
				rsget.Close
			End If
		End If
		getOptTypeMst = buf
	End Function

	Public Function getOptAttrMst(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, j, rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false

		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsUitmAttrMst"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmTmpSeq>0</uitmTmpSeq>"									'�Ӽ�����
			buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'��ǰ�Ӽ��Ӽ���������
			buf = buf & "				<uitmAttrNm></uitmAttrNm>"									'��ǰ�Ӽ���
			buf = buf & "				<uitmCreYn></uitmCreYn>"									'��ǰ�Ӽ���������
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			buf = buf & "	<Dataset id=""dsUitmAttrMst"">"
			buf = buf & "		<rows>"
			j = 0
			strSql = ""
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<uitmTmpSeq>"&j&"</uitmTmpSeq>"								'�Ӽ�����
					buf = buf & "				<uitmAttrTypeSeq>"&rsget("TypeSeq")&"</uitmAttrTypeSeq>"	'��ǰ�Ӽ��Ӽ���������
					buf = buf & "				<uitmAttrNm><![CDATA["&rsget("kindname")&"]]></uitmAttrNm>"	'��ǰ�Ӽ���
					buf = buf & "				<uitmCreYn></uitmCreYn>"									'��ǰ�Ӽ���������
					buf = buf & "			</row>"
					j = j + 1
					rsget.MoveNext
				Loop
			End If
			rsget.Close
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptAttrMst = buf
	End Function

	Public Function getOptCombDtl(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, j, rowType, tmpOption

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false

		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsSellUitmCombDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmTmpCd>0</uitmTmpCd>"									'�Ӽ��ڵ�
			buf = buf & "				<uitmTmpSeq>0</uitmTmpSeq>"									'�Ӽ�����
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close
			buf = buf & "	<Dataset id=""dsSellUitmCombDtl"">"
			buf = buf & "		<rows>"
			If chkMultiOpt = True Then
				'��Ƽ�ɼ��� �� �ؾ���..........
				i = 0
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get2] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						If j = 0 Then
							tmpOption = rsget("itemoption")
						End If

						If tmpOption <> rsget("itemoption") Then
							tmpOption = rsget("itemoption")
							i = i + 1
						End If
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmTmpCd>"&i&"</uitmTmpCd>"												'�Ӽ��ڵ�
						buf = buf & "				<uitmTmpSeq>"&rsget("rnum")&"</uitmTmpSeq>"									'�Ӽ�����
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				 rsget.close
			Else
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmTmpCd>"&rsget("rnum")&"</uitmTmpCd>"									'�Ӽ��ڵ�
						buf = buf & "				<uitmTmpSeq>"&rsget("rnum")&"</uitmTmpSeq>"									'�Ӽ�����
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				 rsget.close
			End If
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptCombDtl = buf
	End Function

    Public Function getCertOrganName(icertOrganName)
		Select Case icertOrganName
			Case Instr(icertOrganName, "FITI���迬����") > 0				getCertOrganName = "8"
			Case Instr(icertOrganName, "�ѱ�ȭ�����ս��迬����") > 0		getCertOrganName = "4"
			Case Instr(icertOrganName, "�ѱ�����������ڽ��迬����") > 0	getCertOrganName = "10"
			Case Instr(icertOrganName, "KOTITI ���迬����") > 0				getCertOrganName = "14"
			Case Instr(icertOrganName, "�ѱ��Ǽ���Ȱ���迬����") > 0		getCertOrganName = "5"
			Case Instr(icertOrganName, "�ѱ��Ƿ����迬����") > 0			getCertOrganName = "7"
			Case Instr(icertOrganName, "�ѱ������������") > 0			getCertOrganName = "3"
			Case Instr(icertOrganName, "�ѱ��Ǽ���Ȱȯ����迬����") > 0	getCertOrganName = "5"
		End Select
    End function

	Public Function getHmallItemSafeInfoToReg(gbcd, gbn)
		Dim buf
		Dim strSql, safetyDiv, certNum, certOrganName, modelName, certDate
		Dim safeCertLawGbcd, safeCertTypeGbcd, safeCertNo, safeCrtiGbcd, speCate
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		speCate = "N"
		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isNull(t.certNum, '') as certNum, isNull(f.modelName, '') as modelName, isNull(f.certDate, '') as certDate, isNull(f.certOrganName, '') as certOrganName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " LEFT JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
			certOrganName = rsget("certOrganName")
			modelName	= rsget("modelName")
			certDate	= rsget("certDate")
		End If
		rsget.Close

		Select Case safetyDiv
			Case "10"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "01"
				safeCertNo				= certNum
			Case "20"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "30"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "03"
			Case "40"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "02"				'��Ȱȭ���� ����Ȯ�θ���..
				safeCertNo				= certNum
			Case "50"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "60"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "03"
			Case "70"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "01"
				safeCertNo				= certNum
			Case "80"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "90"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "03"

		End Select
		safeCrtiGbcd = getCertOrganName(certOrganName)

		Select Case FitemLCsfCd
			Case "R6"	
				safeCertLawGbcd = "10"
				speCate =  "Y"
		End Select

		If speCate="Y" and gbcd = "03" Then
			buf = ""
			buf = buf & "	<Dataset id=""dsItemSafeCertMngDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<safeCertLawGbcd>10</safeCertLawGbcd>"		'���������������ڵ� | 10 ��̾���Ư����, 20 ���������, 30 ��Ȱȭ�о�����, 40 �����ű������
			buf = buf & "				<safeCertTypeGbcd>03</safeCertTypeGbcd>"	'�����������������ڵ� | 01 ��������, 02 ����Ȯ��, 03 ���������ռ�Ȯ��, 04 ���������ؼ����, 05 ��������, 06 ���յ��, 07 ��������
			buf = buf & "				<safeCertDt></safeCertDt>"						'������������ | YYYYMMDD ����
			buf = buf & "				<safeCertNo></safeCertNo>"					'����������ȣ
			buf = buf & "				<safeCrtiGbcd></safeCrtiGbcd>"				'����������������ڵ� | 1 �ѱ��������ڽ��迬����, 2 �ѱ������Ŀ�����, 3 �ѱ������������, 4 �ѱ�ȭ�н��迬����, 5 �ѱ���Ȱȯ����迬����, 6 �ѱ������ȭ���迬����, 7 �ѱ��Ƿ����迬����, 8 FITI���迬����, 9 �ѱ�ǥ����ȸ, 10 �ѱ�����������ڽ��迬����, 11 �ľ�û, 12 ����������ȸ, 13 �������Ŀ���ȸ, 14 KOTITI ���迬����, 
			buf = buf & "				<safeCertClasGbcd>1</safeCertClasGbcd>"						'���������׸񱸺��ڵ� | 0 ����, 1 ���������ڵ�(KC), 2 ����ǰ�����˻�, 3 ��������������ǰ, 4 ��������Ȯ�δ�����ǰ, 5 �����ǰ��������, 6 �Ƿ�������ǰ���㰡, 7 ����Ȯ�δ�������ǰ, 8 ���ؿ����ǰ
			buf = buf & "				<safeCertImgNm></safeCertImgNm>"							'���������̹�����
			buf = buf & "				<certInfIdnfYn>Y</certInfIdnfYn>"							'��������Ȯ�ο��� | Y ����, N ����
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			buf = ""
			buf = buf & "	<Dataset id=""dsItemSafeCertMngDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<safeCertLawGbcd>"&safeCertLawGbcd&"</safeCertLawGbcd>"		'���������������ڵ� | 10 ��̾���Ư����, 20 ���������, 30 ��Ȱȭ�о�����, 40 �����ű������
			buf = buf & "				<safeCertTypeGbcd>"&safeCertTypeGbcd&"</safeCertTypeGbcd>"	'�����������������ڵ� | 01 ��������, 02 ����Ȯ��, 03 ���������ռ�Ȯ��, 04 ���������ؼ����, 05 ��������, 06 ���յ��, 07 ��������
			buf = buf & "				<safeCertDt>"&certDate&"</safeCertDt>"						'������������ | YYYYMMDD ����
			buf = buf & "				<safeCertNo>"&safeCertNo&"</safeCertNo>"					'����������ȣ
			buf = buf & "				<safeCrtiGbcd>"&safeCrtiGbcd&"</safeCrtiGbcd>"				'����������������ڵ� | 1 �ѱ��������ڽ��迬����, 2 �ѱ������Ŀ�����, 3 �ѱ������������, 4 �ѱ�ȭ�н��迬����, 5 �ѱ���Ȱȯ����迬����, 6 �ѱ������ȭ���迬����, 7 �ѱ��Ƿ����迬����, 8 FITI���迬����, 9 �ѱ�ǥ����ȸ, 10 �ѱ�����������ڽ��迬����, 11 �ľ�û, 12 ����������ȸ, 13 �������Ŀ���ȸ, 14 KOTITI ���迬����, 
			buf = buf & "				<safeCertClasGbcd>1</safeCertClasGbcd>"						'���������׸񱸺��ڵ� | 0 ����, 1 ���������ڵ�(KC), 2 ����ǰ�����˻�, 3 ��������������ǰ, 4 ��������Ȯ�δ�����ǰ, 5 �����ǰ��������, 6 �Ƿ�������ǰ���㰡, 7 ����Ȯ�δ�������ǰ, 8 ���ؿ����ǰ
			buf = buf & "				<safeCertImgNm></safeCertImgNm>"							'���������̹�����
			buf = buf & "				<certInfIdnfYn></certInfIdnfYn>"							'��������Ȯ�ο��� | Y ����, N ����
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getHmallItemSafeInfoToReg = buf
	End Function

	Function getHmallItemInfoCdToReg(gbn)
		Dim strSql, buf
		Dim mallinfoCd,infoContent,infotype
		Dim itstCd, itstGbcd, itstTitl, itstCntn
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		buf = ""
		buf = buf & "	<Dataset id=""dsItstDtl"">"
		buf = buf & "		<rows>"

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
			    itstCd		= rsget("itstCd")
			    itstGbcd	= rsget("itstGbcd")
			    itstTitl	= rsget("itstTitl")
				itstCntn	= rsget("itstCntn")

			    If Not (IsNULL(itstCntn)) AND (itstCntn <> "") Then
			    	itstCntn = replace(itstCntn, chr(31), "")
				End If

				buf = buf & "			<row>"
				buf = buf & "				<rowType>"&rowType&"</rowType>"
				buf = buf & "				<itstCd>"&itstCd&"</itstCd>"											'��ǰ������ڵ�
				buf = buf & "				<itstGbcd>"&itstGbcd&"</itstGbcd>"										'��ǰ����������ڵ� | 10 ������ÿ�, 20 ��ǰ������� (��ǰ��������׸��� ��ǰ�ʼ����� ��ȸ API�� �̿��Ͽ� ��ȸ�Ѵ�.)
				buf = buf & "				<itstTitl><![CDATA["&itstTitl&"]]></itstTitl>"							'��ǰ���������	String(200)	
				buf = buf & "				<itstCntn><![CDATA["&itstCntn&"]]></itstCntn>"							'��ǰ���������	String(4000)	
				buf = buf & "			</row>"
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		buf = buf & "		</rows>"
		buf = buf & "	</Dataset>"
		getHmallItemInfoCdToReg = buf
	End Function

	Public Function getHmallSectIdToReg(gbn)
		Dim buf, strSql
		Dim sectAttrGbcd, sectId1, sectId2
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		buf = ""
		buf = buf & "	<Dataset id=""dsDispItemDtl"">"
		buf = buf & "		<rows>"

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_SpecialCategoryMapping_Get] '"& FtenCateLarge &"', '"& FtenCateMid &"' "
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				sectAttrGbcd = ""
				sectId1 = ""
				sectId2 = ""

			    sectAttrGbcd	= rsget("sectAttrGbcd")
			    sectId1			= rsget("sectId1")
			    sectId2			= rsget("sectId2")

				If sectId1 <> "" Then
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<sectAttrGbcd>"&sectAttrGbcd&"</sectAttrGbcd>"				'����Ӽ������ڵ� | 01 �Ϲݸ���
					buf = buf & "				<sectId>"&sectId1&"</sectId>"								'����ID | ��ǰ�� ���� ������ ���ؼ��� 1�� �̻��� Ȱ��ȭ ���� ����� �ʿ���. ��ǰ�� Ȱ��ȭ ������ ��ϵǾ� ���� ���� ��� �ش� �����ͼ��� ���� �߰� ��� ������
					buf = buf & "			</row>"
				End If

				If sectId2 <> "" Then
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<sectAttrGbcd>"&sectAttrGbcd&"</sectAttrGbcd>"				'����Ӽ������ڵ� | 01 �Ϲݸ���
					buf = buf & "				<sectId>"&sectId2&"</sectId>"								'����ID | ��ǰ�� ���� ������ ���ؼ��� 1�� �̻��� Ȱ��ȭ ���� ����� �ʿ���. ��ǰ�� Ȱ��ȭ ������ ��ϵǾ� ���� ���� ��� �ش� �����ͼ��� ���� �߰� ��� ������
					buf = buf & "			</row>"
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		buf = buf & "		</rows>"
		buf = buf & "	</Dataset>"
		getHmallSectIdToReg = buf
	End Function


	Function fngetOptionEditParam(iitemid)
		Dim sqlStr, regedOptArr, i, buf, optionArr, j
		Dim optionLimitNo, optsellYn, boolchk, optTypeName, isSingleOption, slashReplace
		boolchk = False
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemid, outmallOptCode, replace(outmallOptName, '&amp;', '&') as outmallOptName, outmallSellyn, outmalllimitno "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_hmall_regedOption] "
		sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"' "
		rsget.Open sqlStr,dbget
		IF not rsget.EOF THEN
			regedOptArr = rsget.getRows()
		END IF
		rsget.close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemid, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, replace(optionTypeName, char(9), '') as optionTypeName, replace(optionname, char(9), '') as optionname, Len(optionTypeName) as typeLength "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"' "
		rsget.Open sqlStr,dbget
		IF not rsget.EOF THEN
			optionArr = rsget.getRows()
		END IF
		rsget.close

		If IsArray(regedOptArr) Then
			If Ubound(regedOptArr,2) = 0 AND regedOptArr(2, 0) = "���Ͽɼ�" Then			'��ϵ� ���� ��ǰ�϶�
				If FLimitYn = "Y" Then
					optionLimitNo = getLimitEa()
				Else
					If FHmallPrice >= 300000 Then
						optionLimitNo = 30
					Else
						optionLimitNo = 999
					End If
				End If

				If optionLimitNo < 1 Then
					optionLimitNo = 1
				End If

				buf = ""
				buf = buf & "{"
				buf = buf & "  ""itemid"": """&iitemid&""","
				buf = buf & "  ""options"": ["
				buf = buf & "    {"
				buf = buf & "      ""uitmcd"": """& regedOptArr(1, 0) &""","
				buf = buf & "      ""maxSellPossQty"": "&optionLimitNo&","
				buf = buf & "      ""sellGbcd"": """&Chkiif(IsSoldOutLimit5Sell = "True", "11", "00")&""""
				buf = buf & "    }"
				buf = buf & "  ]"
				buf = buf & "}"
			Else
				If IsArray(optionArr) Then
	'				If optionArr(6, 0) = Split(regedOptArr(2, 0), "/")(0) Then
					If optionArr(6, 0) = LEFT(Trim(regedOptArr(2, 0)), Trim(optionArr(8, 0))) Then
						isSingleOption = "Y"
					End If
				End If

				buf = ""
				buf = buf & "{"
				buf = buf & "  ""itemid"": """&iitemid&""","
				buf = buf & "  ""options"": ["
				For i = 0 To Ubound(regedOptArr, 2)
					buf = buf & "    {"
					buf = buf & "      ""uitmcd"": """& regedOptArr(1, i) &""","
					If IsArray(optionArr) Then
						For j = 0 To Ubound(optionArr, 2)
							If isSingleOption = "Y" Then
								slashReplace = replace(Trim(regedOptArr(2, i)), Trim(optionArr(6, 0)) & "/", "")
								slashReplace = replace(slashReplace, "��", "~")
								slashReplace = replace(slashReplace, "��", "&")

								If Trim(slashReplace) = Trim(optionArr(7, j)) Then
									If FLimitYn = "Y" Then
										optionLimitNo = getOptionLimitEa(optionArr(4, j), optionArr(5, j))
									Else
										optionLimitNo = 999
									End If

									If (optionArr(1, j) <> "Y") OR (optionArr(2, j) <> "Y") Then
										optsellYn = "11"
									Else
										optsellYn = "00"
									End If

									If optionLimitNo < 1 Then
										optsellYn = "11"
										optionLimitNo = 1			'�Ǹ� �� ���̶� ��� 1�� �Ǿ� ������ �� �� �Ѥ�;;
									End If

									boolchk = true
									Exit For
								End If
							Else
								slashReplace = replace(regedOptArr(2, i), "/", ",")
								slashReplace = replace(slashReplace, "��", "~")

								If Trim(slashReplace) = Trim(replace(optionArr(7, j), "/", ",")) Then
									If FLimitYn = "Y" Then
										optionLimitNo = getOptionLimitEa(optionArr(4, j), optionArr(5, j))
									Else
										If FHmallPrice >= 300000 Then
											optionLimitNo = 30
										Else
											optionLimitNo = 999
										End If
									End If

									If (optionArr(1, j) <> "Y") OR (optionArr(2, j) <> "Y")  Then
										optsellYn = "11"
									Else
										optsellYn = "00"
									End If

									If optionLimitNo < 1 Then
										optsellYn = "11"
										optionLimitNo = 1			'�Ǹ� �� ���̶� ��� 1�� �Ǿ� ������ �� �� �Ѥ�;;
									End If

									boolchk = true
									Exit For
								End If
							End If
						Next
					End If

					If boolchk = false Then
						optionLimitNo = 1
						optsellYn = "11"
					End If
					buf = buf & "      ""maxSellPossQty"": "&optionLimitNo&","
					buf = buf & "      ""sellGbcd"": """&optsellYn&""""
					buf = buf & "    }"&Chkiif(i = Ubound(regedOptArr, 2), "", ",")  &"  "
				Next
				buf = buf & "  ]"
				buf = buf & "}"
			End If
		End If
		fngetOptionEditParam = buf
	End Function

	Public Function getHmallItemConfirmParameter
		Dim strRst, tt
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>�ٹ�����</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "	<Dataset id=""dsCond"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"
		strRst = strRst & "				<itemCsfDCd />"
		strRst = strRst & "				<venCd>20</venCd>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "</Root>"
		getHmallItemConfirmParameter = strRst
	End Function

	Public Function gethmallItemRegParameter
		Dim strRst, childItemYn
		'################################ �������� �׸� ���� ȣ�� ###############################
		Dim CallSafe, CSafeyn, CSafeGbCd, gbnflag
		CallSafe = getSafetyParam()
		CSafeyn = Split(CallSafe, "|_|")(0)
		CSafeGbCd = Split(CallSafe, "|_|")(1)
		gbnflag = Split(CallSafe, "|_|")(2)
		If CSafeGbCd = "N" Then CSafeGbCd = "" End If

		If FitemLCsfCd = "R6" and CSafeGbCd <> "" Then
			childItemYn = "Y"
		ElseIf gbnflag = "child" Then
			childItemYn = "Y"
		Else
			childItemYn = "N"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"		''������ UTF-8�� �ϱ�@!@@@@@@@@@@@@@@
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>�ٹ�����</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItemCsfCd"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#��ǰ��з��ڵ�
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#��ǰ�ߺз��ڵ�
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#��ǰ�Һз��ڵ�
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#��ǰ���з��ڵ�
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#��ǰ�з������ڵ� | 40 ����Hmall
		strRst = strRst & "				<qaTrgtYn>N</qaTrgtYn>"										'#QA��󿩺� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<safeCertTrgtYn>N</safeCertTrgtYn>"							'#����������󿩺� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<coreMngYn>N</coreMngYn>"									'#�ٽɰ������� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<itstDlbrYn>N</itstDlbrYn>"									'#��ǰ��������ǿ��� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<frdlvSellLimtYn>Y</frdlvSellLimtYn>"						'#�ؿܹ���Ǹ����ѿ��� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItem"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<slitmCd></slitmCd>"										'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "				<slitmNm><![CDATA["&getItemNameFormat()&"]]></slitmNm>"		'#�ǸŻ�ǰ��
		strRst = strRst & "				<sitmCd></sitmCd>"											'��ǰ�ڵ�
		strRst = strRst & "				<bsitmCd></bsitmCd>"										'���ػ�ǰ�ڵ�
		strRst = strRst & "				<bsitmNm></bsitmNm>"										'���ػ�ǰ��
'		strRst = strRst & "				<baseCmpsItemNm></baseCmpsItemNm>"							'�⺻������ǰ��
'		strRst = strRst & "				<addCmpsItemNm></addCmpsItemNm>"							'�߰�������ǰ��
'		strRst = strRst & "				<engItemNm></engItemNm>"									'������ǰ��
'		strRst = strRst & "				<itemUrl></itemUrl>"										'��ǰURL
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#��ǰ��з��ڵ�
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#��ǰ�ߺз��ڵ�
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#��ǰ�Һз��ڵ�
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#��ǰ���з��ڵ�
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#��ǰ�з������ڵ� | 40 ����Hmall
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'���»��ǰ�ڵ�
		strRst = strRst & "				<afcrItemCd></afcrItemCd>"									'���޻��ǰ�ڵ�
		strRst = strRst & "				<frgnDrctBuyYn>N</frgnDrctBuyYn>"							'�ؿ��������� | N/Y
		strRst = strRst & "				<mkcoGbcd>30</mkcoGbcd>"									'#�����籸���ڵ� | 10 ������, 20 ���Կ�, 30 �Ǹſ�
		strRst = strRst & "				<mkcoCd>2347</mkcoCd>"										'#�������ڵ�
		strRst = strRst & "				<mkcoNm>�ٹ�����</mkcoNm>"									'�������
		strRst = strRst & "				<octyCnryGbcd>"&FoctyCnryGbcd&"</octyCnryGbcd>"				'#���������������ڵ�
		strRst = strRst & "				<octyCnryNm>"&FoctyCnryNm&"</octyCnryNm>"					'#������������
		strRst = strRst & "				<prmrOrgCnryGbcd></prmrOrgCnryGbcd>"						'�ֿ���ᱹ�������ڵ�
		strRst = strRst & "				<prmrOrgCnryNm></prmrOrgCnryNm>"							'�ֿ���ᱹ����
		strRst = strRst & "				<prmrOrgNm></prmrOrgNm>"									'�ֿ�����	| �ֿ��ᱹ�������ڵ�/�� ���� ��ϺҰ�
		strRst = strRst & "				<itemMemo></itemMemo>"										'��ǰ�޸�
		strRst = strRst & "				<asGdCntn></asGdCntn>"										'AS�ȳ�����	| ��������� ASó������ 
		strRst = strRst & "				<itemWgt></itemWgt>"										'��ǰ�߷� | frdlvYn �Ӽ��� Y(�ؿܹ��)�� ����ϴ� ��� �ʼ� ���� 30kg �̻��� �ؿܹ�� ���� ��ǰ���� ����"
		strRst = strRst & "				<itemWdthLen></itemWdthLen>"								'��ǰ���α��� | frdlvYn �Ӽ��� Y(�ؿܹ��)�� ����ϴ� ��� �ʼ� �ؿܹ���� ��� �Ҽ��� ���ڸ����� ���"
		strRst = strRst & "				<itemHghLen></itemHghLen>"									'��ǰ���̱��� | frdlvYn �Ӽ��� Y(�ؿܹ��)�� ����ϴ� ��� �ʼ� �ؿܹ���� ��� �Ҽ��� ���ڸ����� ���"
		strRst = strRst & "				<itemHghtLen></itemHghtLen>"								'��ǰ���α��� | frdlvYn �Ӽ��� Y(�ؿܹ��)�� ����ϴ� ��� �ʼ� �ؿܹ���� ��� �Ҽ��� ���ڸ����� ���"
		strRst = strRst & "				<itemGbcd>00</itemGbcd>"									'#��ǰ�����ڵ� | 00 �Ϲݻ�ǰ ,01 ������ǰ ,02 �������Ż�ǰ ,03 ������ǥ��ǰ ,04 ������ǰ ,05 PPL��ǰ ,07 ����������ǰ ,09 ����ǰ ,10 TREND-H��ǰ ,11 Ŭ��H��ǰ"
		strRst = strRst & "				<itemSellGbcd>00</itemSellGbcd>"							'#��ǰ�Ǹű����ڵ� | 00 ����, 11 �Ͻ��ߴ�, 19 �����ߴ�
		strRst = strRst & "				<adltItemYn>"&CHKIIF(IsAdultItem= "Y", "Y", "N")&"</adltItemYn>"	'#���ο�ǰ���� | Y or N
		strRst = strRst & "				<itemRegTcndAgrYn>Y</itemRegTcndAgrYn>"						'#��ǰ��Ͼ�����ǿ��� | Y or N -> �ݵ�� Y
		strRst = strRst & "				<jwlSvrtEnclYn>N</jwlSvrtEnclYn>"							'#������������������ | Y or N (���������������� Y)
		strRst = strRst & "				<giftItemYn>N</giftItemYn>"									'#����ǰ��ǰ���� | 'N'������ ����.
		strRst = strRst & "				<tcommUseYn>N</tcommUseYn>"									'#TĿ�ӽ���뿩�� | Y or N
		strRst = strRst & "				<stckGdYn>N</stckGdYn>"										'#���ȳ����� | Y or N
		strRst = strRst & "				<hmallRsvSellYn>N</hmallRsvSellYn>"							'#HMALL�����Ǹſ��� | Y or N
		strRst = strRst & "				<frgnBuyPrxyYn>N</frgnBuyPrxyYn>"							'#�ؿܱ��Ŵ��࿩�� | Y or N
		strRst = strRst & "				<basktUseNdmtYn>N</basktUseNdmtYn>"							'#��ٱ��ϻ��Ұ����� | Y or N
		strRst = strRst & "				<prsnMsgPossYn>N</prsnMsgPossYn>"							'#�����޽������ɿ��� | Y or N
		strRst = strRst & "				<prsnPackPossYn>N</prsnPackPossYn>"							'#�������尡�ɿ��� | Y or N
		strRst = strRst & "				<addBuyOptUseYn>Y</addBuyOptUseYn>"							'#�߰����ſɼǻ�뿩�� | Y or N
		strRst = strRst & "				<oshpVenAdrSeq>4</oshpVenAdrSeq>"							'#������»��ּҼ��� | Y or N
		strRst = strRst & "				<rtpExchVenAdrSeq>4</rtpExchVenAdrSeq>"						'#��ǰ��ȯ���»��ּҼ��� | Y or N
		strRst = strRst & "				<emgyExchVenAdrSeq></emgyExchVenAdrSeq>"					'��ޱ�ȯ���»��ּҼ���
		strRst = strRst & "				<itntDispYn>Y</itntDispYn>"									'#���ͳ����ÿ���
		strRst = strRst & "				<itemQnaExpsYn>Y</itemQnaExpsYn>"							'#��ǰQNA���⿩��
		strRst = strRst & "				<webExpsPrmoNm><![CDATA[]]></webExpsPrmoNm>"				'���������θ�Ǹ�
'		strRst = strRst & "				<prmo2TxtCntn></prmo2TxtCntn>"								'���θ��2 �������� | ���θ��2 �������� �Է½� ���θ�ǳ��� ��������/�Ͻ� �� ���θ�ǳ��� ��������/�Ͻ�(prmoExpsStrtDtm, prmoExpsStrtTime, prmoExpsEndDtm, prmoExpsEndTime) �ʼ��Է�"
'		strRst = strRst & "				<prmoExpsStrtDtm></prmoExpsStrtDtm>"						'���θ�ǳ��� �������� | ex) 20191118 �Է� prmo2TxtCntn �Է� �� �ʼ�(Y)"
'		strRst = strRst & "				<prmoExpsStrtTime></prmoExpsStrtTime>"						'���θ�ǳ��� �����Ͻ� | 0~23�� �ð��������� �Է� (15�ú����� ��� ���θ�� ������ ��� 15 �Է�) prmo2TxtCntn �Է� �� �ʼ�(Y)"
'		strRst = strRst & "				<prmoExpsEndDtm></prmoExpsEndDtm>"							'���θ�ǳ��� �������� | ex) 20191118 �Է� prmo2TxtCntn �Է� �� �ʼ�(Y)"
'		strRst = strRst & "				<prmoExpsEndTime></prmoExpsEndTime>"						'���θ�ǳ��� �����Ͻ� | 0~24�� �ð��������� �Է� (0~23���ϰ�� 23:00:00 �� ����, 24���� ��� 23:59:59�� �ڵ� ���õ�) prmo2TxtCntn �Է� �� �ʼ�(Y)"
		strRst = strRst & "				<prmoTxtDcCopnYn>N</prmoTxtDcCopnYn>"						'#���θ�ǹ��������������� | Y or N
		strRst = strRst & "				<prmoTxtSpdcYn>N</prmoTxtSpdcYn>"							'#���θ�ǹ�����¦���ο��� | Y or N
		strRst = strRst & "				<prmoTxtSvmtYn>N</prmoTxtSvmtYn>"							'#���θ�ǹ��������ݿ��� | Y or N
		strRst = strRst & "				<prmoTxtFamtFxrtGbcd>2</prmoTxtFamtFxrtGbcd>"				'���θ�ǹ����������������ڵ� | "1 ����, 2 ����"
		strRst = strRst & "				<prmoTxtSvmtPrdcYn>N</prmoTxtSvmtPrdcYn>"					'#���θ�ǹ��������ݼ����ο��� | default 'N'������ ����.
		strRst = strRst & "				<prmoTxtWintYn>N</prmoTxtWintYn>"							'#���θ�ǹ��������ڿ��� | default 'N'������ ����.
		strRst = strRst & "				<prmoTxtSpymDcYn>N</prmoTxtSpymDcYn>"						'#���θ�ǹ����Ͻú����ο��� | default 'N'������ ����.
		strRst = strRst & "				<frdlvYn>N</frdlvYn>"										'#�ؿܹ�ۿ��� | default 'N'������ ����.
		strRst = strRst & "				<packOpenRtpNdmtYn>Y</packOpenRtpNdmtYn>"					'#������¹�ǰ�Ұ����� | Y or N
		strRst = strRst & "				<ostkYn>N</ostkYn>"											'#ǰ������ | default 'N'������ ����.
		strRst = strRst & "				<dlvHopeDtDsntYn>N</dlvHopeDtDsntYn>"						'#������������������ | default 'N'������ ����.
		strRst = strRst & "				<itstHtmlYn>Y</itstHtmlYn>"									'#��ǰ�����HTML���� | Y or N
		strRst = strRst & "				<itstPhotoExpsYn>Y</itstPhotoExpsYn>"						'#��ǰ������������⿩�� | default 'N'������ ����.
		strRst = strRst & "				<dlvItemFormGbcd></dlvItemFormGbcd>"						'��ۻ�ǰ���±����ڵ�	String(2)	"00 �Ϲ�, 10 ���, 20 �õ�
		strRst = strRst & "				<qckDlvPossYn>N</qckDlvPossYn>"								'#����۰��ɿ��� | default 'N'������ ����.
		strRst = strRst & "				<dwtdYn>N</dwtdYn>"											'#��ȸ������ | default 'N'������ ����.
		strRst = strRst & "				<lrpyYn>Y</lrpyYn>"											'#��ȯ�ҿ��� | "Y or N (�ؿܹ���� ��� 'Y'�� �ԷµǾ�� ��) �ؿܹ�� ��ǰ�� ��ȯ�� ��ǰ"
		strRst = strRst & "				<sameItemMxpkPossQty></sameItemMxpkPossQty>"				'���ϻ�ǰ�����尡�ɼ���	Number	
		strRst = strRst & "				<mxpkYn>N</mxpkYn>"											'#�����忩�� | default 'N'������ ����.
		strRst = strRst & "				<packMagnGbcd>20</packMagnGbcd>"							'#������ü�����ڵ�	String(2)	"10 ��� 20 ���»�" 
		strRst = strRst & "				<dlvcGbcd>00</dlvcGbcd>"									'#��ۺ񱸺��ڵ�	String(2)	'00' �Ϲݻ�ǰ ���� �Է�
		strRst = strRst & "				<dlvcPayGbcd>10</dlvcPayGbcd>"								'#��ۺ����ұ����ڵ�	String(2)	"00 ����, 10 ������, 20 ����, 30 ��ġ��ǰ (�ԷºҰ�)"
		strRst = strRst & "				<arpayDlvGdCntn></arpayDlvGdCntn>"							'���ҹ�۾ȳ�����	String(400)	
		strRst = strRst & "				<cvstWtdwPossYn>N</cvstWtdwPossYn>"							'#������ȸ������ | default 'N'������ ����.
		strRst = strRst & "				<prpyDlvCost></prpyDlvCost>"								'���޹�ۺ�� | 00(����)��  ���Է�(null) 10(������)��  ���� �����ϸ� ��ǰ����ۺ��� null �̸� ������ۺ�(�Ҿ���ٱ���) 20(����)��  ���Է�(null) 30(��ġ��ǰ)��  ���Է�(null) ��ȭ�����»�&��ȭ������� ��� ��ǰ����ۺ� ��� �Ұ�
		strRst = strRst & "				<irgnAreaAddDlvCost></irgnAreaAddDlvCost>"					'���������߰���ۺ��
		strRst = strRst & "				<mngWhNo>990</mngWhNo>"										'#����â���ȣ
		strRst = strRst & "				<sbctDlvcoCd>12</sbctDlvcoCd>"								'#���޹�ۻ��ڵ�
		strRst = strRst & "				<dlvMagnGbcd>20</dlvMagnGbcd>"								'#�����ü�����ڵ� | 10 Ȩ����,20 ���»�,30 �ù��"
		strRst = strRst & "				<dlvcChmgGbcd>20</dlvcChmgGbcd>"							'#��ۺ�δ���ü�����ڵ� | 10 Ȩ����,20 ���»�,30 �ù��NL NULL"
		strRst = strRst & "				<dlvFormGbcd>40</dlvFormGbcd>"								'#������±����ڵ� | 00 ���͹��,10 ��ȭ�����,20 ����Ȩ���ù�,30 ���»����ù�,40 ���»�����,50 ��ȭ���������"
		strRst = strRst & "				<rtpWdmtGbcd>2</rtpWdmtGbcd>"								'��ǰȸ����������ڵ� | 1 �������ݼ�,2 ���»�ȸ��"
		strRst = strRst & "				<rtpDlvCost>6000</rtpDlvCost>"								'��ǰ��ۺ��
		strRst = strRst & "				<exchWdmtGbcd>2</exchWdmtGbcd>"								'��ȯȸ����������ڵ� | 1 �������ݼ�,2 ���»�ȸ��"
		strRst = strRst & "				<exchDlvCost>6000</exchDlvCost>"							'��ȯ��ۺ��
		strRst = strRst & "				<custDlvcWdmtGbcd>4</custDlvcWdmtGbcd>"						'����ۺ�ȸ����������ڵ� | 1 ��۹ڽ�����,2 �������Ա�,3 �ù�������,4 ��ǰ ������ ������"
		strRst = strRst & "				<stlmWayScopGbcd>10</stlmWayScopGbcd>"						'�������ܹ��������ڵ� | 10 �� �������ܰ���,20 ����/����ī�常��밡��,30 ��ǰ������ ��� ���� ����"
		strRst = strRst & "				<pntStlmNdmtYn>N</pntStlmNdmtYn>"							'#����Ʈ�����Ұ�����
		strRst = strRst & "				<ostkRishpSmsYn>N</ostkRishpSmsYn>"							'#ǰ�����԰�SMS���� | default 'N'������ ����.
		strRst = strRst & "				<oshpSmsExcldYn>N</oshpSmsExcldYn>"							'#���SMS���ܿ��� | Y or N
		strRst = strRst & "				<hmallItemSrchExcldYn>N</hmallItemSrchExcldYn>"				'HMALL��ǰ�˻����ܿ��� | Y or N
		strRst = strRst & "				<itemTypeGbcd>01</itemTypeGbcd>"							'#��ǰ���������ڵ� | 00 ��ȭ����ǰ, 01 ���»��ǰ, 02 ��ǰ��, 03 ��Ź��ǰ��, 04 ����/����(��ȭ��), 05 ����/����(���»�) ������ ���� ��ǰ ��Ͻ� ��ȭ�� ���»��� ��� 04, �Ϲ����»��� ��� 05 ���
'		strRst = strRst & "				<intgItemGbcd></intgItemGbcd>"								'������ǰ�����ڵ� | 01 ��Ż, 02 ����, 03 ����, 04 �޴���, 05 ����Һ�, 06 ����� ��ǰ��, 07 �ʰ� ��ǰ, 08 ����(Hmall)  ��ǰ�����ڵ�(itemGbcd)���� ������ǰ(04)�� ��쿡�� �ʼ� �Է�
'		strRst = strRst & "				<intgItemStlmYn>N</intgItemStlmYn>"							'#������ǰ����
		strRst = strRst & "				<prchMdaGbcd>40</prchMdaGbcd>"								'#���Ը�ü�����ڵ� | 40 Hmall  20 Ȩ����
		strRst = strRst & "				<frstRegMdaGbcd>02</frstRegMdaGbcd>"						'#���ʵ�ϸ�ü�����ڵ� | 02 ���ͳ�
		strRst = strRst & "				<vatRate><![CDATA[10]]></vatRate>"							'#�ΰ�������
		strRst = strRst & "				<itemTaxnYn>"&Chkiif(Fvatinclude="Y", "Y", "N")&"</itemTaxnYn>"	'#��ǰ�������� | Y or N
		strRst = strRst & "				<venCd>002569</venCd>"										'#���»��ڵ�
		strRst = strRst & "				<ven2Cd></ven2Cd>"											'2�����»��ڵ�
		strRst = strRst & "				<prchMthdGbcd>33</prchMthdGbcd>"							'#���Թ�������ڵ� | 11 ����, 22 Ư��, 33 ������
		strRst = strRst & "				<itemTaxnGbcd>"&Chkiif(Fvatinclude="Y", "001", "000")&"</itemTaxnGbcd>"		'#��ǰ���������ڵ� | 000 �鼼, 001 ����, 002 ����
		strRst = strRst & "				<ringItemYn>N</ringItemYn>"									'#������ǰ���� | default 'N'������ ����.
		strRst = strRst & "				<brndGbcd>40</brndGbcd>"									'#�귣�屸���ڵ� | 40 Hmall�귣��
		strRst = strRst & "				<brndCd>205390</brndCd>"									'#�귣���ڵ� | ��ϵ� �귣�� ��ü ��ȸ(selectBrndList) ���� ��� �ش� MD���� �ű� ��� ��û
		strRst = strRst & "				<itntBrndNm>�ٹ�����</itntBrndNm>"							'#���ͳݺ귣���
		strRst = strRst & "				<dptsPchCd></dptsPchCd>"									'��ȭ����Ī�ڵ� | ��ȭ�����»縸 �Է�
		strRst = strRst & "				<rsptMdCd>8048</rsptMdCd>"									'#�����MD�ڵ�
		strRst = strRst & getAttrInfo()
		strRst = strRst & "				<ordMakeYn>"&FordMakeYn&"</ordMakeYn>"						'#�ֹ����ۿ��� | Y or N (�ֹ����۽� Y)
'		strRst = strRst & "				<baseSectId></baseSectId>"									'#�⺻����ID | �Ǹŵ� �⺻������
'		strRst = strRst & "				<frdlvFormGbcd></frdlvFormGbcd>"							'�ؿܹ�����±����ڵ�  | 1 �ڽ�, 2 ���� (�ؿܹ���� ��� �ʼ� ��)-> �ؿ� ��ۿ��� ���� ���ý� ���� �ִ� ������� ���� 52cm x ���� 44cm�� �ʰ��� �� ����)"
'		strRst = strRst & "				<hscd></hscd>"												'HS�ڵ� | �ؿܹ���� ��� �ʼ� HS�ڵ� �� ��ȿ�� üũ(Harmonized System)"
'		strRst = strRst & "				<frgnOrdPiupSrvYn></frgnOrdPiupSrvYn>"						'�ؿ��ֹ��Ⱦ����񽺿��� | default 'N'������ ����.		
'		strRst = strRst & "				<frdlvNchgYn></frdlvNchgYn>"								'�ؿܹ�۹��Ῡ�� | default 'N'������ ����.
		strRst = strRst & "				<childItemYn>"&childItemYn&"</childItemYn>"					'#��̻�ǰ���� | "Y 13���̸� ��̻�ǰ N 13���̻� ��ǰ"
'		strRst = strRst & "				<hdmalItnlYn></hdmalItnlYn>"								'������� ���� ���� | ��ȭ�� ��ǰ�� ��� �������� Y�� ���� ����
'		strRst = strRst & "				<chkExceptSafeCert></chkExceptSafeCert>"					'����������Ͽ���ó�� | ����������ü ���� ó�� ��� ���»��� ��� ����������� ���߼Ҽ��� ��ǰ ��Ͻ� Y �����ϸ� �������� �ʼ��Է����� �ʾƵ� ��(db���尪 �ƴ�)��� HELP �������μ��� ����
		strRst = strRst & "				<inslItemYn>N</inslItemYn>"									'��ġ��ǰ���� | Y ��ġ��ǰ, N ��ġ��ǰ�ƴ�
'		strRst = strRst & "				<meatHisYn></meatHisYn>"									'�����̷�ǥ�ÿ��� | Y ǥ��, N ��ǥ��
		strRst = strRst & "				<safeMngTrgtYn>"&Chkiif(gbnflag="elec", "Y", "N")&"</safeMngTrgtYn>"					'���ȹ���󿩺� | Y ���ȹ� ��� (���ȹ� ��� ��ǰ�� ��� safeCertTypeGbcd �� �ʼ�), N ���ȹ� ����
		strRst = strRst & "				<chemSafeTrgtYn>"&Chkiif(gbnflag="life", "Y", "N")&"</chemSafeTrgtYn>"							'��Ȱȭ����ǰ ��󿩺� | Y : ��Ȱȭ����ǰ ���, N : ��Ȱȭ����ǰ ����
		strRst = strRst & "				<parlImprYn>N</parlImprYn>"									'������Կ��� | Y ������Ի�ǰ, N ������Ի�ǰ �ƴ�
		strRst = strRst & "				<itemYetaGbcd>00</itemYetaGbcd>"							'#��ǰ�������걸���ڵ� | 00 �Ϲ�, 01 ����, 02 ����, itemTypeGbcd�� 04 Ȥ�� 05�� ������ ��� ������ ��� 01, ������ǰ�� ��� 02 �����ؾ� �� �Ϲݻ�ǰ�� ��� 00���� ó�� �ʿ�
'		strRst = strRst & "				<dawnDlvYn></dawnDlvYn>"									'������ۿ��� | Y or N
'		strRst = strRst & "				<stpicPossYn></stpicPossYn>"								'������Ȱ��ɿ��� | Y or N
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"							'�����Ⱦ����ɿ��� | Y or N
		strRst = strRst & "				<areaDlvCostAddYn>Y</areaDlvCostAddYn>"						'������ۺ���߰����� | Y or N
		strRst = strRst & "				<jejuAddDlvCost>3000</jejuAddDlvCost>"						'���ֵ��߰���ۺ�� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddDlvCost>3000</irgnAddDlvCost>"						'�����߰���ۺ�� | õ���̻� �������� ����
		strRst = strRst & "				<areaRtpCostAddYn>Y</areaRtpCostAddYn>"						'������ǰ����߰����� | Y or N
		strRst = strRst & "				<jejuAddRtpCost>3000</jejuAddRtpCost>"						'���ֵ��߰���ǰ��� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddRtpCost>3000</irgnAddRtpCost>"						'�����߰���ǰ��� | õ���̻� �������� ����
		strRst = strRst & "				<areaExchCostAddYn>Y</areaExchCostAddYn>"					'������ȯ����߰����� | Y or N
		strRst = strRst & "				<jejuAddExchCost>3000</jejuAddExchCost>"					'���ֵ��߰���ȯ��� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddExchCost>3000</irgnAddExchCost>"					'�����߰���ȯ��� | õ���̻� �������� ����
'		strRst = strRst & "				<harmItemYn></harmItemYn>"									'���ػ�ǰ���� | Y or N
'		strRst = strRst & "				<brodEqmtMngTrgtYn></brodEqmtMngTrgtYn>"					'������������󿩺� | Y or N
		strRst = strRst & "				<bbprcCopnDcYn>Y</bbprcCopnDcYn>"							'���ø�����ǥ�� �������� | Y or N
		strRst = strRst & "				<bbprcSpymDcYn>Y</bbprcSpymDcYn>"							'���ø�����ǥ�� �Ͻú����� | Y or N
		strRst = strRst & "				<bbprcSvmtPrdcYn>Y</bbprcSvmtPrdcYn>"						'���ø�����ǥ�� H.Point������ | Y or N
		strRst = strRst & "				<bbprcSpdcYn>Y</bbprcSpdcYn>"								'���ø�����ǥ�� ��¦���� | Y or N
		strRst = strRst & "				<prcExpsBitVal1>1</prcExpsBitVal1>"							'���ݺ񱳳��Ⱑ �������� | 0: �ش����, 1: �ش� �� ���ݺ񱳳��Ⱑ ���� �׸���� ��Ʈ������ ���� �� �Է��ϴ��� �ƴϸ� �ƿ� ��ΰ� �Է��� ���ϴ��� �ؾ� ��. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal2>2</prcExpsBitVal2>"							'���ݺ񱳳��Ⱑ �Ͻú����� | 0: �ش����, 2: �ش� �� ���ݺ񱳳��Ⱑ ���� �׸���� ��Ʈ������ ���� �� �Է��ϴ��� �ƴϸ� �ƿ� ��ΰ� �Է��� ���ϴ��� �ؾ� ��. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal4>4</prcExpsBitVal4>"							'���ݺ񱳳��Ⱑ H.Point������ | 0: �ش����, 4: �ش� �� ���ݺ񱳳��Ⱑ ���� �׸���� ��Ʈ������ ���� �� �Է��ϴ��� �ƴϸ� �ƿ� ��ΰ� �Է��� ���ϴ��� �ؾ� ��. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal8>8</prcExpsBitVal8>"							'���ݺ񱳳��Ⱑ ��¦���� | 0: �ش����, 8: �ش� �� ���ݺ񱳳��Ⱑ ���� �׸���� ��Ʈ������ ���� �� �Է��ϴ��� �ƴϸ� �ƿ� ��ΰ� �Է��� ���ϴ��� �ؾ� ��. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
'		strRst = strRst & "				<frgnDrctDlvYn></frgnDrctDlvYn>"							'�ؿ������(�����ȣ���ʿ�) | Y or N
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallSectIdToReg("I")
		If CSafeyn = "Y" Then
			strRst = strRst & getHmallItemSafeInfoToReg(CSafeGbCd, "I")
		End If
'		strRst = strRst & "	<Dataset id=""dsSlitmBcdDtl"">"											'���ڵ峻��
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<bcdBrndGbcd></bcdBrndGbcd>"								'���ڵ�귣�屸���ڵ� | 10 �Ϲݺ귣��, 20 ��ȭ���귣��
'		strRst = strRst & "				<shrtBcdVal></shrtBcdVal>"									'������ڵ尪
'		strRst = strRst & "				<totBcdVal></totBcdVal>"									'��ü���ڵ尪
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsHmallRsvItemDtl"">"										'Hmall�����Ǹ�����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<sellStrtDt></sellStrtDt>"									'#�ǸŽ�������
'		strRst = strRst & "				<sellEndDt></sellEndDt>"									'#�Ǹ���������
'		strRst = strRst & "				<dlvStrtDt></dlvStrtDt>"									'#��۽�������
'		strRst = strRst & "				<dlvEndDt></dlvEndDt>"										'#�����������
'		strRst = strRst & "				<dlvAdmGdCntn></dlvAdmGdCntn>"								'#��۰����ھȳ�����
'		strRst = strRst & "				<custGdCntn></custGdCntn>"									'#���ȳ�����
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsSlitmAsVenHis"">"										'AS���»�����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<asVenNm></asVenNm>"										'#AS���»��
'		strRst = strRst & "				<asgnrNm></asgnrNm>"										'#����ڸ�
'		strRst = strRst & "				<rgno></rgno>"												'#����ڵ�Ϲ�ȣ
'		strRst = strRst & "				<postNo></postNo>"											'#�����ȣ
'		strRst = strRst & "				<venBaseAdr></venBaseAdr>"									'#���»�⺻�ּ�
'		strRst = strRst & "				<venPtcAdr></venPtcAdr>"									'#���»���ּ�
'		strRst = strRst & "				<tela></tela>"												'#��ȭ������ȣ
'		strRst = strRst & "				<tels></tels>"												'#��ȭ����ȣ
'		strRst = strRst & "				<teli></teli>"												'#��ȭ������ȣ
'		strRst = strRst & "				<extsTel></extsTel>"										'#������ȭ��ȣ
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsItemDlvNdmtDtl"">"										'��ۺҰ�����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<apocGbcd></apocGbcd>"										'#���������ڵ� | 10 ����,������, 11 ����Ź�ۺҰ�����, 12 ����,������ ���� ��ۺҰ�����, 20 ����(��ȭ��������ۺҰ�) ���Ϲݾ�ü���úҰ�!, 21 ����Ұ�, 22 ����/�갣���� �Ұ�, 30 ���ֺҰ�
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsSlitmPrcAthzHis"">"										'#��������
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<prcAplyStrtDtm>"&Replace(Date(), "-", "")&"</prcAplyStrtDtm>"	'#������������Ͻ�
		strRst = strRst & "				<prcAthzGbcd>00</prcAthzGbcd>"								'#���ݰ��籸���ڵ� | 00 MD���δ��
		strRst = strRst & "				<sellPrc>"&MustPrice()&"</sellPrc>"							'#�ǸŰ��� | ���ǸŰ� , VAT ���� �������� �ʽ��ϴ�. ���õ� ���� �״�� �Ǹŵ˴ϴ�.
		strRst = strRst & "				<mrgnRate>"&FMrgnRate&"</mrgnRate>"							'#�������� | ������ �Է�,  '%'�� �Է����� �����ּ���.
		strRst = strRst & "				<dptsOpCd></dptsOpCd>"										'#��ȭ��OP�ڵ� | OP�ڵ�
		strRst = strRst & "				<dptsVenOpCd></dptsVenOpCd>"								'��ȭ�����»�OP�ڵ� | ��ȭ�����»��ΰ�� �ʼ��� �Է�
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'���»��ǰ�ڵ� | ���»��ǰ�ڵ� ����� �ʼ� �Է�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsHpItemDtl"">"											'�ڵ�����ǰ����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<trfsCntn></trfsCntn>"										'#���������
'		strRst = strRst & "				<stplMths></stplMths>"										'#����������
'		strRst = strRst & "				<ccrgAmt></ccrgAmt>"										'#����ݱݾ�
'		strRst = strRst & "				<teRealChrgAmt</teRealChrgAmt>"								'#�ܸ���Ǻδ�ݾ�
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getOptSellUitmDtl()
		strRst = strRst & getOptTypeMst("I")
		strRst = strRst & getOptAttrMst("I")
		strRst = strRst & getOptCombDtl("I")

'		strRst = strRst & "	<Dataset id=""dsAsctSlitmDtl"">"										'���û�ǰ����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<asctItemGbcd></asctItemGbcd>"								'#���û�ǰ�����ڵ� | 10 Ʈ����H
'		strRst = strRst & "				<asctSlitmCd></asctSlitmCd>"								'#�����ǸŻ�ǰ�ڵ�
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallItemInfoCdToReg("I")

		strRst = strRst & "	<Dataset id=""dsHtmlItstMst"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<htmlItstGbcd>00</htmlItstGbcd>"							'HTML��ǰ����������ڵ� | 00 �Ϲ�, 01 ��ǰ, 02 ��������
		strRst = strRst & "				<htmlItstCntn><![CDATA["&getHmallContParamToReg()&"]]></htmlItstCntn>"	'��ǰ����������ڵ� | HTML��ǰ���������
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsMdaSlitmDtl"">"											'��ü�Ǹų���
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"							'�ǸŸ�ü�з��ڵ� | 02 : Hmall, 04 : ����� �ΰ� ��� üũ �ʿ�
		strRst = strRst & "			</row>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>04</sellMdaCsfCd>"							'�ǸŸ�ü�з��ڵ� | 02 : Hmall, 04 : ����� �ΰ� ��� üũ �ʿ�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsCsmtMkcoSlitmDtl"">"									'ȭ��ǰ�������ǸŻ�ǰ����
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"							'ȭ��ǰ��������� | ȭ��ǰ ��ǰ�ΰ�� �ʼ��� �Է��Ͽ��� �մϴ�.
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsItemIntlAddSetupDtl"">"									'������� ������������
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<storpiupExclItemYn></storpiupExclItemYn>"			'������������ǰ(�ù�Ұ�) | dsItem.hdmalIntlYn(������� ��������)�� Y�� ��� �ʼ��Է� �Է°� : Y/N (������Ȱ��� �� Y�� ��쿡�� Y ����)"
'		strRst = strRst & "				<storpiupPossYn></storpiupPossYn>"					'������Ȱ��� | dsItem.hdmalIntlYn(������� ��������)�� Y�� ��� �ʼ��Է� �Է°� : Y/N"
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"					'�����Ⱦ����� | dsItem.hdmalIntlYn(������� ��������)�� Y�� ��� �ʼ��Է� �Է°� : Y/N"
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & "</Root>"
		gethmallItemRegParameter = strRst
	End Function

	Public Function gethmallItemEditParameter
		Dim strRst, childItemYn
		'################################ �������� �׸� ���� ȣ�� ###############################
		Dim CallSafe, CSafeyn, CSafeGbCd, gbnflag
		CallSafe = getSafetyParam()
		CSafeyn = Split(CallSafe, "|_|")(0)
		CSafeGbCd = Split(CallSafe, "|_|")(1)
		gbnflag = Split(CallSafe, "|_|")(2)
		If CSafeGbCd = "N" Then CSafeGbCd = "" End If

		If FitemLCsfCd = "R6" and CSafeGbCd <> "" Then
			childItemYn = "Y"
		ElseIf gbnflag = "child" Then
			childItemYn = "Y"
		Else
			childItemYn = "N"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>�ٹ�����</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItemCsfCd"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#��ǰ��з��ڵ�
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#��ǰ�ߺз��ڵ�
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#��ǰ�Һз��ڵ�
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#��ǰ���з��ڵ�
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#��ǰ�з������ڵ� | 40 ����Hmall
		strRst = strRst & "				<qaTrgtYn>N</qaTrgtYn>"										'#QA��󿩺� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<safeCertTrgtYn>N</safeCertTrgtYn>"							'#����������󿩺� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<coreMngYn>N</coreMngYn>"									'#�ٽɰ������� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<itstDlbrYn>N</itstDlbrYn>"									'#��ǰ��������ǿ��� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "				<frdlvSellLimtYn>Y</frdlvSellLimtYn>"						'#�ؿܹ���Ǹ����ѿ��� | ��ǰ�з���ȸ �� �ش� �������Է�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItem"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"						'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "				<slitmNm><![CDATA["&getItemNameFormat()&"]]></slitmNm>"		'#�ǸŻ�ǰ��
		strRst = strRst & "				<venCd>002569</venCd>"										'#���»��ڵ�
'		strRst = strRst & "				<ven2Cd></ven2Cd>"											'2�����»��ڵ�
'		strRst = strRst & "				<addCmpsItemNm></addCmpsItemNm>"							'�߰�������ǰ��
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'���»��ǰ�ڵ�
		If FmaySoldOut = "Y" OR IsMayLimitSoldout = "Y" Then
			strRst = strRst & "			<itemSellGbcd>11</itemSellGbcd>"							'#�ǸŻ��±��� |00 ����, 11 �Ͻ��ߴ�, 19 �����ߴ�..�Ŵ��󿡴� ���� �� �̰� ������ �����ȵ�
		Else
			strRst = strRst & "			<itemSellGbcd>00</itemSellGbcd>"							'#�ǸŻ��±��� |00 ����, 11 �Ͻ��ߴ�, 19 �����ߴ�..�Ŵ��󿡴� ���� �� �̰� ������ �����ȵ�
		End If
'		strRst = strRst & "				<giftEvntStrtDtm></giftEvntStrtDtm>"						'����ǰ�̺�Ʈ�����Ͻ�
'		strRst = strRst & "				<giftEvntEndDtm>S4</giftEvntEndDtm>"						'����ǰ�̺�Ʈ�����Ͻ�
'		strRst = strRst & "				<giftCntn>S4</giftCntn>"									'����ǰ����
'		strRst = strRst & "				<giftImgNm>S4</giftImgNm>"									'����ǰ�̹�����
'		strRst = strRst & "				<tcommUseYn>N</tcommUseYn>"									'TĿ�ӽ���뿩��
'		strRst = strRst & "				<webExpsPrmoNm></webExpsPrmoNm>"							'���������θ�Ǹ�
'		strRst = strRst & "				<prmo2TxtCntn></prmo2TxtCntn>"								'���θ��2 ��������
'		strRst = strRst & "				<prmoExpsStrtDtm></prmoExpsStrtDtm>"						'���θ�ǳ��� ��������
'		strRst = strRst & "				<prmoExpsStrtTime></prmoExpsStrtTime>"						'���θ�ǳ��� �����Ͻ�
'		strRst = strRst & "				<prmoExpsEndDtm></prmoExpsEndDtm>"							'���θ�ǳ��� ��������
'		strRst = strRst & "				<prmoExpsEndTime></prmoExpsEndTime>"						'���θ�ǳ��� �����Ͻ�
		strRst = strRst & "				<itstHtmlYn>Y</itstHtmlYn>"									'��ǰ�����HTML����
		strRst = strRst & "				<itstPhotoExpsYn>Y</itstPhotoExpsYn>"						'��ǰ������������⿩��
		strRst = strRst & getAttrInfo()
		strRst = strRst & "				<childItemYn>"&childItemYn&"</childItemYn>"					'#��̻�ǰ���� | Y : 13���̸� ��̻�ǰ, N : 13���̻� ��ǰ
'		strRst = strRst & "				<childItemYn>N</childItemYn>"								'#��̻�ǰ���� | Y : 13���̸� ��̻�ǰ, N : 13���̻� ��ǰ
'		strRst = strRst & "				<childItemYn>"&maybeChildYn&"</childItemYn>"				'#��̻�ǰ���� | Y : 13���̸� ��̻�ǰ, N : 13���̻� ��ǰ
		strRst = strRst & "				<safeCertTypeGbcd>"&CSafeGbCd&"</safeCertTypeGbcd>"			'�������������ڵ� | 01 ��������, 02 ����Ȯ��, 03 ������ ���ռ� Ȯ��, 04 ���������ؼ����
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#��ǰ��з��ڵ�
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#��ǰ�ߺз��ڵ�
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#��ǰ�Һз��ڵ�
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#��ǰ���з��ڵ�
'		strRst = strRst & "				<frgnDrctBuyYn></frgnDrctBuyYn>"							'�ؿ���������
		strRst = strRst & "				<safeMngTrgtYn>"&Chkiif(gbnflag="elec", "Y", "N")&"</safeMngTrgtYn>"					'���ȹ���󿩺� | Y ���ȹ� ��� (���ȹ� ��� ��ǰ�� ��� safeCertTypeGbcd �� �ʼ�), N ���ȹ� ����
		strRst = strRst & "				<chemSafeTrgtYn>"&Chkiif(gbnflag="life", "Y", "N")&"</chemSafeTrgtYn>"							'��Ȱȭ����ǰ ��󿩺� | Y : ��Ȱȭ����ǰ ���, N : ��Ȱȭ����ǰ ����
		strRst = strRst & "				<parlImprYn>N</parlImprYn>"									'������Կ��� | Y ������Ի�ǰ, N ������Ի�ǰ �ƴ�
'		strRst = strRst & "				<stpicPossYn></stpicPossYn>"								'������Ȱ��ɿ��� | Y or N
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"							'�����Ⱦ����ɿ��� | Y or N
'		strRst = strRst & "				<brodEqmtMngTrgtYn></brodEqmtMngTrgtYn>"					'������������󿩺� | Y or N
		strRst = strRst & "				<dlvFormGbcd>40</dlvFormGbcd>"								'#������±����ڵ� | 00 ���͹�� ,10 ��ȭ����� ,20 ����Ȩ���ù� ,30 ���»����ù� ,40 ���»����� ,50 ��ȭ���������
		strRst = strRst & "				<areaDlvCostAddYn>Y</areaDlvCostAddYn>"						'������ۺ���߰����� | Y or N
		strRst = strRst & "				<jejuAddDlvCost>3000</jejuAddDlvCost>"						'���ֵ��߰���ۺ�� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddDlvCost>3000</irgnAddDlvCost>"						'�����߰���ۺ�� | õ���̻� �������� ����
		strRst = strRst & "				<areaRtpCostAddYn>Y</areaRtpCostAddYn>"						'������ǰ����߰����� | Y or N
		strRst = strRst & "				<jejuAddRtpCost>3000</jejuAddRtpCost>"						'���ֵ��߰���ǰ��� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddRtpCost>3000</irgnAddRtpCost>"						'�����߰���ǰ��� | õ���̻� �������� ����
		strRst = strRst & "				<areaExchCostAddYn>Y</areaExchCostAddYn>"					'������ȯ����߰����� | Y or N
		strRst = strRst & "				<jejuAddExchCost>3000</jejuAddExchCost>"					'���ֵ��߰���ȯ��� | õ���̻� �������� ����
		strRst = strRst & "				<irgnAddExchCost>3000</irgnAddExchCost>"					'�����߰���ȯ��� | õ���̻� �������� ����
		strRst = strRst & "				<bbprcCopnDcYn>Y</bbprcCopnDcYn>"							'���ø�����ǥ�� �������� | Y or N
		strRst = strRst & "				<bbprcSpymDcYn>Y</bbprcSpymDcYn>"							'���ø�����ǥ�� �Ͻú����� | Y or N
		strRst = strRst & "				<bbprcSvmtPrdcYn>Y</bbprcSvmtPrdcYn>"						'���ø�����ǥ�� H.Point������ | Y or N
		strRst = strRst & "				<bbprcSpdcYn>Y</bbprcSpdcYn>"								'���ø�����ǥ�� ��¦���� | Y or N
		strRst = strRst & "				<prcExpsBitVal1>1</prcExpsBitVal1>"							'���ݺ񱳳��Ⱑ �������� | 0: �ش����, 1: �ش�
		strRst = strRst & "				<prcExpsBitVal2>2</prcExpsBitVal2>"							'���ݺ񱳳��Ⱑ �Ͻú����� | 0: �ش����, 2: �ش�
		strRst = strRst & "				<prcExpsBitVal4>4</prcExpsBitVal4>"							'���ݺ񱳳��Ⱑ H.Point������ | 0: �ش����, 4: �ش�
		strRst = strRst & "				<prcExpsBitVal8>8</prcExpsBitVal8>"							'���ݺ񱳳��Ⱑ ��¦���� | 0: �ش����, 8: �ش�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		If CSafeyn = "Y" Then
			strRst = strRst & getHmallItemSafeInfoToReg(CSafeGbCd, "U")
		End If
'		strRst = strRst & "	<Dataset id=""dsSlitmBcdDtl"">"
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>UPDATE</rowType>"
'		strRst = strRst & "				<bcdBrndGbcd></bcdBrndGbcd>"								'���ڵ�귣�屸���ڵ� | 10 �Ϲݺ귣��, 20 ��ȭ���귣��
'		strRst = strRst & "				<shrtBcdVal></shrtBcdVal>"									'������ڵ尪
'		strRst = strRst & "				<totBcdVal></totBcdVal>"									'��ü���ڵ尪
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getOptTypeMst("U")
		strRst = strRst & getOptAttrMst("U")
		strRst = strRst & getOptCombDtl("U")

'		strRst = strRst & "	<Dataset id=""dsAsctSlitmDtl"">"
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>UPDATE</rowType>"
'		strRst = strRst & "				<asctItemGbcd></asctItemGbcd>"								'���û�ǰ�����ڵ� | 10 Ʈ����H
'		strRst = strRst & "				<asctSlitmCd></asctSlitmCd>"								'�����ǸŻ�ǰ�ڵ�
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallItemInfoCdToReg("U")

		strRst = strRst & "	<Dataset id=""dsHtmlItstMst"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<htmlItstGbcd>00</htmlItstGbcd>"							'HTML��ǰ����������ڵ� | 00 �Ϲ�, 01 ��ǰ, 02 ��������
		strRst = strRst & "				<htmlItstCntn><![CDATA["&getHmallContParamToReg()&"]]></htmlItstCntn>"	'��ǰ����������ڵ� | HTML��ǰ���������
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsMdaSlitmDtl"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"								'�ǸŸ�ü�з��ڵ�
		strRst = strRst & "			</row>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>04</sellMdaCsfCd>"								'�ǸŸ�ü�з��ڵ�
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallSectIdToReg("U")
		strRst = strRst & "</Root>"
		gethmallItemEditParameter = strRst
	End Function

	Public Function getHmallPriceParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""sessionVO"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>�ٹ�����</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "	<Dataset id=""dsItemPrcHistTran"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"							'#��ǰ�ڵ�
		strRst = strRst & "				<prcAplyStrtDtm>"&Replace(Date(), "-", "")&"</prcAplyStrtDtm>"	'#���������� | 20131204
		strRst = strRst & "				<prcAplyStrtTime></prcAplyStrtTime>"							'����������۽ð� | �������� ��-�� (����4�� 30���� 1630���� �ۼ�)
		strRst = strRst & "				<prcDcEndDtm></prcDcEndDtm>"									'����������	| 20131208
		strRst = strRst & "				<prcDcEndTime></prcDcEndTime>"									'��������ð� | �������� ��-�� (����4�� 30���� 1630���� �ۼ�)
		strRst = strRst & "				<prcAthzGbcd>00</prcAthzGbcd>"									'#�����û�����ڵ� | 00 : ��û, 41: ��û���
		strRst = strRst & "				<sellPrc>"&MustPrice()&"</sellPrc>"								'#�ǸŰ���
		strRst = strRst & "				<mrgnRate>"&FMrgnRate&"</mrgnRate>"								'#������
		strRst = strRst & "				<dptsVenOpCd></dptsVenOpCd>"									'��ȭ�����»�OP�ڵ�	String(2)	
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"								'���»��ǰ�ڵ�	String(20)	
		strRst = strRst & "				<prmoCopyYn>Y</prmoCopyYn>"										'���θ�Ǻ��翩�� | ���� ���θ������(������, ����, �Ͻú�����, ������)��  �ű� ���ݼ����� �״�� �����ؼ� ����� ��� Y
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "</Root>"
		getHmallPriceParameter = strRst
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CHmall
	Public FItemList()
	Public FOneItem
	Public FResultCount
	Public FTotalCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount

	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectMakerid
	Public FRectMatchCate
	Public FRectMatchShipping
	Public FRectGosiEqual
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectFailCntExists
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectNotinmakerid
	Public FRectPriceOption
	Public FRectMwdiv

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType

	Public FRectDeliveryType

	Public Sub getHmallNotRegOnlyOneItem
		strSql = ""
		strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Hmall_Reg_Get] " & FRectItemID
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
				FOneItem.FAdultType			= db2html(rsget("adultType"))
				FOneItem.Fvatinclude		= rsget("vatinclude")
				FOneItem.FordMakeYn			= rsget("ordMakeYn")
		End If
		rsget.Close
	End Sub

	Public Sub getHmallNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "		SELECT itemid FROM ("
			addSql = addSql & "			SELECT itemid"
			addSql = addSql & " 		,count(*) as optCNT"
			addSql = addSql & " 		,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
			addSql = addSql & " 		,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
			addSql = addSql & " 		FROM db_item.dbo.tbl_item_option"
			addSql = addSql & " 		WHERE itemid in (" & FRectItemID & ")"
			addSql = addSql & " 		and isusing='Y'"
			addSql = addSql & " 		GROUP BY itemid"
			addSql = addSql & " 	) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.hmallStatCD,-9) as hmallStatCD, isNull(R.hmallPrice, 0) as hmallPrice "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & "	,(SELECT [db_etcmall].[dbo].[getHmallMargin] (" & FRectItemID & ")) as mrgnRate"
		strSql = strSql & "	,S.octyCnryGbcd, S.octyCnryNm"
		strSql = strSql & "	,LEFT(am.CateKey, 2) as itemLCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 4) as itemMCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 6) as itemSCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 8) as itemCsfGbcd "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_hmall_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_hmall_sourceCodeName] (" & FRectItemID & ") as S on i.itemid = S.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_category as tm on am.CateKey = tm.CateKey "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
'		strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "		'�ö��/ȭ�����/�ؿ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo - i.LimitSold > "&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.

'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100) >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((sellcash - buycash)/ sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((orgprice - orgsuplycash)/ orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN sailyn = 'N' AND (Round(((sellcash - buycash)/ sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and isnull(R.hmallStatCD,0) < 3  "
		strSql = strSql & " and cm.mapCnt is Not Null "		'ī�װ� ��Ī ��ǰ��
		strSql = strSql & " and i.itemdiv not in ('06') "	'�ֹ����۹��� ��ǰ ����
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
		End If
		rsget.Close
	End Sub

	Public Sub getHmallEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & " ,m.hmallGoodNo, m.hmallSellyn, m.regImageName, isNull(m.hmallprice, 0) as hmallprice "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or ((i.sailyn <> 'Y') and (i.sellcash + round(i.orgprice * 0.5, 0) < m.hmallprice)) "	'������ �ƴϰ� ������ 50%�̻��� �ǸŰ��� hmall�� ��ϵ� ���
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & " 	or i.itemdiv in ('06') "		''�ֹ����۹��� ��ǰ ǰ��ó��
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "
		strSql = strSql & "		or ((i.sellcash < 50000) AND (i.itemname like '%������%' or i.itemname like '%���� ���%')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & "	,(SELECT [db_etcmall].[dbo].[getHmallMargin] (" & FRectItemID & ")) as mrgnRate"
		strSql = strSql & "	,S.octyCnryGbcd, S.octyCnryNm"
		strSql = strSql & "	,LEFT(am.CateKey, 2) as itemLCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 4) as itemMCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 6) as itemSCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 8) as itemCsfGbcd "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_hmall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_hmall_sourceCodeName] (" & FRectItemID & ") as S on i.itemid = S.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Hmall_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.hmallGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FHmallGoodNo		= rsget("hmallGoodNo")
				FOneItem.FHmallSellYn		= rsget("hmallSellYn")
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName		= rsget("regImageName")
				FOneItem.FHmallprice		= rsget("hmallprice")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
		End If
		rsget.Close
	End Sub

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
End Class

'// ��ǰ�̹��� ���翩�� �˻�
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

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

'Hmall ��ǰ�ڵ� ���
Function getHmallGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ISNULL(hmallgoodno, '') as hmallgoodno FROM db_etcmall.dbo.tbl_hmall_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getHmallGoodno = rsget("hmallgoodno")
	End If
	rsget.Close
End Function

'Hmall ��ǰ�ڵ� ���
Function getHmallGoodno2(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ISNULL(hmallgoodno, '') as hmallgoodno FROM db_etcmall.dbo.tbl_hmall_regitem WHERE itemid = '"&iitemid&"' and APIaddImg = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getHmallGoodno2 = rsget("hmallgoodno")
	End If
	rsget.Close
End Function

'�ٹ����� �⺻ �̹��� ���
Function getTenBasicImage(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT basicimage " & VBCRLF
	strSql = strSql & " FROM db_item.dbo.tbl_item  " & VBCRLF
	strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getTenBasicImage = rsget("basicimage")
	End If
	rsget.Close
End Function
%>