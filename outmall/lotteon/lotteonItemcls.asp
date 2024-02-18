<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "lotteon"
CONST CUPJODLVVALID = TRUE			''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5				'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST APIATTRURL = "https://onpick-api.lotteon.com"
CONST CDEFALUT_STOCK = 99999

Class CLotteonItem
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
	Public FLotteonStatCD
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
	Public FLotteonGoodNo
	Public FLotteonprice
	Public FLotteonSellYn
	Public FStd_cat_id
	Public FDisp_cat_id

	Public FAdultType
	Public FLastStatCheckDate
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999999" Then
			getOrderMaxNum = 999999
		End If
	End Function

	Public Function getRegedOptionCnt
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as Cnt  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption "
		sqlStr = sqlStr & " WHERE mallid= 'lotteon' "
		sqlStr = sqlStr & " and itemoption <> '0000' "
		sqlStr = sqlStr & " and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			getRegedOptionCnt = rsget("Cnt")
		rsget.Close
	End Function

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

	Public Function getLimitEa()
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

			If FLotteonPrice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					' If (FSellCash < Round(FSsgprice * 0.25, 0)) Then
					' 	MustPrice = CStr(GetRaiseValue(Round(FSsgprice * 0.25, 0)/10)*10)
					' Else
						MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					' End If
				End If
			End If
		End If
	End Function

	'// Lotteon �Ǹſ��� ��ȯ
	Public Function getLotteonSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getLotteonSellYn = "Y"
			Else
				getLotteonSellYn = "N"
			End If
		Else
			getLotteonSellYn = "N"
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

	Public Function getShopLeadTime()
		Dim CateLargeMid, leadTime
		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "040010", "040011", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "040121", "055070", "055080", "055090", "055100", "055110", "055120", "055222"
				leadTime = 15
			Case "045001",  "045002", "045003", "045004", "045005", "045006", "045007", "045008", "045009", "045010", "045011", "045012"
			 	leadTime = 10
			Case "070010", "070020", "070030", "070040", "070050", "070070", "070110", "070120", "070140", "070150", "070160", "070200", "070201", "070202", "070203", "080007", "080010", "080020", "080030", "080031", "080040", "080050", "080051", "080060", "080070", "080071", "080080", "080090", "090005", "090010", "090011", "090020", "090030", "090040", "090050", "090060", "090061", "090070", "090071", "090080"
				leadTime = 7
			Case "050010", "050020", "050030", "050040", "050045", "050050", "050070", "050110", "050120", "050666", "050777", "060010", "060020", "060040", "060050", "060060", "060070", "060080", "060090", "060120", "060130", "060140", "060150", "060160", "100010", "100020", "100030", "100040", "100060", "100070", "100080", "100090", "100100", "100110", "100120", "100130", "100140", "100150", "100201", "100300"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select
		getShopLeadTime = leadTime
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
        buf = LeftB(buf, 130)
        getItemNameFormat = buf
    end function

	Public function getOriginCode()
		If Fsourcearea = "�ѱ�" OR Fsourcearea = "���ѹα�" Then
			getOriginCode = "KR"
		Else
			getOriginCode = "ETC"
		End If
	End Function


	Public function getBrandCode()
		Select Case Fmakerid
			Case "disney10x10"
				getBrandCode = "P778"
			Case "sanrio10x10"
				getBrandCode = "P47543"
			Case "universal10x10"
				getBrandCode = "P11805"
			Case "peanuts10x10"
				getBrandCode = "P5270"
			Case "sanx10x10"
				getBrandCode = "P15324"
			Case "cncglobalkr"
				getBrandCode = "P2399"
			Case Else
				getBrandCode = ""
		End Select
	End Function

    public function getItemNameFormat2()
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
        getItemNameFormat2 = buf
    end function

	Public Function getLotteonKeywordsParameter(obj)
		Dim arrRst, arrRst2, q, Keyword1, strRst
		Dim retKeyword, i, commaSplit
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
			retKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			retKeyword = keyword1
		End If

		If retKeyword = "" Then
			Set obj("spdLst")(null)("scKwdLst") = jsArray()
				obj("spdLst")(null)("scKwdLst") = null
		Else
			commaSplit = Split(retKeyword,",")
			Set obj("spdLst")(null)("scKwdLst") = jsArray()								'�˻�Ű������ | 5�� ���ϸ� ��� ����
				For i = 0 To Ubound(commaSplit)
					obj("spdLst")(null)("scKwdLst")(i) = commaSplit(i)
				Next
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

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
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

	'��ǰ�����������
	Public Function getLotteonInfoCdParameter(obj)
		Dim strSql, buf, i
		Dim mallinfoCd, infoContent, mallinfodiv
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00002') THEN '�������� ����' "
		strSql = strSql & "     WHEN (M.infoCd='10000') THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' "
		strSql = strSql & "     WHEN c.infotype='P' THEN '�ٹ����� ���ູ���� 1644-6035' "
		strSql = strSql & " 	WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '�������� ����' "
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & "  SELECT TOP 1 itemid, certNum FROM db_item.dbo.tbl_safetycert_tenReg where itemid = '"&FItemID&"' "
		strSql = strSql & " ) as tr on I.itemid = tr.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = '"& CMALLNAME &"' and IC.itemid='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			If CStr(rsget("mallinfodiv")) = "35"  Then
				mallinfodiv = "38"
			ElseIf CStr(rsget("mallinfodiv")) = "47"  Then
				mallinfodiv = "39"
			ElseIf CStr(rsget("mallinfodiv")) = "48"  Then
				mallinfodiv = "40"
			Else
				mallinfodiv = CStr(rsget("mallinfodiv"))
			End If

			Set obj("spdLst")(null)("pdItmsInfo") = jsObject()
				obj("spdLst")(null)("pdItmsInfo")("pdItmsCd") = mallinfodiv
				Set obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst") = jsArray()						'#��ǰǰ���׸���
			i = 0
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
				End If

				Set obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i) = jsObject()
					obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i)("pdArtlCd") = mallinfoCd		'#��ǰ�׸��ڵ�
					obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i)("pdArtlCnts") = infoContent	'#��ǰ�׸񳻿� | �ش� ��������׸��� �׸��� �Է��Ѵ�.

				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
    End Function

	'����������� �Ķ���� ����
	Public Function getLotteonCertInfoParameter(obj)
		Dim strSql
		Dim safetyDiv, safetyId, certNum, certOrganName, isRegCert
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN 'ELC_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN 'ELC_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN 'ELC_SUPS' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN 'LIFE_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN 'LIFE_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN 'LIFE_SUPS' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN 'CHL_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN 'CHL_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN 'CHL_SUPS' end as safetyId "
		strSql = strSql & " , t.certNum, f.certOrganName, f.makerName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			certNum			= rsget("certNum")
			certOrganName	= rsget("certOrganName")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			If safetyDiv = "30" OR safetyDiv = "60" OR safetyDiv = "90" Then
				certNum = ""
			End If

			Set obj("spdLst")(null)("sftyAthnLst") = jsArray()
				Set obj("spdLst")(null)("sftyAthnLst")(0) = jsObject()
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnTypCd") = safetyId		'#�������������ڵ� [�����ڵ� : SFTY_ATHN_TYP_CD] | CHL_SUPS : [�����ǰ]���������ռ�Ȯ��, CHL_ATHN : [�����ǰ]��������, CHL_CFM : [�����ǰ]����Ȯ��, CMCN_TNTT : [�����ű�����]��������, CMCN_REG : [�����ű�����]���յ��, CMCN_ATHN : [�����ű�����]��������, LIFE_SUPS : [��Ȱ��ǰ]���������ռ�Ȯ��, LIFE_ATHN : [��Ȱ��ǰ]��������, LIFE_CFM : [��Ȱ��ǰ]����Ȯ��, ELC_SUPS : [�����ǰ]���������ռ�Ȯ��, ELC_ATHN : [�����ǰ]��������, ELC_CFM : [�����ǰ]����Ȯ��, LIFE_STD : [��Ȱ��ǰ]���������ؼ�, CHEM_LIFE : [ȭ����ǰ] ��Ȱȭ����ǰ ������������Ȯ�νŰ��ȣ / ���ι�ȣ, CHEM_BIOC : [ȭ����ǰ] �������ǰ ���ι�ȣ, ETC : ��Ÿ
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnOrgnNm") = certOrganName	'�������������
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnNo") = certNum			'����������ȣ
		Else
			Set obj("spdLst")(null)("sftyAthnLst") = jsArray()
				obj("spdLst")(null)("sftyAthnLst") = null
		End If
	End Function

	'ǥ��ī�װ��Ӽ����
	Public Function getLotteonStdCateAttrParameter(obj)
		' Set obj("spdLst")(null)("scatAttrLst") = jsArray()
		' 	Set obj("spdLst")(null)("scatAttrLst")(null) = jsObject()
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optCd") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optValCd") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optVal") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("dtlsVal") = ""
		Set obj("spdLst")(null)("scatAttrLst") = jsArray()
			obj("spdLst")(null)("scatAttrLst") = null
	End Function

	'��ǰ���� �Ķ���� ����
	Public Function getLotteonContParamToReg(obj)
		Dim strRst, strSQL, retContents, retOrderComment
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteon.jpg></p><br />")
		strRst = strRst & ("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & ("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & ("<tr>")
		strRst = strRst & ("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""��ǰ��"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & ("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '����', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & ("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat2
		strRst = strRst & ("</p>")
		strRst = strRst & ("</td>")
		strRst = strRst & ("</tr>")
		strRst = strRst & ("</table>")
		strRst = strRst & ("</div>")

		If ForderComment <> "" Then
			strRst = strRst & "<div align=""center""><br />" & nl2br(Fordercomment) & "<br /></div>"
		End If

		If Fitemsize <> "" Then
			strRst = strRst & "- ������ : " & Fitemsize & "<br />"
		End if

		If Fitemsource <> "" Then
			strRst = strRst & "- ��� : " &  Fitemsource & "<br />"
		End If

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br />")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
		End Select
		'# �߰� ��ǰ �����̹��� ����
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br />")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br />")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br />")

		'#��� ���ǻ���
		strRst = strRst & ("<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_lotteon.jpg>")
		strRst = strRst & ("</div>")
		retContents = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = rsget("textVal")
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteon.jpg></p><br />" & strRst & "<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_lotteon.jpg></div>"
			retContents = strRst
		End If
		rsget.Close

		Set obj("spdLst")(null)("epnLst") = jsArray()
			Set obj("spdLst")(null)("epnLst")(0) = jsObject()
				obj("spdLst")(null)("epnLst")(0)("pdEpnTypCd") = "DSCRP"		'#��ǰ���������ڵ� [�����ڵ� : PD_EPN_TYP_CD] | DSCRP : ��ǰ�����, AS_CNTS : A/S���뼳��, PRCTN : ���ǻ��׼���
				obj("spdLst")(null)("epnLst")(0)("cnts") = retContents			'#���� | html�Է½� ����Ѵ�.
		' If ForderComment <> "" Then
		' 	retOrderComment = "<div align=""center""><br />" & Fordercomment & "<br /></div>"
		' 	Set obj("spdLst")(null)("epnLst")(1) = jsObject()
		' 		obj("spdLst")(null)("epnLst")(1)("pdEpnTypCd") = "PRCTN"		'#��ǰ���������ڵ� [�����ڵ� : PD_EPN_TYP_CD] | DSCRP : ��ǰ�����, AS_CNTS : A/S���뼳��, PRCTN : ���ǻ��׼���
		' 		obj("spdLst")(null)("epnLst")(1)("cnts") = retOrderComment		'#���� | html�Է½� ����Ѵ�.
		' End If
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

	Public Function getLotteonAddImageParam(obj)
		Dim addImages
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP 30 gubun, ImgType, addimage_400, addimage_600, addimage_1000 "
		strSql = strSql & " FROM db_item.[dbo].tbl_item_addimage "
		strSql = strSql & " WHERE itemid=" & Fitemid
		strSql = strSql & " and isnull(addimage_400, '') <> '' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				addImages = addImages & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & "|"
				rsget.MoveNext
				If i>=9 Then Exit For
			Next
		End If
		rsget.Close

		If Right(addImages,1) = "|" Then
			addImages = Left(addImages,Len(addImages)-1)
		End If
		getLotteonAddImageParam = addImages
	End Function

	Public Function getLotteonOptionParameter(obj)
		Dim addImages, addImgSplit, i, j
		Dim limitsu, strSql
		Dim vlimitno, vlimitsold, vitemoption, voptionname, voptlimitno, voptlimitsold, voptsellyn, voptlimityn, voptaddprice
		Dim vMustprice
		vMustprice = mustPrice()
		addImages = getLotteonAddImageParam(obj)
		addImgSplit = Split(addImages, "|")

		If FOptionCnt = 0 Then			'��ǰ
			obj("spdLst")(null)("sitmYn") = "N"													'#�Ǹ��ڴ�ǰ���� [Y, N] | Y�̸� ��ǰ�Ӽ������ �����ؾ� �Ѵ�. N�̸� ��ǰ�Ӽ������ ���� ���Ѵ�. �ɼ��� ���� ��ǰ �Ѱ����� �����ȴ�.
			Set obj("spdLst")(null)("itmLst") = jsArray()										'��ǰ���
				Set obj("spdLst")(null)("itmLst")(null) = jsObject()
					obj("spdLst")(null)("itmLst")(null)("eitmNo") = "0000"						'��ü��ǰ��ȣ
					obj("spdLst")(null)("itmLst")(null)("sortSeq") = "1"						'#���ļ���
					obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"							'#���ÿ��� [Y, N]
					Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()			'�������ܿ��ܸ�� [�����ڵ� : PY_MNS_CD]
						obj("spdLst")(null)("itmLst")(null)("itmOptLst") = null
					Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()			'��ǰ�̹������ | ��ǰ�� �ϳ� �̻��� �̹����� ����Ͽ��� �Ѵ�. ��ǰ�� �ִ� 10���� �̹����� ����� �� �ִ�.
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
						If IsArray(addImgSplit) = True Then
							For i = 1 to Ubound(addImgSplit) + 1
								Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i) = jsObject()
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("origImgFileNm") = addImgSplit(i-1) '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("rprtImgYn") = "N"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
							Next
						End If
	'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'�÷�Ĩ�̹������
	'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
	'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
	'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'��ǰ������������
	'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#��ǰ�뷮 | ���ش����� ���ؿ뷮�� ǥ��ī�װ� ���� ������ ������. ex) ǥ��ī�װ��� ���ش����� ml, ���ؿ뷮�� 100���� ���εǾ� �ִ� ��� 100ml�� ������ ǥ�õȴ�.
					obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice						'#�ǸŰ�
					obj("spdLst")(null)("itmLst")(null)("stkQty") = getLimitEa()					'#������ | ���������ΰ� Y�� ��쿡�� �ʼ���
		Else							'�ɼ�
			obj("spdLst")(null)("sitmYn") = "Y"													'#�Ǹ��ڴ�ǰ���� [Y, N] | Y�̸� ��ǰ�Ӽ������ �����ؾ� �Ѵ�. N�̸� ��ǰ�Ӽ������ ���� ���Ѵ�. �ɼ��� ���� ��ǰ �Ѱ����� �����ȴ�.
			Set obj("spdLst")(null)("itmLst") = jsArray()											'��ǰ���

			Dim vattr_id, vattr_nm
			Dim vattr_val_id, vattr_val_nm
			strSql = ""
			strSql = strSql & " SELECT TOP 1 a.attr_id, a.attr_nm, a.attr_disp_nm "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_Attribute as a "
			strSql = strSql & " WHERE attr_id in ( "
			strSql = strSql & " 	SELECT attr_id FROM db_etcmall.dbo.tbl_lotteon_StdCategory_Attr WHERE std_cat_id = '"& FStd_cat_id &"'  "
			strSql = strSql & " ) and attr_pi_type= 'I' and attr_disp_nm = '����'  "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				vattr_id = rsget("attr_id")
				vattr_nm = rsget("attr_nm")
			End If
			rsget.Close

			If vattr_id = "" Then
				rw "�´� �Ӽ� ����"
				Exit Function
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 attr_val_id, attr_val_nm FROM db_etcmall.dbo.tbl_lotteon_Attribute_Values "
				strSql = strSql & " WHERE attr_id = '"& vattr_id &"' and attr_val_nm = '��Ƽ' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					vattr_val_id = rsget("attr_val_id")
					vattr_val_nm = rsget("attr_val_nm")
				End If
				rsget.Close
			End If

			If vattr_val_id = "" Then
				rw "�´� �Ӽ��� ����"
				Exit Function
			End If

			strSql = ""
			strSql = strSql & " SELECT i.itemid, i.limityn, i.limitno ,i.limitsold, o.itemoption, optionname" & VBCRLF
			strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice " & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			strSql = strSql & " WHERE i.itemid = "&Fitemid
			strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				For i = 1 to rsget.RecordCount
					If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''���ϻ�ǰ
						vitemoption = "0000"
						voptionname = "���ϻ�ǰ"
						limitsu = getLimitEa()
						voptaddprice		= 0
					Else
						vitemoption 		= rsget("itemoption")
						voptionname 		= rsget("optionname")
						voptlimitno 		= rsget("optlimitno")
						voptlimitsold 		= rsget("optlimitsold")
						voptaddprice		= rsget("optaddprice")
						If FLimityn = "Y" Then
							If voptlimitno - voptlimitsold - 5 < 1 Then
								limitsu = 0
							Else
								limitsu = voptlimitno - voptlimitsold - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If
					End If
					Set obj("spdLst")(null)("itmLst")(null) = jsObject()
						obj("spdLst")(null)("itmLst")(null)("eitmNo") = vitemoption						'��ü��ǰ��ȣ
						obj("spdLst")(null)("itmLst")(null)("sortSeq") = i								'#���ļ���
						obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"								'#���ÿ��� [Y, N]
						Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()				'o��ǰ�Ӽ���� | �Ǹ��ڴ�ǰ���ΰ� Y�� ��� �ʼ���
							Set obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optCd") = vattr_id	'#�ɼ��ڵ� [�Ӽ���� ���� �׸�] | ��ǰ�� �ɼǿ� �ش��ϴ� �ɼ��ڵ带 �Է��Ͽ��� �Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optNm") = vattr_nm	'#�ɼǸ� [�Ӽ���� ���� �׸�] | �ش� ��ǰ�� �ɼǸ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optValCd") = vattr_val_id	'o�ɼǰ��ڵ� [�Ӽ���� ���� �׸�] | �Է��ϰ��� �ϴ� �ɼǰ��� �ɼǰ��ڵ尡 �������� �ʴ� ��쿡�� �ɼǰ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optVal") = vattr_val_nm	'o�ɼǰ� [�Ӽ���� ���� �׸�] | �ش� ��ǰ�� �ɼǰ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("dtlsVal") = voptionname	'���ΰ� | ���ΰ��� �Է��ϴ� ��� 1. �������� ���� ������ �Է½�, 2. �ɼǰ��� ���� �߰� ǥ��
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'��ǰ�̹������ | ��ǰ�� �ϳ� �̻��� �̹����� ����Ͽ��� �Ѵ�. ��ǰ�� �ִ� 10���� �̹����� ����� �� �ִ�.
							Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
							If IsArray(addImgSplit) = True Then
								For j = 1 to Ubound(addImgSplit) + 1
									Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j) = jsObject()
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("origImgFileNm") = addImgSplit(j-1) '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("rprtImgYn") = "N"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
								Next
							End If
		'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'�÷�Ĩ�̹������
		'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
		'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
		'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'��ǰ������������
		'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#��ǰ�뷮 | ���ش����� ���ؿ뷮�� ǥ��ī�װ� ���� ������ ������. ex) ǥ��ī�װ��� ���ش����� ml, ���ؿ뷮�� 100���� ���εǾ� �ִ� ��� 100ml�� ������ ǥ�õȴ�.
						obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice + voptaddprice		'#�ǸŰ�
						obj("spdLst")(null)("itmLst")(null)("stkQty") = limitsu							'#������ | ���������ΰ� Y�� ��쿡�� �ʼ���
					rsget.MoveNext
				Next
			End If
			rsget.Close
		End If
	End Function

	Public Function getLotteonOptionEditParameter(obj)
		Dim addImages, addImgSplit, i, j
		Dim limitsu, strSql, arrRows
		Dim vlimitno, vlimitsold, vitemoption, voptionname, voptlimitno, voptlimitsold, voptsellyn, voptlimityn, voptaddprice
		Dim vMustprice, sitmNo
		vMustprice = mustPrice()
		addImages = getLotteonAddImageParam(obj)
		addImgSplit = Split(addImages, "|")

		strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_optEditParamList_lotteon '"&CMallName&"'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
		    arrRows = rsget.getRows
		End If
		rsget.close

		If UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z" Then
			sitmNo = arrRows(15, 0)
		End If

		If FOptionCnt = 0 AND (UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z") Then			'��ǰ
			Set obj("spdLst")(null)("itmLst") = jsArray()										'��ǰ���
				Set obj("spdLst")(null)("itmLst")(null) = jsObject()
					obj("spdLst")(null)("itmLst")(null)("eitmNo") = "0000"						'��ü��ǰ��ȣ
					obj("spdLst")(null)("itmLst")(null)("sitmNo") = ""&sitmNo&""				'�Ǹ��ڴ�ǰ��ȣ
					obj("spdLst")(null)("itmLst")(null)("sortSeq") = "1"						'���ļ���
					obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"							'���ÿ��� [Y, N]
					Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()			'�������ܿ��ܸ�� [�����ڵ� : PY_MNS_CD]
						obj("spdLst")(null)("itmLst")(null)("itmOptLst") = null
					Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'��ǰ�̹������ | ��ǰ�� �ϳ� �̻��� �̹����� ����Ͽ��� �Ѵ�. ��ǰ�� �ִ� 10���� �̹����� ����� �� �ִ�.
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
						If IsArray(addImgSplit) = True Then
							For i = 1 to Ubound(addImgSplit) + 1
								Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i) = jsObject()
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("origImgFileNm") = addImgSplit(i-1) '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("rprtImgYn") = "N"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
							Next
						End If
	'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'�÷�Ĩ�̹������
	'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
	'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
	'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'��ǰ������������
	'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#��ǰ�뷮 | ���ش����� ���ؿ뷮�� ǥ��ī�װ� ���� ������ ������. ex) ǥ��ī�װ��� ���ش����� ml, ���ؿ뷮�� 100���� ���εǾ� �ִ� ��� 100ml�� ������ ǥ�õȴ�.
					obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice						'#�ǸŰ�
					obj("spdLst")(null)("itmLst")(null)("stkQty") = getLimitEa()					'#������ | ���������ΰ� Y�� ��쿡�� �ʼ���
		Else							'�ɼ�
			Set obj("spdLst")(null)("itmLst") = jsArray()											'��ǰ���

			Dim vattr_id, vattr_nm
			Dim vattr_val_id, vattr_val_nm
			strSql = ""
			strSql = strSql & " SELECT TOP 1 a.attr_id, a.attr_nm, a.attr_disp_nm "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_Attribute as a "
			strSql = strSql & " WHERE attr_id in ( "
			strSql = strSql & " 	SELECT attr_id FROM db_etcmall.dbo.tbl_lotteon_StdCategory_Attr WHERE std_cat_id = '"& FStd_cat_id &"'  "
			strSql = strSql & " ) and attr_pi_type= 'I' and attr_disp_nm = '����'  "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				vattr_id = rsget("attr_id")
				vattr_nm = rsget("attr_nm")
			Else
				rw "�´� �Ӽ� ����"
				Exit Function
			End If
			rsget.Close

			If vattr_id <> "" Then
				strSql = ""
				strSql = strSql & " SELECT attr_val_id, attr_val_nm FROM db_etcmall.dbo.tbl_lotteon_Attribute_Values "
				strSql = strSql & " WHERE attr_id = '"& vattr_id &"' and attr_val_nm = '��Ƽ' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					vattr_val_id = rsget("attr_val_id")
					vattr_val_nm = rsget("attr_val_nm")
				Else
					rw "�´� �Ӽ��� ����"
					rsget.Close
					Exit Function
				End If
				rsget.Close
			End If

			If IsArray(arrRows) Then
				For i = 0 To UBound(arrRows, 2)
					vitemoption 		= arrRows(1, i)
					voptionname 		= arrRows(3, i)
					voptaddprice		= arrRows(16, i)
					sitmNo 				= arrRows(15, i)
					If FLimityn = "Y" Then
						If arrRows(4, i) - 5 < 1 Then
							limitsu = 0
						Else
							limitsu = arrRows(4, i) - 5
						End If
					Else
						limitsu = CDEFALUT_STOCK
					End If

					Set obj("spdLst")(null)("itmLst")(null) = jsObject()
						obj("spdLst")(null)("itmLst")(null)("eitmNo") = vitemoption						'��ü��ǰ��ȣ
						obj("spdLst")(null)("itmLst")(null)("sitmNo") = ""&sitmNo&""					'�Ǹ��ڴ�ǰ��ȣ
						obj("spdLst")(null)("itmLst")(null)("sortSeq") = i								'#���ļ���
						obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"								'#���ÿ��� [Y, N]

					If (ArrRows(11,i)=0) and ArrRows(12,i) = "1" AND ArrRows(15,i) = "" Then		'�ɼǸ��� �ٸ��� �ɼ��ڵ尪�� ���� �� ==> ��ǰ�߰� �ǹ�// preged 0
						Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()				'o��ǰ�Ӽ���� | �Ǹ��ڴ�ǰ���ΰ� Y�� ��� �ʼ���
							Set obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optCd") = vattr_id	'#�ɼ��ڵ� [�Ӽ���� ���� �׸�] | ��ǰ�� �ɼǿ� �ش��ϴ� �ɼ��ڵ带 �Է��Ͽ��� �Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optNm") = vattr_nm	'#�ɼǸ� [�Ӽ���� ���� �׸�] | �ش� ��ǰ�� �ɼǸ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optValCd") = vattr_val_id	'o�ɼǰ��ڵ� [�Ӽ���� ���� �׸�] | �Է��ϰ��� �ϴ� �ɼǰ��� �ɼǰ��ڵ尡 �������� �ʴ� ��쿡�� �ɼǰ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optVal") = vattr_val_nm	'o�ɼǰ� [�Ӽ���� ���� �׸�] | �ش� ��ǰ�� �ɼǰ��� �Է��Ѵ�.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("dtlsVal") = voptionname	'���ΰ� | ���ΰ��� �Է��ϴ� ��� 1. �������� ���� ������ �Է½�, 2. �ɼǰ��� ���� �߰� ǥ��
					End If

						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'��ǰ�̹������ | ��ǰ�� �ϳ� �̻��� �̹����� ����Ͽ��� �Ѵ�. ��ǰ�� �ִ� 10���� �̹����� ����� �� �ִ�.
							Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
							If IsArray(addImgSplit) = True Then
								For j = 1 to Ubound(addImgSplit) + 1
									Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j) = jsObject()
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypCd") = "IMG"	'#���������ڵ� [�����ڵ� : EPSR_TYP_CD] | IMG : �̹���
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypDtlCd") = "IMG_SQRE"	'#�����������ڵ� [�����ڵ� : EPSR_TYP_DTL_CD] | IMG_SQRE : ��������:�̹��� > ���簢��, IMG_LNTH : ��������:�̹��� > ������
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("origImgFileNm") = addImgSplit(j-1) '#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("rprtImgYn") = "N"	'#��ǥ�̹������� [Y, N] | ��ǥ�̹����� �ϳ��� ���� ����
								Next
							End If
		'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'�÷�Ĩ�̹������
		'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
		'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#�����̹������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.jpg
		'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'��ǰ������������
		'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#��ǰ�뷮 | ���ش����� ���ؿ뷮�� ǥ��ī�װ� ���� ������ ������. ex) ǥ��ī�װ��� ���ش����� ml, ���ؿ뷮�� 100���� ���εǾ� �ִ� ��� 100ml�� ������ ǥ�õȴ�.
						obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice + voptaddprice		'#�ǸŰ�
						obj("spdLst")(null)("itmLst")(null)("stkQty") = limitsu							'#������ | ���������ΰ� Y�� ��쿡�� �ʼ���

				Next
			End If
		End If
	End Function

	'��ǰ��� Json
	Public Function getLotteonItemRegParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj, tenBeasongDay
		tenBeasongDay = getShopLeadTime()

		If FItemdiv = "06" OR FItemdiv = "16" Then
			dvPdTypCd = "OD_MFG"
			sndBgtNday = "15"
		Else
			If tenBeasongDay > 3 Then
				dvPdTypCd = "OD_MFG"
				sndBgtNday = tenBeasongDay
			Else
				dvPdTypCd = "GNRL"
				sndBgtNday = "3"
			End If
		End If

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()														'��ϻ�ǰ���
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"										'#�ŷ�ó�׷��ڵ� | SR : �Ϲݼ���
					obj("spdLst")(null)("trNo") = afflTrCd										'#�ŷ�ó��ȣ
					obj("spdLst")(null)("lrtrNo") = ""											'�����ŷ�ó��ȣ
					obj("spdLst")(null)("scatNo") = FStd_cat_id									'#ǥ��ī�װ���ȣ
					Set obj("spdLst")(null)("dcatLst") = jsArray()								'#����ī�װ���� | �Ӽ������ API�� ���Ͽ� ǥ��ī�װ��� ���ε� ����ī�װ��� ������ �޴´�. ���ε� ����ī�װ� �߿��� �ϳ� �̻� �����Ͽ� �Է��Ѵ�.
						Set obj("spdLst")(null)("dcatLst")(null) = jsObject()
							obj("spdLst")(null)("dcatLst")(null)("mallCd") = "LTON"				'#�������ڵ� | LTON : �Ե�ON
							obj("spdLst")(null)("dcatLst")(null)("lfDcatNo") = FDisp_cat_id		'#leaf����ī�װ���ȣ
'							obj("spdLst")(null)("dcatLst")(null)("dcatNo") = ""					'--���ô� �ִµ� ������ ����..;
					obj("spdLst")(null)("epdNo") = ""&FItemid&""								'��ü��ǰ��ȣ
					obj("spdLst")(null)("slTypCd") = "GNRL"										'#�Ǹ������ڵ� | ����ǰ�� ����ǰ��� API�� ����Ѵ�. GNRL : �Ϲ��ǸŻ�ǰ, CNSL : ����ǸŻ�ǰ
					obj("spdLst")(null)("pdTypCd") = "GNRL_GNRL"								'#��ǰ�����ڵ� | ����ǰ�� ����ǰ��� API�� ����Ѵ�. GNRL_GNRL : �Ϲ��Ǹ�_�Ϲݻ�ǰ, GNRL_ECPN : �Ϲ��Ǹ�e������ǰ, GNRL_GFTV : �Ϲ��Ǹ�_��ǰ��, GNRL_ZRWON : �Ϲ��Ǹ�_0����ǰ, CNSL_CNSL : ����Ǹ�_����ǰ
					obj("spdLst")(null)("gftvShpCd") = null										'o��ǰ���������ڵ尡 GNRL_GFTV(��ǰ��)�� ��쿡�� �ʼ� �Է�, ����ϻ�ǰ���� ��쿡�� e���� �׸��� �Է��Ͽ��� �Ѵ�. | PPR : ����, MBL : �����
					obj("spdLst")(null)("spdNm") = getItemNameFormat()							'#�Ǹ��ڻ�ǰ�� | �Էµ� �Ǹ��ڻ�ǰ���� ��ǰ�� ������ ���� ���û�ǰ������ ����ȴ�.
					obj("spdLst")(null)("brdNo") = getBrandCode()								'�귣���ȣ [�Ӽ���� ���� �׸�] | �Ӽ���� API�� ���Ͽ� ���ŵ� �귣���ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("mfcrNm") = CStr(FMakerName)							'������� | TXT ������ �Է��Ѵ�.
					obj("spdLst")(null)("oplcCd") = getOriginCode()								 '#�������ڵ� | ��Ÿ�� ��쿡�� "��ǰ�� ����"�ڵ�(ETC) �Է�
					obj("spdLst")(null)("mdlNo") = ""											'�𵨹�ȣ
					obj("spdLst")(null)("barCd") = ""											'���ڵ�
					obj("spdLst")(null)("tdfDvsCd") = CHKIIF(FVatInclude="N","02","01")			'#���������ڵ� [�����ڵ� : TDF_DVS_CD] | 01: ����, 02 : �鼼, 03 : ����, 04 : �ش����
					obj("spdLst")(null)("slStrtDttm") = FormatDate(now(), "00000000000000")		'#�ǸŽ����Ͻ� [YYYYMMDDHH24MISS ex) 20190801100000]
					obj("spdLst")(null)("slEndDttm") = "99991231235959"							'#�Ǹ������Ͻ� [YYYYMMDDHH24MISS ex) 20190801100000]
					Call getLotteonInfoCdParameter(obj)											'#��ǰǰ��������
					Call getLotteonCertInfoParameter(obj)										'�����������
					Call getLotteonStdCateAttrParameter(obj)									'ǥ��ī�װ��Ӽ����
					If FItemdiv = "06" Then
					Set obj("spdLst")(null)("itypOptLst") = jsArray()							'�Է����ɼǸ�� | �ִ� 5���� �Է����ɼ��� ������ �� �ִ�.
						Set obj("spdLst")(null)("itypOptLst")(null) = jsObject()
							obj("spdLst")(null)("itypOptLst")(null)("itypOptDvsCd") = "TXT"		'#�Է����ɼǱ����ڵ� [�����ڵ� : ITYP_OPT_DVS_CD] | NO : ����, TXT : �ؽ�Ʈ, DATE : �޷���, TIME : �ð�������
							obj("spdLst")(null)("itypOptLst")(null)("itypOptNm") = "�ؽ�Ʈ�� �Է��ϼ���"	'#�Է����ɼǸ�
					End If
					Set obj("spdLst")(null)("purPsbQtyInfo") = jsObject()						'���Ű��ɼ�������
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurYn") = "N"				'#��ǰ���ּұ��ſ��� [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurQty") = null			'o��ǰ���ּұ��ż��� | ��ǰ���ּұ��ſ��ΰ� Y�� ��� �ʼ��Է��Ѵ�.
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMaxPurPsbQtyYn") = "Y"		'#��ǰ���ִ뱸�Ű��ɼ������� [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("maxPurQty") = getOrderMaxNum		'o��ǰ���ִ뱸�ż��� | ��ǰ���ִ뱸�Ű��ɼ������ΰ� Y�� ��� �ʼ��Է��Ѵ�.
					obj("spdLst")(null)("ageLmtCd") = Chkiif(IsAdultItem()="Y", "19", "0")		'#���������ڵ� 0 : ������ ���Ű���, 15 : 15���̻� ���Ű���, 19 : 19���̻� ���Ű���
					obj("spdLst")(null)("prstPsbYn") = "N"										'�������ɿ��� [Y, N] | ����Ʈ:N
'					obj("spdLst")(null)("prstPckPsbYn") = ""									'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("prstMsgPsbYn") = ""									'--���ô� �ִµ� ������ ����..;
					obj("spdLst")(null)("prcCmprEpsrYn") = "Y"									'���ݺ񱳳��⿩�� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("bookCultCstDdctYn") = "N"								'������ȭ�� �������� [Y, N] | ����Ʈ:N �ŷ�ó�� ǥ��ī�װ��� ��� ������ȭ�� ������� �ش��ϴ� ��쿡�� �������ΰ� Y�̴�.
					obj("spdLst")(null)("isbnCd") = ""											'oISBN | ������ȭ�� �������ΰ� Y�̰� ī�װ��� �������� ī�װ��� ��� ISBN NO�� �Է��Ѵ�.
'					obj("spdLst")(null)("impCoNm") = ""											'���Ի�� | TXT �Է�
'					obj("spdLst")(null)("impDvsCd") = "NONE"									'���Ա����ڵ� [�����ڵ� : IMP_DVS_CD] | ���Ի���� �ִ� ��� �Է��Ѵ�. DRC_IMP : ������, PRL_IMP : �������, NONE : �ش����
					obj("spdLst")(null)("cshbltyPdYn") = "N"									'ȯ�ݼ���ǰ���� [Y, N] | ǥ��ī�װ� �Ӽ��� ��� �޴´�. ȯ�ݼ� ��ǰ���� �����Ǵ� ��� �ֹ����� �������ܿ� ���� ���Ű� ���ѵȴ�. ����Ʈ:N
'					obj("spdLst")(null)("dnDvPdYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("toysPdYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("intgSlPdNo") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("nmlPdYn") = ""											'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("prmmPdYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("otltPdYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("prmmInstPdYn") = ""									'--���ô� �ִµ� ������ ����..;
					obj("spdLst")(null)("brkHmapPkcpPsbYn") = "N"								'�������ſ��� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("ctrtTypCd") = "A"										'��������ڵ�[�����ڵ� : CTRT_TYP_CD] | A : �߰�, B : ��Ź
'					Set obj("spdLst")(null)("pdSzInfo") = jsObject()							'��ۻ��������� | ������ �Է� �����ϴ�.
'						obj("spdLst")(null)("pdSzInfo")("pdWdthSz") = ""						'��ǰ���λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdLnthSz") = ""						'��ǰ���λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdHghtSz") = ""						'��ǰ���̻����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckWdthSz") = ""						'���尡�λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckLnthSz") = ""						'���弼�λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckHghtSz") = ""						'������̻����� (cm)
					obj("spdLst")(null)("pdStatCd") = "NEW"										'#��ǰ�����ڵ� [�����ڵ� : PD_STAT_CD] | ��ǰ�����ڵ尡 ����ǰ(NEW)�� �ƴ� ��쿡�� ���������ڵ�� ���ϱ����ڵ带 USD�� �Ͽ� ��ǰ�����̹����� �ݵ�� ����Ͽ��� �Ѵ�.
					obj("spdLst")(null)("dpYn") = "Y"											'���ÿ��� [Y, N] | ����Ʈ:Y
'					obj("spdLst")(null)("ltonDpYn") = ""										'--���ô� �ִµ� ������ ����..;
					Call getLotteonKeywordsParameter(obj)										'�˻�Ű������ | 5�� ���ϸ� ��� ����
'					Set obj("spdLst")(null)("pdFileLst") = jsArray()							'o��ǰ���������ϸ�� | ��ǰ�����ڵ尡 ����ǰ(NEW)�� �ƴ� ��쿡�� ���������ڵ�� ���ϱ����ڵ带 USD�� �Ͽ� ��ǰ�����̹����� �ݵ�� ����Ͽ��� �Ѵ�.
'						Set obj("spdLst")(null)("pdFileLst")(null) = jsObject()
'							obj("spdLst")(null)("pdFileLst")(null)("fileTypCd") = ""			'#���������ڵ� [�����ڵ� : FILE_TYP_CD] | USD : ��ǰ����, TAG_LBL : Tag/�ɾ��, PD : ��ǰ
'							obj("spdLst")(null)("pdFileLst")(null)("fileDvsCd") = ""			'#���ϱ����ڵ� [�����ڵ� : FILE_DVS_CD] | USD : ��ǰ����, TAG_LBL : Tag/�ɾ��, 3D : ��ǰ3D�̹���, WDTH : ��ǰ������, VDO_FILE : ��ǰ������_FILE, VDO_URL : ��ǰ������_URL
'							obj("spdLst")(null)("pdFileLst")(null)("origFileNm") = ""			'#�������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.mp4
					Call getLotteonContParamToReg(obj)											'��ǰ������
					Set obj("spdLst")(null)("pyMnsExcpLst") = jsArray()							'�������ܿ��ܸ�� [�����ڵ� : PY_MNS_CD]
						obj("spdLst")(null)("pyMnsExcpLst") = null
					obj("spdLst")(null)("cnclPsbYn") = "Y"										'��Ұ��ɿ��� [Y, N] | ��� �Ұ��� ��ǰ�� ��쿡�� 'N'���� ���� ����Ʈ:Y
					obj("spdLst")(null)("immdCnclPsbYn") = "N"									'�����Ұ��ɿ��� [Y, N] | Ư�� ����(��� ��)������ ���� ���� �ٷ� ��� ������ ��� "Y"�� ���� ����Ʈ: Y
					obj("spdLst")(null)("dmstOvsDvDvsCd") = "DMST"								'�����ؿܹ�۱����ڵ� [�����ڵ� : DMST_OVS_DV_DVS_CD] | ����Ʈ:�������, DMST : �������, OVS : �ؿܹ߼�, RVRS_DPUR : ������
					obj("spdLst")(null)("pstkYn") = "N"											'������� [Y, N] ����Ʈ:N
					obj("spdLst")(null)("dvProcTypCd") = "LO_ENTP"								'#���ó�������ڵ� [�����ڵ� : DV_PROC_TYP_CD] | LO_CNTR : eĿ�ӽ� ���͹��, LO_ENTP : eĿ�ӽ� ��ü���
					obj("spdLst")(null)("dvPdTypCd") = dvPdTypCd								'#��ۻ�ǰ�����ڵ� [�����ڵ� : DV_PD_TYP_CD] | TDY_SND : ���ù߼�(0��), GNRL : �Ϲݻ�ǰ(3��), OD_MFG : �ֹ����ۻ�ǰ(15��), FREE_INST : ���ἳġ��ǰ(3��), CHRG_INST : ���ἳġ��ǰ(3��), PRMM_INST : �����̾���ġ��ǰ(365��), ECPN : e����(0��), GFTV : ��ǰ��(3��), OVS : �ؿܹ��(15��)
					obj("spdLst")(null)("sndBgtNday") = sndBgtNday								'�߼ۿ����ϼ� | ��ۻ�ǰ�����ڵ忡 ���� �ִ� �߼ۿ����ϼ��� �Է��Ѵ�.
					Set obj("spdLst")(null)("sndBgtDdInfo") = jsObject()						'�߼ۿ���������
						obj("spdLst")(null)("sndBgtDdInfo")("nldySndCloseTm") = "1500"			'#���� �߼۸����ð� [HH24MI ex) 1000]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndPsbYn") = "Y"				'#����� �߼۰��ɿ��� [Y, N]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndCloseTm") = "1300"			'o����� �߼۸����ð� [HH24MI ex) 1000] | ����� �߼� ���ɿ��� Y�� ��� �ʼ�
					obj("spdLst")(null)("dvRgsprGrpCd") = "GN101"								'#��۱ǿ��׷��ڵ� | ��۸���� ���Ͽ� �����Ǵ� �ڵ带 �Է��Ѵ�. | GN000(����), GN004(����), GN006(�����갣), GN101(����(�Ϻ���������), GN102(����(���ֵ� �� �������� ����), GN103(���� �� ������), GN104(���� + �ؿ�), GN105(����)
					obj("spdLst")(null)("dvMnsCd") = "DPCL"										'#��ۼ����ڵ� [�����ڵ� : DV_MNS_CD] �ܰǸ� �Է°��� | DGNN_DV : ������(�������), DPCL : �ù�, NONE_DV : �����, REG_MAIL : ���, ZIP : ����
					obj("spdLst")(null)("owhpNo") = DVPCd(1)									'#�������ȣ | �ŷ�ó API "(�Ϲ� Seller��) �Ǹ��� �����/��ǰ�� ���"�� ���Ͽ� ��ϵ� �������ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("hdcCd") = "0002"										'#�ù���ڵ� [�����ڵ� : DV_CO_CD] | 0002 : �������
					obj("spdLst")(null)("dvCstPolNo") = DVPCd(0)								'#��ۺ���å��ȣ | �ŷ�ó�� API�� ���� ����ϵ� ��ۺ���å��ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("adtnDvCstPolNo") = DVPCd(3)							'�߰���ۺ���å��ȣ | �ŷ�ó�� API�� ���� ����ϵ� �߰���ۺ���å��ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("cmbnDvPsbYn") = "Y"									'�չ�۰��ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("dvCstStdQty") = "0"									'��ۺ���ؼ��� | ����Ʈ:0
					obj("spdLst")(null)("qckDvUseYn") = "N"										'����ۻ�뿩�� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("crdayDvPsbYn") = "N"									'���Ϲ�۰��ɿ��� [Y, N] | ����Ʈ:N
'					Set obj("spdLst")(null)("crdayDvInfo") = jsObject()							'o���Ϲ������ | ���Ϲ�۰��ɿ��ΰ� Y�� ��� �ʼ���
'						obj("spdLst")(null)("crdayDvInfo")("odCloseTm") = ""					'#�ֹ������ð� [HH24MI ex) 1000] | ���Ϲ�۰��ɿ��ΰ� Y�� ��� �ʼ���
					obj("spdLst")(null)("spicUseYn") = "N"										'����Ʈ�Ȼ�뿩�� [Y, N] | ����Ʈ:N
					Set obj("spdLst")(null)("spicInfo") = jsObject()							'����Ʈ������ | ����Ʈ�Ȼ�뿩�� Y�� ��� �ʼ�
						obj("spdLst")(null)("spicInfo") = null
'					obj("spdLst")(null)("spicEusePdYn") = ""									'--���ô� �ִµ� ������ ����..;
					obj("spdLst")(null)("hpDdDvPsbYn") = "N"									'����Ϲ�۰��ɿ��� [Y, N] ����Ʈ:N
'					obj("spdLst")(null)("hpDdDvPsbPrd") = ""									'����Ϲ�۰��ɱⰣ | ����Ϲ�۰��ɿ��� Y�� ��� �ʼ�
					obj("spdLst")(null)("saveTypCd") = "NONE"									'���������ڵ� [�����ڵ� : SAVE_TYP_CD] | ����Ʈ:�ش���� RFRG : ����, FRZN : �õ�, FRSH : �ż�, NONE : �ش����
'					obj("spdLst")(null)("shopCnvMsgPsbYn") = ""									'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("rgnLmtPdYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("fprdDvPsbYn") = ""										'--���ô� �ִµ� ������ ����..;
'					obj("spdLst")(null)("spcfSqncPdYn") = ""									'--���ô� �ִµ� ������ ����..;
					obj("spdLst")(null)("rtngPsbYn") = "Y"										'��ǰ���ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("xchgPsbYn") = "Y"										'��ȯ���ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("echgPsbYn") = "N"										'�±�ȯ���ɿ��� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("cmbnRtngPsbYn") = "Y"									'�չ�ǰ���ɿ��� [Y, N] | �չ�۰��ɿ��ΰ� Y�� ��� Y, N ���� ����. N�� ��� N�� ���� ����
					obj("spdLst")(null)("rtngHdcCd") = ""										'��ǰ�ù���ڵ� | 0002 : �������
					obj("spdLst")(null)("rtngRtrvPsbYn") = "Y"									'��ǰȸ�����ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("rtrpNo") = DVPCd(2)									'#ȸ������ȣ | �ŷ�ó API "(�Ϲ� Seller��) �Ǹ��� �����/��ǰ�� ���"�� ���Ͽ� ��ϵ� ȸ������ȣ�� �Է��Ѵ�.
'					Set obj("spdLst")(null)("ecpnInfo") = jsObject()							'(����)e�������� | �ش� ��ǰ�� e������ ��쿡�� �Է��Ѵ�.
'					Set obj("spdLst")(null)("rntlPdInfo") = jsObject()							'(����)��Ż��ǰ���� | ��ǰ������ ��Ż�� ��� �ʼ���
'					Set obj("spdLst")(null)("opngPdInfo") = jsObject()							'(����)��������ǰ���� | ��ǰ���������ڵ尡 �Ϲ��Ǹ�_0����ǰ(GNRL_ZRWON)�� �ش��ϴ� ��������ǰ�� ��� �ʼ��Է��Ѵ�.
					obj("spdLst")(null)("stkMgtYn") = "Y"										'#���������� [Y, N] | 'N'�� ��� ��� 999,999,999�� ����. ����� �������� �ʴ´�.
					Call getLotteonOptionParameter(obj)											'��ǰ���
'					Set obj("spdLst")(null)("slrRcPdLst") = jsArray()							'������õ��ǰ��� | �ִ� 10������ ��� �����ϴ�.
'						Set obj("spdLst")(null)("slrRcPdLst")(null) = jsObject()
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSpdNo") = ""			'#������õ�Ǹ��ڻ�ǰ��ȣ
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSitmNo") = ""			'#������õ�Ǹ��ڴ�ǰ��ȣ
'							obj("spdLst")(null)("slrRcPdLst")(null)("epsrPrirRnkg") = ""		'#����켱����
		getLotteonItemRegParameter = obj.jsString
'   response.write getLotteonItemRegParameter
'   response.end
	End Function

	'��ǰ���� Json
	Public Function getLotteonItemEditParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj, tenBeasongDay
		tenBeasongDay = getShopLeadTime()

		If FItemdiv = "06" OR FItemdiv = "16" Then
			dvPdTypCd = "OD_MFG"
			sndBgtNday = "15"
		Else
			If tenBeasongDay > 3 Then
				dvPdTypCd = "OD_MFG"
				sndBgtNday = tenBeasongDay
			Else
				dvPdTypCd = "GNRL"
				sndBgtNday = "3"
			End If
		End If

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()														'��ϻ�ǰ���
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"										'#�ŷ�ó�׷��ڵ� | SR : �Ϲݼ���
					obj("spdLst")(null)("trNo") = afflTrCd										'#�ŷ�ó��ȣ
					obj("spdLst")(null)("lrtrNo") = ""											'�����ŷ�ó��ȣ
					obj("spdLst")(null)("scatNo") = FStd_cat_id									'#ǥ��ī�װ���ȣ
					Set obj("spdLst")(null)("dcatLst") = jsArray()								'#����ī�װ���� | �Ӽ������ API�� ���Ͽ� ǥ��ī�װ��� ���ε� ����ī�װ��� ������ �޴´�. ���ε� ����ī�װ� �߿��� �ϳ� �̻� �����Ͽ� �Է��Ѵ�.
						Set obj("spdLst")(null)("dcatLst")(null) = jsObject()
							obj("spdLst")(null)("dcatLst")(null)("mallCd") = "LTON"				'#�������ڵ� | LTON : �Ե�ON
							obj("spdLst")(null)("dcatLst")(null)("lfDcatNo") = FDisp_cat_id		'#leaf����ī�װ���ȣ
					obj("spdLst")(null)("spdNo") = ""&FLotteonGoodNo&""							'#�Ǹ��ڻ�ǰ��ȣ
					obj("spdLst")(null)("spdNm") = getItemNameFormat()							'�Ǹ��ڻ�ǰ�� | �Էµ� �Ǹ��ڻ�ǰ���� ��ǰ�� ������ ���� ���û�ǰ������ ����ȴ�.
					obj("spdLst")(null)("brdNo") = getBrandCode()								'�귣���ȣ [�Ӽ���� ���� �׸�] | �Ӽ���� API�� ���Ͽ� ���ŵ� �귣���ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("mfcrNm") = CStr(FMakerName)							'������� | TXT ������ �Է��Ѵ�.
					obj("spdLst")(null)("oplcCd") = getOriginCode()								 '�������ڵ� | ��Ÿ�� ��쿡�� "��ǰ�� ����"�ڵ�(ETC) �Է�
					obj("spdLst")(null)("mdlNo") = ""											'�𵨹�ȣ
					obj("spdLst")(null)("barCd") = ""											'���ڵ�
					obj("spdLst")(null)("tdfDvsCd") = CHKIIF(FVatInclude="N","02","01")			'#���������ڵ� [�����ڵ� : TDF_DVS_CD] | 01: ����, 02 : �鼼, 03 : ����, 04 : �ش����
					obj("spdLst")(null)("slStrtDttm") = FormatDate(now(), "00000000000000")		'#�ǸŽ����Ͻ� [YYYYMMDDHH24MISS ex) 20190801100000]
					obj("spdLst")(null)("slEndDttm") = "99991231235959"							'#�Ǹ������Ͻ� [YYYYMMDDHH24MISS ex) 20190801100000]
					Call getLotteonInfoCdParameter(obj)											'#��ǰǰ��������
					Call getLotteonCertInfoParameter(obj)										'�����������
					If FItemdiv = "06" Then
					Set obj("spdLst")(null)("itypOptLst") = jsArray()							'�Է����ɼǸ�� | �ִ� 5���� �Է����ɼ��� ������ �� �ִ�.
						Set obj("spdLst")(null)("itypOptLst")(null) = jsObject()
							obj("spdLst")(null)("itypOptLst")(null)("itypOptDvsCd") = "TXT"		'#�Է����ɼǱ����ڵ� [�����ڵ� : ITYP_OPT_DVS_CD] | NO : ����, TXT : �ؽ�Ʈ, DATE : �޷���, TIME : �ð�������
							obj("spdLst")(null)("itypOptLst")(null)("itypOptNm") = "�ؽ�Ʈ�� �Է��ϼ���"	'#�Է����ɼǸ�
					End If
					Set obj("spdLst")(null)("purPsbQtyInfo") = jsObject()						'���Ű��ɼ�������
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurYn") = "N"				'#��ǰ���ּұ��ſ��� [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurQty") = null			'o��ǰ���ּұ��ż��� | ��ǰ���ּұ��ſ��ΰ� Y�� ��� �ʼ��Է��Ѵ�.
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMaxPurPsbQtyYn") = "Y"		'#��ǰ���ִ뱸�Ű��ɼ������� [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("maxPurQty") = getOrderMaxNum		'o��ǰ���ִ뱸�ż��� | ��ǰ���ִ뱸�Ű��ɼ������ΰ� Y�� ��� �ʼ��Է��Ѵ�.
					obj("spdLst")(null)("prstPsbYn") = "N"										'�������ɿ��� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("prcCmprEpsrYn") = "Y"									'���ݺ񱳳��⿩�� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("bookCultCstDdctYn") = "N"								'������ȭ�� �������� [Y, N] | ����Ʈ:N �ŷ�ó�� ǥ��ī�װ��� ��� ������ȭ�� ������� �ش��ϴ� ��쿡�� �������ΰ� Y�̴�.
					obj("spdLst")(null)("isbnCd") = ""											'oISBN | ������ȭ�� �������ΰ� Y�̰� ī�װ��� �������� ī�װ��� ��� ISBN NO�� �Է��Ѵ�.
'					obj("spdLst")(null)("impCoNm") = ""											'���Ի�� | TXT �Է�
'					obj("spdLst")(null)("impDvsCd") = "NONE"									'���Ա����ڵ� [�����ڵ� : IMP_DVS_CD] | ���Ի���� �ִ� ��� �Է��Ѵ�. DRC_IMP : ������, PRL_IMP : �������, NONE : �ش����
					obj("spdLst")(null)("cshbltyPdYn") = "N"									'ȯ�ݼ���ǰ���� [Y, N] | ǥ��ī�װ� �Ӽ��� ��� �޴´�. ȯ�ݼ� ��ǰ���� �����Ǵ� ��� �ֹ����� �������ܿ� ���� ���Ű� ���ѵȴ�. ����Ʈ:N
					obj("spdLst")(null)("brkHmapPkcpPsbYn") = "N"								'�������ſ��� [Y, N] | ����Ʈ:N
'					Set obj("spdLst")(null)("pdSzInfo") = jsObject()							'��ۻ��������� | ������ �Է� �����ϴ�.
'						obj("spdLst")(null)("pdSzInfo")("pdWdthSz") = ""						'��ǰ���λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdLnthSz") = ""						'��ǰ���λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdHghtSz") = ""						'��ǰ���̻����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckWdthSz") = ""						'���尡�λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckLnthSz") = ""						'���弼�λ����� (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckHghtSz") = ""						'������̻����� (cm)
					obj("spdLst")(null)("dpYn") = "Y"											'���ÿ��� [Y, N] | ����Ʈ:Y
					Call getLotteonKeywordsParameter(obj)										'�˻�Ű������ | 5�� ���ϸ� ��� ����
'					Set obj("spdLst")(null)("pdFileLst") = jsArray()							'o��ǰ���������ϸ�� | ��ǰ�����ڵ尡 ����ǰ(NEW)�� �ƴ� ��쿡�� ���������ڵ�� ���ϱ����ڵ带 USD�� �Ͽ� ��ǰ�����̹����� �ݵ�� ����Ͽ��� �Ѵ�.
'						Set obj("spdLst")(null)("pdFileLst")(null) = jsObject()
'							obj("spdLst")(null)("pdFileLst")(null)("fileTypCd") = ""			'#���������ڵ� [�����ڵ� : FILE_TYP_CD] | USD : ��ǰ����, TAG_LBL : Tag/�ɾ��, PD : ��ǰ
'							obj("spdLst")(null)("pdFileLst")(null)("fileDvsCd") = ""			'#���ϱ����ڵ� [�����ڵ� : FILE_DVS_CD] | USD : ��ǰ����, TAG_LBL : Tag/�ɾ��, 3D : ��ǰ3D�̹���, WDTH : ��ǰ������, VDO_FILE : ��ǰ������_FILE, VDO_URL : ��ǰ������_URL
'							obj("spdLst")(null)("pdFileLst")(null)("origFileNm") = ""			'#�������ϸ�(��θ�) | ���ϸ��� ������ �ٿ�ε尡 ������ ��θ� �Է��Ѵ�. ex) http://abc.com/12/34/56/78_90.mp4
					Call getLotteonContParamToReg(obj)											'��ǰ������
					Set obj("spdLst")(null)("pyMnsExcpLst") = jsArray()							'�������ܿ��ܸ�� [�����ڵ� : PY_MNS_CD]
						obj("spdLst")(null)("pyMnsExcpLst") = null
					obj("spdLst")(null)("cnclPsbYn") = "Y"										'��Ұ��ɿ��� [Y, N] | ��� �Ұ��� ��ǰ�� ��쿡�� 'N'���� ���� ����Ʈ:Y
					obj("spdLst")(null)("immdCnclPsbYn") = "N"									'�����Ұ��ɿ��� [Y, N] | Ư�� ����(��� ��)������ ���� ���� �ٷ� ��� ������ ��� "Y"�� ���� ����Ʈ: Y
					obj("spdLst")(null)("dvPdTypCd") = dvPdTypCd								'#��ۻ�ǰ�����ڵ� [�����ڵ� : DV_PD_TYP_CD] | TDY_SND : ���ù߼�(0��), GNRL : �Ϲݻ�ǰ(3��), OD_MFG : �ֹ����ۻ�ǰ(15��), FREE_INST : ���ἳġ��ǰ(3��), CHRG_INST : ���ἳġ��ǰ(3��), PRMM_INST : �����̾���ġ��ǰ(365��), ECPN : e����(0��), GFTV : ��ǰ��(3��), OVS : �ؿܹ��(15��)
					obj("spdLst")(null)("sndBgtNday") = sndBgtNday								'�߼ۿ����ϼ� | ��ۻ�ǰ�����ڵ忡 ���� �ִ� �߼ۿ����ϼ��� �Է��Ѵ�.
					Set obj("spdLst")(null)("sndBgtDdInfo") = jsObject()						'�߼ۿ���������
						obj("spdLst")(null)("sndBgtDdInfo")("nldySndCloseTm") = "1500"			'#���� �߼۸����ð� [HH24MI ex) 1000]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndPsbYn") = "Y"				'#����� �߼۰��ɿ��� [Y, N]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndCloseTm") = "1300"			'o����� �߼۸����ð� [HH24MI ex) 1000] | ����� �߼� ���ɿ��� Y�� ��� �ʼ�
					obj("spdLst")(null)("dvRgsprGrpCd") = "GN101"								'��۱ǿ��׷��ڵ� | ��۸���� ���Ͽ� �����Ǵ� �ڵ带 �Է��Ѵ�. | GN000(����), GN004(����), GN006(�����갣), GN101(����(�Ϻ���������), GN102(����(���ֵ� �� �������� ����), GN103(���� �� ������), GN104(���� + �ؿ�), GN105(����)
					obj("spdLst")(null)("dvMnsCd") = "DPCL"										'#��ۼ����ڵ� [�����ڵ� : DV_MNS_CD] �ܰǸ� �Է°��� | DGNN_DV : ������(�������), DPCL : �ù�, NONE_DV : �����, REG_MAIL : ���, ZIP : ����
					obj("spdLst")(null)("owhpNo") = DVPCd(1)									'#�������ȣ | �ŷ�ó API "(�Ϲ� Seller��) �Ǹ��� �����/��ǰ�� ���"�� ���Ͽ� ��ϵ� �������ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("hdcCd") = "0002"										'#�ù���ڵ� [�����ڵ� : DV_CO_CD] | 0002 : �������
					obj("spdLst")(null)("dvCstPolNo") = DVPCd(0)								'#��ۺ���å��ȣ | �ŷ�ó�� API�� ���� ����ϵ� ��ۺ���å��ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("adtnDvCstPolNo") = DVPCd(3)							'�߰���ۺ���å��ȣ | �ŷ�ó�� API�� ���� ����ϵ� �߰���ۺ���å��ȣ�� �Է��Ѵ�.
					obj("spdLst")(null)("cmbnDvPsbYn") = "Y"									'�չ�۰��ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("dvCstStdQty") = "0"									'��ۺ���ؼ��� | ����Ʈ:0
					obj("spdLst")(null)("qckDvUseYn") = "N"										'����ۻ�뿩�� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("crdayDvPsbYn") = "N"									'���Ϲ�۰��ɿ��� [Y, N] | ����Ʈ:N
'					Set obj("spdLst")(null)("crdayDvInfo") = jsObject()							'o���Ϲ������ | ���Ϲ�۰��ɿ��ΰ� Y�� ��� �ʼ���
'						obj("spdLst")(null)("crdayDvInfo")("odCloseTm") = ""					'#�ֹ������ð� [HH24MI ex) 1000] | ���Ϲ�۰��ɿ��ΰ� Y�� ��� �ʼ���
					obj("spdLst")(null)("spicUseYn") = "N"										'����Ʈ�Ȼ�뿩�� [Y, N] | ����Ʈ:N
					Set obj("spdLst")(null)("spicInfo") = jsObject()							'����Ʈ������ | ����Ʈ�Ȼ�뿩�� Y�� ��� �ʼ�
						obj("spdLst")(null)("spicInfo") = null
					obj("spdLst")(null)("hpDdDvPsbYn") = "N"									'����Ϲ�۰��ɿ��� [Y, N] ����Ʈ:N
'					obj("spdLst")(null)("hpDdDvPsbPrd") = ""									'����Ϲ�۰��ɱⰣ | ����Ϲ�۰��ɿ��� Y�� ��� �ʼ�
					obj("spdLst")(null)("saveTypCd") = "NONE"									'���������ڵ� [�����ڵ� : SAVE_TYP_CD] | ����Ʈ:�ش���� RFRG : ����, FRZN : �õ�, FRSH : �ż�, NONE : �ش����
					obj("spdLst")(null)("rtngPsbYn") = "Y"										'��ǰ���ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("xchgPsbYn") = "Y"										'��ȯ���ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("echgPsbYn") = "N"										'�±�ȯ���ɿ��� [Y, N] | ����Ʈ:N
					obj("spdLst")(null)("cmbnRtngPsbYn") = "Y"									'�չ�ǰ���ɿ��� [Y, N] | �չ�۰��ɿ��ΰ� Y�� ��� Y, N ���� ����. N�� ��� N�� ���� ����
					obj("spdLst")(null)("rtngHdcCd") = ""										'��ǰ�ù���ڵ� | 0002 : �������
					obj("spdLst")(null)("rtngRtrvPsbYn") = "Y"									'��ǰȸ�����ɿ��� [Y, N] | ����Ʈ:Y
					obj("spdLst")(null)("rtrpNo") = DVPCd(2)									'ȸ������ȣ | �ŷ�ó API "(�Ϲ� Seller��) �Ǹ��� �����/��ǰ�� ���"�� ���Ͽ� ��ϵ� ȸ������ȣ�� �Է��Ѵ�.
'					Set obj("spdLst")(null)("ecpnInfo") = jsObject()							'(����)e�������� | �ش� ��ǰ�� e������ ��쿡�� �Է��Ѵ�.
'					Set obj("spdLst")(null)("rntlPdInfo") = jsObject()							'(����)��Ż��ǰ���� | ��ǰ������ ��Ż�� ��� �ʼ���
'					Set obj("spdLst")(null)("opngPdInfo") = jsObject()							'(����)��������ǰ���� | ��ǰ���������ڵ尡 �Ϲ��Ǹ�_0����ǰ(GNRL_ZRWON)�� �ش��ϴ� ��������ǰ�� ��� �ʼ��Է��Ѵ�.
					obj("spdLst")(null)("stkMgtYn") = "Y"										'#���������� [Y, N] | 'N'�� ��� ��� 999,999,999�� ����. ����� �������� �ʴ´�.
					Call getLotteonOptionEditParameter(obj)										'��ǰ���
'					Set obj("spdLst")(null)("slrRcPdLst") = jsArray()							'������õ��ǰ��� | �ִ� 10������ ��� �����ϴ�.
'						Set obj("spdLst")(null)("slrRcPdLst")(null) = jsObject()
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSpdNo") = ""			'#������õ�Ǹ��ڻ�ǰ��ȣ
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSitmNo") = ""			'#������õ�Ǹ��ڴ�ǰ��ȣ
'							obj("spdLst")(null)("slrRcPdLst")(null)("epsrPrirRnkg") = ""		'#����켱����
		getLotteonItemEditParameter = obj.jsString
'    response.write obj.jsString
'    response.end
	End Function

	'��ǰ ����ȸ Json
	Public Function getLotteonItemViewParameter
		Dim strRst
		Dim obj
		Set obj = jsObject()
			obj("trGrpCd") = "SR"
			obj("trNo") = afflTrCd
			obj("spdNo") = FLotteonGoodNo
		getLotteonItemViewParameter = obj.jsString
	End Function

	'��ǰ ������ Json
	Public Function getLotteonQuantityParameter
		Dim strRst
		Dim obj, sqlStr, arrRows, limitsu

		Set obj = jsObject()
			Set obj("itmStkLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, o.optlimitno, o.optlimitsold "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " WHERE r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "���ϻ�ǰ") Then
				Set obj("itmStkLst")(null) = jsObject()
					obj("itmStkLst")(null)("trGrpCd") = "SR"
					obj("itmStkLst")(null)("trNo") = afflTrCd
					obj("itmStkLst")(null)("spdNo") = FLotteonGoodNo
					obj("itmStkLst")(null)("sitmNo") = arrRows(1,0)
					obj("itmStkLst")(null)("stkQty") = getLimitEa()
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						limitsu = ""
						If FLimityn = "Y" Then
							If arrRows(3, i) - arrRows(4, i) - 5 < 1 Then
								limitsu = 0
							Else
								limitsu = arrRows(3, i) - arrRows(4, i) - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If

						Set obj("itmStkLst")(i) = jsObject()
							obj("itmStkLst")(i)("trGrpCd") = "SR"
							obj("itmStkLst")(i)("trNo") = afflTrCd
							obj("itmStkLst")(i)("spdNo") = FLotteonGoodNo
							obj("itmStkLst")(i)("sitmNo") = arrRows(1, i)
							obj("itmStkLst")(i)("stkQty") = limitsu
					Next
				End If
			End If
		getLotteonQuantityParameter = obj.jsString
	End Function

	'��ǰ ���ݼ��� Json
	Public Function getLotteonPriceParameter
		Dim strRst
		Dim obj, sqlStr, arrRows
		Dim vMustprice
		vMustprice = mustPrice()

		Set obj = jsObject()
			Set obj("itmPrcLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, isnull(o.optAddPrice, 0) optAddPrice "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "���ϻ�ǰ") Then
				Set obj("itmPrcLst")(null) = jsObject()
					obj("itmPrcLst")(null)("trGrpCd") = "SR"
					obj("itmPrcLst")(null)("trNo") = afflTrCd
					obj("itmPrcLst")(null)("spdNo") = FLotteonGoodNo
					obj("itmPrcLst")(null)("sitmNo") = arrRows(1,0)
					obj("itmPrcLst")(null)("slPrc") = vMustprice
					obj("itmPrcLst")(null)("hstStrtDttm") = FormatDate(now(), "00000000000000")		'#���ݽ����Ͻ�
					obj("itmPrcLst")(null)("hstEndDttm") = "99991231235959"							'#���������Ͻ�
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						Set obj("itmPrcLst")(i) = jsObject()
							obj("itmPrcLst")(i)("trGrpCd") = "SR"
							obj("itmPrcLst")(i)("trNo") = afflTrCd
							obj("itmPrcLst")(i)("spdNo") = FLotteonGoodNo
							obj("itmPrcLst")(i)("sitmNo") = arrRows(1, i)
							obj("itmPrcLst")(i)("slPrc") = vMustprice + arrRows(3, i)						'#�ǸŰ�
							obj("itmPrcLst")(i)("hstStrtDttm") = FormatDate(now(), "00000000000000")		'#���ݽ����Ͻ�
							obj("itmPrcLst")(i)("hstEndDttm") = "99991231235959"							'#���������Ͻ�
					Next
				End If
			End If
		getLotteonPriceParameter = obj.jsString
	End Function

	'��ǰ �ǸŻ��� ���� Json
	Public Function getLotteonSellynParameter(ichgSellYn)
		Dim strRst
		Dim obj, slStatCd
		Select Case ichgSellYn
			Case "Y"	slStatCd = "SALE"		'�Ǹ���
			Case "N"	slStatCd = "SOUT"		'ǰ��
			Case "X"	slStatCd = "END"		'�Ǹ�����
		End Select

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"
					obj("spdLst")(null)("trNo") = afflTrCd
					obj("spdLst")(null)("spdNo") = FLotteonGoodNo
					obj("spdLst")(null)("slStatCd") = slStatCd
		getLotteonSellynParameter = obj.jsString
	End Function

	'��ǰ �ǸŻ��� ���� Json
	Public Function getLotteonOptStatusParameter()
		Dim strRst
		Dim obj, sqlStr, arrRows, optsellyn

		If rsget.state = "1" Then
			rsget.close
		End If

		Set obj = jsObject()
			Set obj("sitmLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, isnull(o.optAddPrice, 0) optAddPrice "
			sqlStr = sqlStr & " , isnull(o.optionname, '') as optionname, isnull(o.isUsing, '') as isUsing, isnull(o.optsellyn, '') as optsellyn "
			sqlStr = sqlStr & " , (o.optlimitno - o.optlimitsold - 5) as optLimit "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "���ϻ�ǰ") Then
				optsellyn = Chkiif(FMaySoldOut="Y", "SOUT", "SALE")
				Set obj("sitmLst")(null) = jsObject()
					obj("sitmLst")(null)("trGrpCd") = "SR"
					obj("sitmLst")(null)("trNo") = afflTrCd
					obj("sitmLst")(null)("spdNo") = FLotteonGoodNo
					obj("sitmLst")(null)("sitmNo") = arrRows(1,0)
					obj("sitmLst")(null)("slStatCd") = optsellyn
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						optsellyn = ""
						'itempoption�� ���ٸ�(�ɼǻ���) / itemopti
						If (arrRows(0, i) = "") OR (arrRows(5, i) <> "Y") OR (arrRows(6, i) <> "Y") THEN
							optsellyn = "SOUT"
						ElseIf FLimityn = "Y" AND (arrRows(7, i) < 1) Then
							optsellyn = "SOUT"
						Else
							optsellyn = "SALE"
						End If

						Set obj("sitmLst")(i) = jsObject()
							obj("sitmLst")(i)("trGrpCd") = "SR"
							obj("sitmLst")(i)("trNo") = afflTrCd
							obj("sitmLst")(i)("spdNo") = FLotteonGoodNo
							obj("sitmLst")(i)("sitmNo") = arrRows(1, i)
							obj("sitmLst")(i)("slStatCd") = optsellyn
					Next
				End If
			End If
		getLotteonOptStatusParameter = obj.jsString
	End Function

End Class

Class CLotteon
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
	Public Sub getLotteonNotRegOneItem
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
		strSql = strSql & " , (SELECT db_etcmall.dbo.getOutmallKeywords ('"& CMALLNAME &"', i.itemid) ) as keywords "
		strSql = strSql & "	, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, isNULL(R.lotteonStatCD,-9) as lotteonStatCD, IsNull(R.lotteonPrice, 0) as lotteonPrice "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & " , am.std_cat_id, am.disp_cat_id "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_lotteon_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
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
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "				'�ö��/ȭ�����/�ؿ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_lotteon_regItem WHERE lotteonStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.	'lotteon��ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteonItem
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
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FLotteonStatCD		= rsget("lotteonStatCD")
				FOneItem.FLotteonPrice		= rsget("lotteonPrice")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FStd_cat_id		= rsget("std_cat_id")
				FOneItem.FDisp_cat_id		= rsget("disp_cat_id")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getLotteonNotEditOneItem
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
		strSql = strSql & "	, m.lotteonGoodNo, m.lotteonprice, m.lotteonSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, isnull(m.lastStatCheckDate, '1900-01-01 00:00:00.000') as lastStatCheckDate "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.std_cat_id, am.disp_cat_id, isNULL(m.lotteonStatCD,-9) as lotteonStatCD "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.lotteonGoodno is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteonItem
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
				FOneItem.FLotteonGoodNo		= rsget("lotteonGoodNo")
				FOneItem.FLotteonprice		= rsget("lotteonprice")
				FOneItem.FLotteonSellYn		= rsget("lotteonSellYn")

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
                FOneItem.FStd_cat_id		= rsget("std_cat_id")
				FOneItem.FDisp_cat_id		= rsget("disp_cat_id")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FLotteonStatCD		= rsget("lotteonStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FLastStatCheckDate = rsget("lastStatCheckDate")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

End Class

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

function APIURL()
	If application("Svr_Info") = "Dev" Then
		APIURL = "https://dev-openapi.lotteon.com"
	Else
		APIURL = "https://openapi.lotteon.com"
	End If
end function

function APIkey()
	If application("Svr_Info") = "Dev" Then
		APIkey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
	Else
		APIkey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
	End If
end function

function afflTrCd()
	If application("Svr_Info") = "Dev" Then
		afflTrCd = "LO10001101"
	Else
		afflTrCd = "LD304013"
	End If
end function

function DVPCd(v)
	'v : 0(��ۺ� ��å), 1(�����), 2(ȸ����), 3(�߰���ۺ�)
	If v = "0" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1000529"
		Else
			DVPCd = "DLD706463"
		End If
	ElseIf v = "1" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1300153"
		Else
			DVPCd = "BPLD304013"
		End If
	ElseIf v = "2" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1300153"
		Else
			DVPCd = "PLD333127"
		End If
	ElseIf v = "3" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "2009166"
		Else
			DVPCd = "2009166"
		End If
	End If
end function
%>