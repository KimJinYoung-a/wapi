<%
CONST CMAXMARGIN = 15			'' MaxMagin��..
CONST CMALLNAME = "gsshop"
CONST CMAXLIMITSELL = 5			'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CGSSHOPMARGIN = 12
CONST CUPJODLVVALID = True		''��ü ���ǹ�� ��� ���ɿ���
CONST COurCompanyCode = 1003890	'' ���»��ڵ�
CONST COurRedId = "TBT"

Class CGSShopItem
	Public FItemid
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
	Public ForderComment
	Public FoptionCnt
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FGsshopGoodNo
	Public FGsshopprice
	Public FGsshopSellYn
	Public FbasicimageNm
	Public FregImageName

	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public FIsNulltoTimeout

	Public FUserid
	Public FSocname
	Public FSocname_kor
	Public FDeliver_name
	Public FReturn_zipcode
	Public FReturn_address
	Public FReturn_address2
	Public FMaeipdiv
	Public FDeliveryCd
	Public FDeliveryAddrCd
	Public FBrandcd
	Public FDivname

	Public FIcnt
	Public FDivcode
	Public Fcdd_Name
	Public Fcdl_Name
	Public Fcdm_Name
	Public Fcds_Name

	Public FSafecode
	Public FSafecode_NAME
	Public FIsvat
	Public FIsvat_NAME
	Public FInfodiv1
	Public FInfodiv2
	Public FInfodiv3
	Public FInfodiv4
	Public FInfodiv5
	Public FInfodiv6


	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FDispNo
	Public FDispNm
	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public Fdisptpcd
	Public FCateIsUsing
	Public FD_NAME

	Public FDispThnNm

	Public FItemoption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold

	Public Fvatinclude
	Public FGSShopStatCd
	Public FOptNotMatch
	Public FAdultType

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
		If (Fdisptpcd="B") Then
			getDisptpcdName = "<font color='blue'>����</font>"
		Elseif (Fdisptpcd = "D") Then
			getDisptpcdName = "�Ϲ�"
		Else
			getDisptpcdName = Fdisptpcd
		End if
	End Function


	'// ǰ������
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
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

	'�ɼ� �ǸŻ��� ����
	Public Function isOptNotMatch()
		Dim strSql, arrRows, isOptionExists, tmpCnt
		Dim bufcnt, i, optLimit, optlimityn, isUsing, optsellyn, optNameDiff, forceExpired
		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_gsshop '"&CMALLNAME&"'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		tmpCnt = 0
		isOptNotMatch = "N"
		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				optNameDiff			= (ArrRows(12,i)=1)
				forceExpired		= (ArrRows(13,i)=1)
				If ((forceExpired) or (optNameDiff) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5)) Then
					tmpCnt = tmpCnt + 1
				End If
			Next

			If FOptionCnt = 1 AND tmpCnt = 1 AND i = 1 Then
				isOptNotMatch = "Y"
			ElseIf (FOptionCnt >= 1) AND (i = tmpCnt) Then
				isOptNotMatch = "Y"
			End If
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, vBigPrice, vSmallPrice, ownItemCnt
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
			If FGsshopprice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					If (FSellCash < Round(FGsshopprice * 0.25, 0)) Then
						MustPrice = CStr(GetRaiseValue(Round(FGsshopprice * 0.25, 0)/10)*10)
					Else
						MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
	End Function

	public function GetGSLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetGSLmtQty = 0
			Else
				GetGSLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetGSLmtQty = 999
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
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

	Public Function getOptionLimitNo()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (IsOptionSoldOut) Then
			getOptionLimitNo = 0
		Else
			If (Foptlimityn = "Y") Then
				If (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO) Then
					getOptionLimitNo = 0
				Else
					getOptionLimitNo = Foptlimitno - Foptlimitsold - CLIMIT_SOLDOUT_NO
				End If
			Else
				getOptionLimitNo = 999
			End if
		End If
	End Function

	'// GSShop �Ǹſ��� ��ȯ
	Public Function getGSShopSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getGSShopSellYn = "Y"
			Else
				getGSShopSellYn = "N"
			End If
		Else
			getGSShopSellYn = "N"
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

	Function getItemNameFormat()
		Dim buf
'		buf = "[�ٹ�����]"&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(FItemName,"'","")						'���� ��ǰ�� �տ� [�ٹ�����] ����

		If Left(FItemName, Len(Trim(FSocname_kor)) + 2) = "[" & FSocname_kor & "]" Then
		ElseIf (Left(FItemName, len(FSocname_kor)) <> FSocname_kor) Then
			buf = FSocname_kor & " " & Replace(FItemName,"'","")		'[�ٹ�����] ���� ���� / �귣���ѱ۸� ���� / 2020-07-30 ���� ����
		End If

		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"&","��")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"+","%2B")
		buf = replace(buf,":","%3A")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
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

	'��ǰ�з��� ��������
	Public Function getGSShopItemSafeInfoParam()
		Dim buf, strSql, regCertCnt, regSafetydiv
		Dim safeCertGbnCd, safeCertOrgCd, safeCertModelNm, safeCertNo, safeCertDt
		If FDivcode = "" Then			'��ǰ�з��� �������� ī�װ�
			rw "��ǰ�з��� �������ּ���"
			Exit Function
			response.end
		End If

		buf = ""
		If (FSafecode = "3") Then		'SafeCode�� 3(����)�̶��..
			buf = buf & "&safeCertGbnCd=0"		'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ, 4 : �����ǰ��������Ȯ��
			buf = buf & "&safeCertOrgCd=0"		'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			buf = buf & "&safeCertModelNm="		'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			buf = buf & "&safeCertNo="			'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			buf = buf & "&safeCertDt="			'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
'			buf = buf & "&safeCertFileNm="		'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
		Else							'SafeCode�� 1(�ʼ�,����)�̶��..
			If (Fsafetyyn) = "Y" AND (FSafecode = "1" OR FSafecode = "2") Then			'SafeCode�� 1(�ʼ�,����)�̰� �ٹ����ٿ� �����������ΰ� Y���
				strSql = ""
				strSql = strSql & " SELECT COUNT(*) as cnt, safetydiv FROM db_item.dbo.tbl_safetycert_tenReg WHERE itemid = " &Fitemid& " GROUP BY safetydiv "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.EOF Then
					regCertCnt = rsget("cnt")
					regSafetydiv = rsget("safetydiv")
				End If
				rsget.Close

				If regCertCnt > 0 AND (regSafetydiv = "30" OR regSafetydiv = "60" OR regSafetydiv = "90") Then
					If regSafetydiv = "30" Then
						safeCertGbnCd = "7"
						safeCertOrgCd = "701"
					ElseIf regSafetydiv = "60" Then
						safeCertGbnCd = "8"
						safeCertOrgCd = "801"
					ElseIf regSafetydiv = "90" Then
						safeCertGbnCd = "C"
						safeCertOrgCd = "C01"
					End If
					buf = buf & "&safeCertGbnCd="&safeCertGbnCd								'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ, 4 : �����ǰ��������Ȯ��
					buf = buf & "&safeCertOrgCd="&safeCertOrgCd								'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertModelNm="											'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertNo="												'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertDt=" 												'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
	'				buf = buf & "&safeCertFileNm=Y"											'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				ElseIf regCertCnt > 0 AND (regSafetydiv <> "30" AND regSafetydiv <> "60" AND regSafetydiv <> "90") Then
					strSql = ""
					strSql = strSql & " EXEC [db_item].[dbo].[usp_API_GSShop_SafeInfo_Get] " & FItemid
					rsget.CursorLocation = adUseClient
					rsget.CursorType=adOpenStatic
					rsget.Locktype=adLockReadOnly
					rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
						Do Until rsget.EOF
							buf = buf & "&safeCertGbnCd=" & rsget("safeCertGbnCd")			'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ, 4 : �����ǰ��������Ȯ��
							buf = buf & "&safeCertOrgCd=" & rsget("safeCertOrgCd")			'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
							buf = buf & "&safeCertModelNm=" & rsget("safeCertModelNm")		'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
							buf = buf & "&safeCertNo=" & rsget("safeCertNo")				'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
							buf = buf & "&safeCertDt=" & rsget("safeCertDt") 				'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			'				buf = buf & "&safeCertFileNm=Y"									'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
							rsget.MoveNext
						Loop
					End If
					rsget.Close
				Else
					buf = buf & "&safeCertGbnCd=0"		'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ
					buf = buf & "&safeCertOrgCd=0"		'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertModelNm="		'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertNo="			'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
					buf = buf & "&safeCertDt="			'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
	'				buf = buf & "&safeCertFileNm="		'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				End If
			Else						'�� ���� ���� ���� �ش���� ó��
				buf = buf & "&safeCertGbnCd=0"		'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ
				buf = buf & "&safeCertOrgCd=0"		'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertModelNm="		'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertNo="			'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertDt="			'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
'				buf = buf & "&safeCertFileNm="		'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			End If
		End If
		getGSShopItemSafeInfoParam = buf
	End Function

	Public Function getGSCateParam()
		Dim strSql, bufcnt, cateKey, buf, cateGbn, isDefaultCate
		buf = ""
		strSql = ""
		strSql = strSql & " SELECT TOP 2 c.CateKey, c.cateGbn "
		strSql = strSql & " FROM db_item.dbo.tbl_gsshop_cate_mapping as m "
		strSql = strSql & " JOIN db_temp.dbo.tbl_gsshop_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : �귣�� / D : �Ϲ�
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
'rw strSql
'response.end
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				cateGbn = "S"			'���ظ��� : S / ��ȭ������ : D / ��Ʈ�ʽ����� : P / BP���� : B
				isDefaultCate = "N"
				If rsget("cateGbn") = "B" Then
					cateGbn = "P"
					isDefaultCate = "Y"
				End If

			    cateKey  = rsget("CateKey")
				buf = buf & "&prdSectListSectid="&cateKey
				buf = buf & "&prdSectListSectGbn="&cateGbn
				buf = buf & "&prdSectListSectStdYn="&isDefaultCate
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSCateParam = bufcnt&"|_|"&buf
	End Function

	'���»�������/�� | �⺻�� : �ǸŰ�*(1-0.13) // ����12��
    Function getGSShopSuplyPrice()
		'getGSShopSuplyPrice = CLNG(FSellCash * (100-CGSSHOPMARGIN) / 100)
		getGSShopSuplyPrice = CLNG(MustPrice * (100-CGSSHOPMARGIN) / 100)
    End Function

	'��ǰ �з��� MDID ����
	Public Function getMdIdMapping(divCode)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT TOP 1 mdid "
		strSql = strSql & " FROM db_item.[dbo].[tbl_gsshop_mdid_mapping]  "
		strSql = strSql & " WHERE divcode = '"& divCode &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			getMdIdMapping = rsget("mdid")
		Else 
			getMdIdMapping = "80055"
		End If
		rsget.Close
	End Function

   ''�ֹ����� ����
    Public Function getzCostomMadeInd()
		Dim ordMnfcYn, ordMnfcTypCd, ordMnfcTermDdcnt, ordMnfcCntnt
		Dim buf
		If (Fitemdiv="06" or Fitemdiv="16" or FtenCateLarge="040") Then
			If Fitemdiv = "06" Then
				ordMnfcTypCd = "10"
				ordMnfcCntnt = "�ֹ����ۿ�û����"
			ElseIf Fitemdiv="16" OR FtenCateLarge="040" Then
				ordMnfcTypCd = "20"
			End If

			If FtenCateLarge="040" Then
				ordMnfcTermDdcnt = 15
			ElseIf (FrequireMakeDay > 5) Then
				ordMnfcTermDdcnt = FrequireMakeDay
			ElseIf (FrequireMakeDay < 1) Then
				ordMnfcTermDdcnt = 5
			Else
				ordMnfcTermDdcnt = FrequireMakeDay + 1
			End If
			ordMnfcYn = "Y"
		Else
			ordMnfcYn = "N"
		End If

		buf = ""
		buf = buf & "&ordMnfcYn="&ordMnfcYn					'(*)�ֹ����ۿ���
		buf = buf & "&ordMnfcTypCd="&ordMnfcTypCd			'(*)�ֹ����������ڵ� | �ֹ����ۿ��ΰ� 'Y'�� ��� �ʼ��Է��׸��Դϴ�.('N'�� ���� NULL) NULL : �ش����, 10 : ��������, 20 : �ֹ�������, 30 : �ֹ��ļ���
		buf = buf & "&ordMnfcCntnt="&ordMnfcCntnt			'(*)�ֹ����۳��� | �ֹ����������� 10�� ���������� ��� �ʼ��Է��׸��Դϴ�.
		buf = buf & "&ordMnfcTermDdcnt="&ordMnfcTermDdcnt	'(*)�ֹ����۱Ⱓ�ϼ� | �ֹ����ۿ��ΰ� 'Y'�� ��� �ʼ��Է��׸��Դϴ�.('N'�� ���� NULL)
		getzCostomMadeInd = buf
    End Function

	'//New ��ǰ��� �Ķ���� ����
	Public Function getGSShopItemNewRegParameter(v)
		Dim strRst
		Dim DeliverCd, DeliverAddrCd, brandcd
		'################################ �ù��/��ǰ�� ���� Ȯ�� #################################
'2017-04-24 ���� ����..�ٹ�� ����� CJ�� ����� ������..���� : ���ε��� ���� �� ��������� �� ��..
'		If (Fdeliverytype = "9") OR (Fdeliverytype = "7") OR (Fdeliverytype = "2") Then	'��ü����̶��
'			DeliverCd		= FDeliveryCd
'			DeliverAddrCd	= FDeliveryAddrCd
'			DeliverCd = "CJ"															'CJ�ù�
'			DeliverAddrCd = "0001"														'0001�� ��� ���� �Ϸ�(������ ����)
'		Else																			'�ٹ���
'			DeliverCd = "CJ"															'CJ�ù�
'			DeliverAddrCd = "0001"														'0001�� ��� ���� �Ϸ�(������ ����)
'		End If

		DeliverCd = "HJ"															'�����ù�
		DeliverAddrCd = "0001"														'0001�� ��� ���� �Ϸ�(������ ����)
		brandcd = "115985"
		'##########################################################################################

		'################################ �̹��� ����Ʈ ���� ȣ�� #################################
		Dim CallImage, CntImage, NmImage
		CallImage = getGSShopAddImageParam()
		CntImage = Split(CallImage, "|_|")(0)
		NmImage = Split(CallImage, "|_|")(1)
		'##########################################################################################

		'################################ �Ӽ�(�ɼ�) �׸� ���� ȣ�� ###############################
		Dim CallOpt, COptyn, CntOpt, NmOpt
		CallOpt = getGSShopOptionParam()
		COptyn = Split(CallOpt, "|_|")(0)
		CntOpt = Split(CallOpt, "|_|")(1)
		NmOpt = Split(CallOpt, "|_|")(2)
		'##########################################################################################

		'################################ �������� �׸� ���� ȣ�� #################################
		Dim CallCate, CntCate, NmCate
		CallCate = getGSCateParam()
		CntCate = Split(CallCate, "|_|")(0)
		NmCate = Split(CallCate, "|_|")(1)
		'##########################################################################################

		'################################ ���� ��� �׸� ���� ȣ�� ################################
		Dim CallInfoCd, CntInfoCd, NmInfoCd
		CallInfoCd = getGSShopItemInfoCdParam()
		CntInfoCd = Split(CallInfoCd, "|_|")(0)
		NmInfoCd = Split(CallInfoCd, "|_|")(1)
		'##########################################################################################
		'���� ���� �� �ݺ�����Ʈ �Ǽ�
		strRst = ""
		strRst = strRst & "regGbn=I"														'(*)��ϱ��� | I : �ű�, U : ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����	| �ش� ���»縦 �ĺ��Ҽ� �ִ� �����빮�� 3��(�� : TBT)�� ����
		strRst = strRst & "&regSubjCd=SUP"													'(*)�����ü�ڵ� | ���� ������ ��� : MD, ���»簡 ������ ��� : SUP
		strRst = strRst & "&prdCntntListCnt="&CntImage										'(*)�̹�������Ʈ�Ǽ� | ��ǰ�̹�������Ʈ (prdCntntList) �ݺ�Ƚ���� �����մϴ�.
'		strRst = strRst & "&prdDescdGnrlListCnt=0"											'(*)�Ϲݱ��������Ʈ�Ǽ� | ���λ����� ���� �ؽ�Ʈ������̸�, ������ 0 Ȥ�� NULL�� ����
'		strRst = strRst & "&prdDescdHtmlItmListCnt="										'(*)�̹����׸���������Ʈ�Ǽ� | �����������ʵ� : 0 Ȥ�� NULL�� ����
		strRst = strRst & "&attrPrdListCnt="&CntOpt											'(*)�Ӽ�[�ɼ�]����Ʈ�Ǽ�
		strRst = strRst & "&prdSectListCnt="&CntCate										'������������Ʈ�Ǽ�
		strRst = strRst & "&prdGovPublsItmListCnt="&CntInfoCd								'(*)���ΰ���׸񸮽�Ʈ�Ǽ� | 1���̻��Է�
		strRst = strRst & "&prdDescdHtmlImgListCnt=0"										'(*)��ǰ�󼼱�����̹����Ǽ� | ��� �̹��������� ��ϵ� �� ����� �̹��� �Ǽ� ���°�� null �Ǵ� 0
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		strRst = strRst & "&brandCd="&brandcd												'(*)�귣���ڵ� | 7152�� ������ �ִ���?
		strRst = strRst & "&dlvPickMthodCd=3200"											'(*)��ۼ��Ź���ڵ� | 3200 : ����(�ù�)-��ü����
		strRst = strRst & "&dlvsCoCd="&DeliverCd											'(*)�ù���ڵ� | ����ù���ڵ�, �켱CJ�� ���
		strRst = strRst & "&saleStrDtm="&FormatDate(now(), "00000000000000")				'(*)�ǸŽ����Ͻ�
		strRst = strRst & "&saleEndDtm=29991231235959"										'(*)�Ǹ������Ͻ� | ��ǰ�� �ߴ�(�Ǹ�����)�Ϸ��� �ߴܽ����� �Ǹ������Ͻø� �Է��մϴ�.
		strRst = strRst & "&cardUseLimitYn=N"												'ī�������ѿ���
		strRst = strRst & "&baseAccmLimitYn=Y"												'(*)�⺻���������ѿ��� | �⺻�� : Y
		strRst = strRst & "&selAccmApplyYn=Y"												'(*)�������������뿩�� | �⺻�� : Y
		strRst = strRst & "&selAccRt="														'(*)���������� | �⺻�� : NULL
		strRst = strRst & "&immAccmDcLimitYn=Y"												'(*)����������������ѿ��� | �⺻�� : Y
		strRst = strRst & "&immAccmDcRt="													'(*)��������� | �⺻�� : NULL
		strRst = strRst & "&mnfcCoNm="&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)	'(*)�������
'		strRst = strRst & "&operMdId=80055"													'(*)�mdid
		strRst = strRst & "&operMdId="& getMdIdMapping(FDivcode)							'(*)�mdid
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)��ǰ�з��ڵ�
		strRst = strRst & "&orgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)�������� | ��ǰ�� ���������� �Է��մϴ�. ��)�̱�,�ѱ�,�߱� ��
		strRst = strRst & "&prdNm="&DDotFormat(getItemNameFormat, 10)						'(*)��ǰ��(����) | ����忡 �ԷµǴ� ��ǰ���Դϴ�.
		strRst = strRst & "&regChanlGrpCd=GE"												'(*)���ä�α׷��ڵ� | �Ǹ��� ��ǰ�� ä�α׷��ڵ��Դϴ�. GE : ���ͳݻ�ǰ
		strRst = strRst & "&ordPrdTypCd=02"													'(*)�ֹ���ǰ�����ڵ� | �Ӽ��� �ֹ����ɼ���(���)�� �����ϴ� �����ڵ��Դϴ�.02 : ��ǰ�Ӽ����ֹ��������� 01 : ��ǰ�� �ֹ���������
		'strRst = strRst & "&chrDlvYn="&CHKIIF(FSellcash>=30000, "N", "Y")					'(*)�����ۿ���
		strRst = strRst & "&taxTypCd="&CHKIIF(FVatInclude="N","01","02")					'(*)���������ڵ� | ��ǰ�� ���������� �Է��մϴ�. 01 : �鼼, 02 : ����, 03 : ����
		strRst = strRst & "&dlvDtGuideCd=N"													'(*)������ھȳ��ڵ� | �⺻�� : N
		strRst = strRst & "&prdTypCd="&CHKIIF(COptyn = "Y","S","P")							'(*)��ǰ�����ڵ� | ��ǰ�� �Ӽ�(�ɼ�)�� ������ �Է��մϴ�. P : �Ϲ� (�Ӽ������� ���� ���) S : �Ӽ� (����/������/����/����� �ִ� ���) | P�� ����� �Ŀ� S�κ����ϸ� �ɼ��߰�����//S->P�� �Ϲݻ�ǰ ��ȯ�� �� ��
		strRst = strRst & "&oboxCd="														'(*)�������ڵ� | �⺻�� : NULL
		strRst = strRst & "&chrDlvYn=Y"	'2016-06-21 19:28 ������ ����..3���� �̻��̸� N�� �������� 30000�� �̸� 2500�� �ڵ� : 7237257 �� �������� ������ Y�� ����..
		strRst = strRst & "&chrDlvcAmt=3000"												'�����ۺ�ݾ�
		strRst = strRst & "&shipLimitAmt=50000"												'�����ۺ�������رݾ�
		strRst = strRst & "&exchRtpChrYn=Y"													'(*)��ȯ��ǰ���Ῡ�� | ��ȯ,��ǰ�� ��ۺ� ������ ���θ� �Է��մϴ�.
		strRst = strRst & "&rtpAmt=6000"													'��ǰ�� | ��ǰ�� ����� �ݾ��� �Է� (��ȯ��ǰ���Ῡ�θ� Y�� �����ؾ� �ݿ���)
		strRst = strRst & "&exchAmt=6000"													'��ȯ�� | ��ȯ�� ����� �ݾ��� �Է� (��ȯ��ǰ���Ῡ�θ� Y�� �����ؾ� �ݿ���)
		strRst = strRst & "&chrDlvAddYn=N"													'(*)�������߰�����
		strRst = strRst & "&ilndDlvPsblYn=Y"												'���������۰��ɿ���
		strRst = strRst & "&jejuDlvPsblYn=Y"												'���ֵ���۰��ɿ���
		strRst = strRst & "&dd3InDlvNoadmtRegonYn=N"										'3�ϳ���ۺҰ���������
		strRst = strRst & "&ilndChrDlvYn=Y"													'�������������ۿ��� | ����-�ù��ϰ�츸 �߰�������
		strRst = strRst & "&ilndChrDlvcAmt=3000"											'�������������ۺ�	�������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndExchRtpChrYn=Y"												'�������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndRtpAmt=6000"												'���������ǰ�� | �������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndExchAmt=6000"												'�������汳ȯ�� | �������� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuChrDlvYn=Y"													'���ֵ������ۿ��� | ����-�ù��ϰ�츸 �߰������� ����
		strRst = strRst & "&jejuChrDlvcAmt=3000"											'���ֵ������ۺ� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuExchRtpChrYn=Y"												'���ֵ���ȯ��ǰ���Ῡ��	���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuRtpAmt=6000"												'���ֵ���ǰ�� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuExchAmt=6000"												'���ֵ���ȯ�� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&prdGbnCd=00"													'(*)��ǰ�����ڵ� | �Ϲݻ�ǰ,����ǰ,��ǰ�� �����ϴ� ���Դϴ�.00 : �Ϲݻ�ǰ, 02 : ����ǰ-��ü����
		strRst = strRst & "&bundlDlvCd=A01"													'(*)��������ڵ� | ������� ����/�Ұ����� �����ϴ� ���Դϴ�. A01 : ����, A02 : �Ұ���
		strRst = strRst & "&modelNo="														''''�𵨹�ȣ
		strRst = strRst & "&cpnApplyTypCd=09"												'(*)�������������ڵ� | �������� ���� �Ǵ� �����ϴ� ���Դϴ�. 00 : �������, 03 : ��ǰ������ ����, 09 : ��������
		If Fitemdiv="06" OR Fitemdiv="16" OR FtenCateLarge="040" Then
			strRst = strRst & "&openAftRtpNoadmtYn=Y"										'(*)�����Ĺ�ǰ�Ұ����� | �⺻�� : Y,N	(�ֹ������� Y // �ƴѰ� N)
		Else
			strRst = strRst & "&openAftRtpNoadmtYn=N"										'(*)�����Ĺ�ǰ�Ұ����� | �⺻�� : Y,N	(�ֹ������� Y // �ƴѰ� N)
		End If
		strRst = strRst & "&istTypCd="														'(*)�԰������ڵ� | �⺻�� : NULL
'		strRst = strRst & "&chrDlvcCd=7237257"												'(*)�����ۺ��ڵ�
		strRst = strRst & "&prdRelspAddrCd="&DeliverAddrCd									'(*)��ǰ������ּ��ڵ�
		strRst = strRst & "&prdRetpAddrCd="&DeliverAddrCd									'(*)��ǰ�ݼ����ּ��ڵ�
		strRst = strRst & "&separOrdNoadmtYn=N"												'(*)�ܵ��ֹ��Ұ����� | �⺻�� : N
		strRst = strRst & "&gftTypCd=00"													'(*)����ǰ�����ڵ� | 00 : �ǸŻ�ǰ, 02 : ����ǰ-��ü����
		strRst = strRst & "&prchTypCd=03"													'(*)���������ڵ� | 03 : ���������
		strRst = strRst & "&zrwonSaleYn=N"													'(*)0���Ǹſ���
		strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)�������»��ڵ� | ���»��ڵ�� �����ϰ� �Է�
		strRst = strRst & getzCostomMadeInd													'(*)�ֹ����ۿ��� �� �׸� �Լ�ȣ��
		strRst = strRst & "&attrTypExposCd=L"												'(*)�Ӽ����������ڵ� | L : ����Ʈ
		strRst = strRst & "&adultCertYn="&Chkiif(IsAdultItem() = "Y", "Y", "N")&""			'(*)������������	(�켱��N����)
		strRst = strRst & "&barcdNo="														'���ڵ��ȣ
		strRst = strRst & "&apntDlvDlvsCoCd="												'(*)��������ù���ڵ� | �⺻�� : NULL
		strRst = strRst & "&apntPickDlvsCoCd="												'(*)���������ù���ڵ� | �⺻�� : NULL
		strRst = strRst & "&gnuinYn=N"														'(*)��ǰ���� | �⺻�� : N
		strRst = strRst & "&frmlesPrdTypCd=N"												'(*)������ǰ�����ڵ� | �⺻�� : N
		strRst = strRst & "&rsrvSalePrdYn=N"												'�����Ǹſ���
		'�����̻��� �ɼ��̶�� �ɼ�Ÿ�Ը��� �������� ������Ű�� CJMall�� ���� 2~3�� �ɼ��� ������ �ʰ� �ϳ��� �ɼǿ� �� �ְ�..
		strRst = strRst & "&attrTypNm1="&CHKIIF(COptyn = "Y","����","")						'�Ӽ�������1 | �Ӽ������� �Ӽ��� Ÿ��Ʋ�� �����ϰ��� �Ҷ� ���̴� �÷�. ������ �������, ���� ���� ǥ�õȴ�.
		strRst = strRst & "&attrTypNm2="													'�Ӽ�������2 | �Ӽ������� Ÿ��Ʋ�� �����ϰ��� �Ҷ� ���̴� �÷� ������ �������, ������ �� ǥ�õȴ�.
		strRst = strRst & "&attrTypNm3="													'�Ӽ�������3 | �Ӽ������� Ÿ��Ʋ�� �����ϰ��� �Ҷ� ���̴� �÷� ������ �������, ��Ÿ�� ���� ǥ�õȴ�.
		strRst = strRst & "&attrTypNm4="													'�Ӽ�������4 | �Ӽ������� Ÿ��Ʋ�� �����ϰ��� �Ҷ� ���̴� �÷� ������ �������, ����ǰ ���� ǥ�õȴ�.
		strRst = strRst & "&attrSaleEndStModYn="											'�Ӽ��Ǹ�������¼������� | �Ӽ�����(S) ��ǰ�ǸŻ��¸� ������ �� ����ϴ� �׸�����, ��ǰ������ ���� �� ���� �� �Ӽ���ǰ�� ���µ� �Բ� ���� �� �����Ϸ��� Y, ��ǰ�����Ϳ� �Ӽ� ������ ���º��� ���� �ÿ� N
		'��ǰȮ��(prdAddInfo)
		strRst = strRst & "&prdBaseCmposCntnt="&Trim(chrbyte(getItemNameFormat,56,"Y"))		'(*)��ǰ�⺻�������� | ��ǰ��� �����ϰ� �Է�
		strRst = strRst & "&orgprdPkgCnt=1"													'(*)��ǰ���尹��
		strRst = strRst & "&prdAddCmposCntnt="												'��ǰ�߰���������
		strRst = strRst & "&addCmposPkgCnt="												'�߰��������尳��
		strRst = strRst & "&addCmposOrgpNm="												'�߰�������������
		strRst = strRst & "&addCmposMnfcCoNm="												'�߰������������
		strRst = strRst & "&prdGftCmposCntnt="												'��ǰ����ǰ��������
		strRst = strRst & "&gftPkgCnt="														'����ǰ���尳��
		strRst = strRst & "&gftCmposOrgpNm="												'����ǰ������������
		strRst = strRst & "&gftCmposMnfcCoNm="												'����ǰ�����������
		strRst = strRst & "&prdUnitValCd40=A01"												'(*)��ǰ�������� | A01 : 2.5kg�̸�, A02 : 2.5kg�̻� ~ 5kg�̸�, A03 : 5kg�̻� ~ 20kg�̸�, A04 : 30kg�̻�, A05 : 20kg�̻� ~ 30kg�̸�
		strRst = strRst & "&prdUnitValCd20=B01"												'(*)��ǰ�������� | B01 : 80cm�̸�, B02 : 80cm�̻� ~ 120cm�̸�, B03 : 120cm�̻� ~ 160cm�̸�, B04 : 160cm�̻�
		'��ǰ��������(prdSchdInfo)
'		strRst = strRst & "&prdSchdInfoRsrvOrdStrDt="										'�����ֹ����ɽ����Ͻ� | ��ǰ�⺻�� �����Ǹſ��ΰ� 'Y'�� ��츸 �ʼ��Է��׸��Դϴ�.
'		strRst = strRst & "&prdSchdInfoRsrvOrdEndDt="										'�����ֹ����������Ͻ� | ��ǰ�⺻�� �����Ǹſ��ΰ� 'Y'�� ��츸 �ʼ��Է��׸��Դϴ�.
'		strRst = strRst & "&prdSchdInfoRsrvRelsStrDt="										'�����������Ͻ� | ��ǰ�⺻�� �����Ǹſ��ΰ� 'Y'�� ��츸 �ʼ��Է��׸��Դϴ�.
'		strRst = strRst & "&prdSchdInfoRsrvRelsEndDt="										'������������Ͻ� | ��ǰ�⺻�� �����Ǹſ��ΰ� 'Y'�� ��츸 �ʼ��Է��׸��Դϴ�.
		'��ǰ����(prdPrc)
		strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdPrcSalePrc="&MustPrice										'(*)�ǸŰ���
'		strRst = strRst & "&prdPrcPrchPrc="													'(SYS)���԰��� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
		strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)���»�������/���ڵ� | 01 : ��
		strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice()						'(*)���»�������/�� | �⺻�� : �ǸŰ�*(1-0.12)
		'�����ǰ��(prdNmChg)
		strRst = strRst & "&prdNmChgValidStrDtm="&FormatDate(now(), "00000000000000")		'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdNmChgValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdNmChgExposPrdNm=" & Trim(chrbyte(getItemNameFormat,56,"Y"))					'(*)�����ǰ�� | GSShop�����ǰ��
		'��ǰ�̹���(prdCntntList)
		strRst = strRst & NmImage

		If v = 1 Then
			'��ǰ�󼼱����(prdDescdHtml)
			strRst = strRst & getGSShopItemContParam()
		Else
			strRst = strRst & "&prdDescdHtmlDescdExplnCntnt=" & Server.URLEncode("<div align=""center""><p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		End If

		'��ǰ�⺻-�Ӽ�
		strRst = strRst & NmOpt
		'��ǰ���ø���(prdSectList)
		strRst = strRst & NmCate															'(*)�����������̵�
		'��������(prdSafeCertInfo)
		strRst = strRst & getGSShopItemSafeInfoParam()
		'���ΰ���׸�(prdGovPublsItmList)
		strRst = strRst & NmInfoCd
		'rw strRst
		'response.end
		getGSShopItemNewRegParameter = strRst
	End Function

	'��ǰǰ������
	public function getGSShopItemInfoCdParam()
		Dim strSql, bufcnt, buf, certNum
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		' strSql = ""
		' strSql = strSql & " SELECT TOP 1 certNum "
		' strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg "
		' strSql = strSql & " WHERE itemid='"&FItemID&"' "
		' rsget.CursorLocation = adUseClient
		' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		' If Not(rsget.EOF or rsget.BOF) then
		' 	certNum = rsget("certNum")
		' End If
		' rsget.Close

		' buf = ""
		' strSql = ""
		' strSql = strSql & " SELECT TOP 100 M.* , " & vbcrlf
		' strSql = strSql & "		CASE " & vbcrlf
        ' strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') and isnull(IC.safetyNum, '') = '' THEN '"&certNum&"' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') and isnull(IC.safetyNum, '') <> '' THEN IC.safetyNum " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00001') THEN '��ǰ��������' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00002') THEN '������������' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00003') THEN '�ֿ��������' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00004') THEN '�ش����' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00005') THEN '������ǰ' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00006') THEN '�ǰ���ɽ�ǰ' " & vbcrlf
		' strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		' strSql = strSql & "			WHEN c.infotype='P' AND c.infoCd <> '22009' THEN '�ٹ����� ���ູ���� 1644-6035' " & vbcrlf
		' strSql = strSql & "			WHEN LEN(F.infocontent) <= 1 THEN F.infocontent + ' ����' " & vbcrlf
		' strSql = strSql & "		ELSE convert(varchar(500),F.infocontent) " & vbcrlf
		' strSql = strSql & " END AS infocontent " & vbcrlf
		' strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		' strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		' strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		' strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		' strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' and IC.itemid='"&FItemID&"' "
		' rsget.CursorLocation = adUseClient
		' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		' bufcnt = rsget.RecordCount
		' If Not(rsget.EOF or rsget.BOF) then
		' 	Do until rsget.EOF
		' 	    mallinfoCd  = rsget("mallinfoCd")
		' 	    infoContent = rsget("infoContent")
		' 		infocd		= rsget("infocd")
		' 		mallinfodiv = rsget("mallinfodiv")
		' 		If isnull(infoContent) Then
		' 			infoContent = ""
		' 		End If

		' 		infoContent = replace(infoContent, "&", "��")
		' 		infoContent = replace(infoContent, "?", "��")
		' 		infoContent = replace(infoContent, "%", "��")

		' 		buf = buf & "&govPublsItmCd="&mallinfoCd						'(*)���ΰ���׸�
		' 		buf = buf & "&govPublsItmCntnt="&infoContent					'(*)���ΰ���׸񳻿�
		' 		rsget.MoveNext
		' 	Loop
		' End If
		' rsget.Close

		buf = ""
		strSql = ""
		strSql = strSql & " EXEC db_item.dbo.usp_API_GSShop_InfoCodeMap_Get " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
				If isnull(infoContent) Then
					infoContent = ""
				End If

				infoContent = replace(infoContent, "&", "��")
				infoContent = replace(infoContent, "?", "��")
				infoContent = replace(infoContent, "%", "��")

				buf = buf & "&govPublsItmCd="&mallinfoCd						'(*)���ΰ���׸�
				buf = buf & "&govPublsItmCntnt="&infoContent					'(*)���ΰ���׸񳻿�
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSShopItemInfoCdParam = bufcnt&"|_|"&buf
	End Function

	'//��ǰ���� �Ķ���� ����
	Public Function getGSShopItemContParam()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		strRst = strRst & Server.URLEncode("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & Server.URLEncode("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & Server.URLEncode("<tr>")
		strRst = strRst & Server.URLEncode("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""��ǰ��"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & Server.URLEncode("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '����', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & Server.URLEncode("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat
		strRst = strRst & Server.URLEncode("</p>")
		strRst = strRst & Server.URLEncode("</td>")
		strRst = strRst & Server.URLEncode("</tr>")
		strRst = strRst & Server.URLEncode("</table>")
		strRst = strRst & Server.URLEncode("</div>")

		If ForderComment <> "" Then
			strRst = strRst & Server.URLEncode("- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>")
		End If

		Fitemcontent = replace(Fitemcontent,"&nbsp;"," ")
		Fitemcontent = replace(Fitemcontent,"&nbsp"," ")
		Fitemcontent = replace(Fitemcontent,"&"," ")
		Fitemcontent = replace(Fitemcontent,chr(13)," ")
		Fitemcontent = replace(Fitemcontent,chr(10)," ")
		Fitemcontent = replace(Fitemcontent,chr(9)," ")

		Select Case FUsingHTML
			Case "Y"
				'strRst = strRst & Server.URLEncode(Fitemcontent & "<br>")
				strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case "H"
				'strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case Else
				'strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
				strRst = strRst & nl2br(ReplaceBracket(Fitemcontent)) & "<br>"
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
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_gsshop.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		'(*)����������� | GSShop�� ����Ǵ� HTML�����		prdDescdHtmlDescdExplnCntnt
		getGSShopItemContParam = "&prdDescdHtmlDescdExplnCntnt=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = Server.URLEncode("<div align=""center"">"& strtextVal & "</div>")
			getGSShopItemContParam = "&prdDescdHtmlDescdExplnCntnt=" & strRst
		End If
		rsget.Close
	End Function

	'//��ǰ���� �Ķ���� ����
	Public Function getGSShopItemContParamEucKR()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		strRst = strRst & Server.URLEncode("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & Server.URLEncode("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & Server.URLEncode("<tr>")
		strRst = strRst & Server.URLEncode("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""��ǰ��"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & Server.URLEncode("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '����', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & Server.URLEncode("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & Server.URLEncode(getItemNameFormat)
		strRst = strRst & Server.URLEncode("</p>")
		strRst = strRst & Server.URLEncode("</td>")
		strRst = strRst & Server.URLEncode("</tr>")
		strRst = strRst & Server.URLEncode("</table>")
		strRst = strRst & Server.URLEncode("</div>")

		If ForderComment <> "" Then
			strRst = strRst & Server.URLEncode("- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>")
		End If

		Fitemcontent = replace(Fitemcontent,"&nbsp;"," ")
		Fitemcontent = replace(Fitemcontent,"&nbsp"," ")
		Fitemcontent = replace(Fitemcontent,"&"," ")
		Fitemcontent = replace(Fitemcontent,chr(13)," ")
		Fitemcontent = replace(Fitemcontent,chr(10)," ")
		Fitemcontent = replace(Fitemcontent,chr(9)," ")

		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				'strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case "H"
				strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				'strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case Else
				strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
				'strRst = strRst & nl2br(ReplaceBracket(Fitemcontent)) & "<br>"
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
					strRst = strRst & Server.URLEncode("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#��� ���ǻ���
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_gsshop.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		'(*)����������� | GSShop�� ����Ǵ� HTML�����		prdDescdHtmlDescdExplnCntnt
		getGSShopItemContParamEucKR = "&prdDescdHtmlDescdExplnCntnt=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = Server.URLEncode("<div align=""center"">"& strtextVal & "</div>")
			getGSShopItemContParamEucKR = "&prdDescdHtmlDescdExplnCntnt=" & strRst
		End If
		rsget.Close
If (session("ssBctID") = "kjy8517") Then
	' rw FUsingHTML & "!!!!!!!!!!!!!!!!!!!!!!"
	' rw getGSShopItemContParamEucKR
end if
	End Function

	'��ǰ �̹���
	Public Function getGSShopAddImageParam()
		Dim strRstCnt, strRst, strSQL, i
		'���� ������� �̹���
		'(*)�̹���url | ���� ū �̹����� URL �Է��ϸ� �ڵ�������¡ ó���� (GSShop �ִ��̹��� : 550x550)
		strRst = "&prdCntntListCntntUrlNm="&Server.URLEncode(FbasicImage)
		strRstCnt = 1
		'�̴ϻ�����  �̹���
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "&prdCntntListCntntUrlNm=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
					strRstCnt = strRstCnt + 1
				End If
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getGSShopAddImageParam = strRstCnt&"|_|"&strRst
	End Function

	'�ɼ� �Ķ���� ����
	Public Function getGSShopOptionParam()
		Dim strSql, strRst, itemSu, validSellno, optionname, fixday, optaddprice
		Dim ret, bufcnt, optyn, i
		ret = ""
		strSql = ""
		strSql = strSql & " SELECT T.* "
		strSql = strSql & " INTO #T1 "
		strSql = strSql & "	FROM ( "
		strSql = strSql & " 	SELECT i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(96),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " 	, o.optlimitno, o.optlimitsold " & VBCRLF
		strSql = strSql & " 	,Case When (o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold > 5) Then 'Y' " & VBCRLF
		strSql = strSql & " 		  When (o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold <= 5) Then 'N' " & VBCRLF
		strSql = strSql & " 		  When (isnull(o.itemid, '') = '') Then 'Y' " & VBCRLF		'-- �ɼ� ���� ��ǰ�̳� Y�� ó��
		strSql = strSql & " 	else o.optsellyn end as optsellyn " & VBCRLF
		strSql = strSql & " 	, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " 	,DATALENGTH(o.optionname) as optnmLen, isnull(r.outmallOptCode,'') as outmallOptCode" & VBCRLF
		strSql = strSql & " 	FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' and o.optaddprice=0 " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.[dbo].tbl_outmall_regedoption as r on i.itemid = r.itemid and o.itemoption = r.itemoption and r.mallid = '"&CMALLNAME&"' " & VBCRLF
		strSql = strSql & " 	WHERE i.itemid = "&Fitemid
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " WHERE T.optsellyn = 'Y' "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " SELECT t.itemoption "
		strSql = strSql & " INTO #T2 "
		strSql = strSql & " FROM #T1 as t "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_outmall_regedoption as r on  r.outmallOptName = t.optionname  "
		strSql = strSql & " WHERE r.mallid = '"&CMALLNAME&"' "
		strSql = strSql & " and r.itemid = " & Fitemid
		strSql = strSql & " and t.outmallOptCode = '' "
		strSql = strSql & " GROUP BY t.itemoption "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " SELECT * FROM #T1 "
		strSql = strSql & " WHERE itemoption not in ( "
		strSql = strSql & " 	SELECT itemoption FROM #T2 "
		strSql = strSql & " ) "
		strSql = strSql & " ORDER BY optaddprice, itemoption "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''���ϻ�ǰ
				    optionname="����"
				    FItemoption = "0000"
					itemSu = GetGSLmtQty()
					optyn	= "N"
					bufcnt	= 1
				Else
					FItemoption		= rsget("itemoption")
					optionname 		= db2Html(rsget("optionname"))
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					itemSu = getOptionLimitNo()
					optyn	= "Y"
					If rsget("optnmLen") > 80 Then
					    optionname=DdotFormat(optionname,40)
					End If
				End If

				'2016-01-22 14:38 ������ URL ���ڵ��� �ɼǸ�����
				optionname = replace(optionname,"&","��")
				optionname = replace(optionname,"%","����")
				optionname = replace(optionname,"+","%2B")
				optionname = replace(optionname,","," ")

				ret = ret & "&attrPrdListSupAttrPrdCd="&FItemoption							'Null�̶���� Null�� �����ϸ� �� ��'(SYS)���»�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
'				ret = ret & "&attrPrdListAttrPrdCd="&Chkiif(rsget("outmallOptCode") <> "", rsget("outmallOptCode"), "")	'(*)(SYS)GS�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
				ret = ret & "&attrPrdListAttrValCd1=00000"									'(*)�Ӽ����ڵ�1 | �⺻�� : 00000
				ret = ret & "&attrPrdListAttrValCd2=00000"									'(*)�Ӽ����ڵ�2 | �⺻�� : 00000
				ret = ret & "&attrPrdListAttrValCd3=00000"									'(*)�Ӽ����ڵ�3 | �⺻�� : 00000
				ret = ret & "&attrPrdListAttrValCd4=00000"									'(*)�Ӽ����ڵ�4 | �⺻�� : 00000
				ret = ret & "&attrPrdListSaleStrDtm="&FormatDate(now(), "00000000000000")	'(*)�ǸŽ����Ͻ�
				ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)�Ǹ������Ͻ�
				ret = ret & "&attrPrdListModelNo="											'�𵨹�ȣ
				ret = ret & "&attrPrdListAttrVal1="&optionname								'(*)�Ӽ���1 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ���� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
				ret = ret & "&attrPrdListAttrVal2="&ChkIIF(optyn="Y","None","����")			'(*)�Ӽ���2 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ����� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
				ret = ret & "&attrPrdListAttrVal3="&ChkIIF(optyn="Y","None","����")			'(*)�Ӽ���3 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ��Ÿ�ϰ� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
				ret = ret & "&attrPrdListAttrVal4="&ChkIIF(optyn="Y","None","����")			'(*)�Ӽ���4 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ����ǰ�� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��, (��ǰ�� �������ؼ� �ִ� ����ǰ)
'				ret = ret & "&attrPrdListArsAttrVal1="										'(*)�ڵ��ֹ��Ӽ���1 | �⺻�� : NULL
'				ret = ret & "&attrPrdListArsAttrVal2="										'(*)�ڵ��ֹ��Ӽ���2 | �⺻�� : NULL
'				ret = ret & "&attrPrdListArsAttrVal3="										'(*)�ڵ��ֹ��Ӽ���3 | �⺻�� : NULL
'				ret = ret & "&attrPrdListArsAttrVal4="										'(*)�ڵ��ֹ��Ӽ���4 | �⺻�� : NULL
'				ret = ret & "&attrPrdListAttrPkgCnt="										'(*)�Ӽ����尳�� | �⺻�� : NULL
				ret = ret & "&attrPrdListAttrCmposCntnt="									'(*)�Ӽ��������� | �⺻�� : NULL
				ret = ret & "&attrPrdListOrgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)��������
				ret = ret & "&attrPrdListMnfcCoNm="&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)	'(*)�������
				ret = ret & "&attrPrdListSafeStockQty=5"									'(*)���������� | ����������Ϸ� ������ �������� ���MD���� �˸��� ��
				ret = ret & "&attrPrdListTempoutYn=N"										'(*)�Ͻ�ǰ������ | �⺻�� : N
'				ret = ret & "&attrPrdListTempoutDtm="										'�Ͻ�ǰ���Ͻ�
				ret = ret & "&attrPrdListChanlGrpCd=AZ"										'(*)ä�α׷��ڵ� | AZ : DM��(DM�� ������ ������ ä��)
				ret = ret & "&attrPrdListOrdPsblQty="&itemSu								'(*)�ֹ����ɼ���
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getGSShopOptionParam = optyn&"|_|"&bufcnt&"|_|"&ret
	End Function

	Public Function getGSShopImageEditParameter()
		Dim strRst
		'################################ �̹��� ����Ʈ ���� ȣ�� #################################
		Dim CallImage, CntImage, NmImage
		CallImage = getGSShopAddImageParam()
		CntImage = Split(CallImage, "|_|")(0)
		NmImage = Split(CallImage, "|_|")(1)
		'##########################################################################################
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=I"														'(*)�������� I : �̹��� ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&prdCntntListCnt="&CntImage										'(*)�̹�������Ʈ�Ǽ� | ��ǰ�̹�������Ʈ (prdCntntList) �ݺ�Ƚ���� �����մϴ�.
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��ǰ�̹���(prdCntntList)
		strRst = strRst & NmImage
		getGSShopImageEditParameter = strRst
	End Function

	Public Function getGSShopSafeCertEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=C"														'(*)�������� C : ������������
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��������(prdSafeCertInfo)
		strRst = strRst & getGSShopItemSafeInfoParam()
		getGSShopSafeCertEditParameter = strRst
	End Function

	Public Function getGSShopItemEditParameter()
		Dim strRst
		Dim DeliverCd, DeliverAddrCd
		DeliverCd = "HJ"																	'�����ù�
		DeliverAddrCd = "0001"														'0001�� ��� ���� �Ϸ�(������ ����)

		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=A"														'(*)�������� A: ��ǰ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&regSubjCd=SUP"													'(*)�����ü�ڵ� | ���� ������ ��� : MD, ���»簡 ������ ��� : SUP
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		strRst = strRst & "&dlvsCoCd="&DeliverCd											'(*)�ù���ڵ� | ����ù���ڵ�, �켱CJ�� ���
		strRst = strRst & "&orgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)�������� | ��ǰ�� ���������� �Է��մϴ�. ��)�̱�,�ѱ�,�߱� ��
		strRst = strRst & "&chrDlvYn=Y"														'(*)�����ۿ��� | ������ ��� ����
		strRst = strRst & "&chrDlvcAmt=3000"												'�����ۺ�ݾ�
		strRst = strRst & "&shipLimitAmt=50000"												'�����ۺ�������رݾ�
		strRst = strRst & "&exchRtpChrYn=Y"													'(*)��ȯ��ǰ���Ῡ�� | ������ ��� ����
		strRst = strRst & "&rtpAmt=6000"													'��ǰ�� | ��ǰ�� ����� �ݾ��� �Է� (��ȯ��ǰ���Ῡ�θ� Y�� �����ؾ� �ݿ���)
		strRst = strRst & "&exchAmt=6000"													'��ȯ�� | ��ȯ�� ����� �ݾ��� �Է� (��ȯ��ǰ���Ῡ�θ� Y�� �����ؾ� �ݿ���)
		strRst = strRst & "&chrDlvAddYn=N"													'(*)�������߰����� | ������ ��� ����
		strRst = strRst & "&ilndDlvPsblYn=Y"												'���������۰��ɿ���
		strRst = strRst & "&jejuDlvPsblYn=Y"												'���ֵ���۰��ɿ���
		strRst = strRst & "&dd3InDlvNoadmtRegonYn=N"										'3�ϳ���ۺҰ���������
		strRst = strRst & "&ilndChrDlvYn=Y"													'�������������ۿ��� | ����-�ù��ϰ�츸 �߰�������
		strRst = strRst & "&ilndChrDlvcAmt=3000"											'�������������ۺ�	�������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndExchRtpChrYn=Y"												'�������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndRtpAmt=6000"												'���������ǰ�� | �������� �߰���ۺ� ������ ���
		strRst = strRst & "&ilndExchAmt=6000"												'�������汳ȯ�� | �������� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuChrDlvYn=Y"													'���ֵ������ۿ��� | ����-�ù��ϰ�츸 �߰������� ����
		strRst = strRst & "&jejuChrDlvcAmt=3000"											'���ֵ������ۺ� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuExchRtpChrYn=Y"												'���ֵ���ȯ��ǰ���Ῡ��	���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuRtpAmt=6000"												'���ֵ���ǰ�� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&jejuExchAmt=6000"												'���ֵ���ȯ�� | ���ֵ� �߰���ۺ� ������ ���
		strRst = strRst & "&prdRelspAddrCd="&DeliverAddrCd									'(*)��ǰ������ּ��ڵ�
		strRst = strRst & "&prdRetpAddrCd="&DeliverAddrCd									'(*)��ǰ�ݼ����ּ��ڵ�
		getGSShopItemEditParameter = strRst
	End Function

	'// ��ǰ �����(��ǰ����) ���� �Ķ���� ����
	Public Function getGSShopContentsEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=D"														'(*)�������� D : ����� ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��ǰ�̹��������(prdDescdHtml)
		strRst = strRst & getGSShopItemContParamEucKR()
		getGSShopContentsEditParameter = strRst
	End Function


	'// ��ǰ ���� ��� �׸� ���� �Ķ���� ����
	Public Function getGSShopInfodivEditParameter()
		'################################ ���� ��� �׸� ���� ȣ�� ################################
		Dim CallInfoCd, CntInfoCd, NmInfoCd
		CallInfoCd = getGSShopItemInfoCdParam()
		CntInfoCd = Split(CallInfoCd, "|_|")(0)
		NmInfoCd = Split(CallInfoCd, "|_|")(1)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=G"														'(*)�������� G : ���� ��� �׸� ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&prdGovPublsItmListCnt="&CntInfoCd								'(*)���ΰ���׸񸮽�Ʈ�Ǽ� | 1���̻��Է�
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)��ǰ�з��ڵ�
		'���ΰ���׸�(prdGovPublsItmList)
		strRst = strRst & NmInfoCd
		'rw strRst
		'response.end
		getGSShopInfodivEditParameter = strRst
	End Function

	'// ���ø��� ���� �Ķ���� ����
	Public Function getGSShopCategoryEditParameter()
		'################################ �������� �׸� ���� ȣ�� #################################
		Dim CallCate, CntCate, NmCate
		CallCate = getGSCateParam()
		CntCate = Split(CallCate, "|_|")(0)
		NmCate = Split(CallCate, "|_|")(1)
		'##########################################################################################

		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=M"														'(*)�������� M : ��������
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&prdSectListCnt="&CntCate										'(*)������������Ʈ�Ǽ� | 1���̻��Է�
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��ǰ���ø���(prdSectList)
		strRst = strRst & NmCate
		'rw strRst
		'response.end
		getGSShopCategoryEditParameter = strRst
	End Function

	'�ɼ� �ǸŻ��� ����
	Public Function getGSShopOptionEditParam()
		Dim strSql, arrRows, isOptionExists, tmpCnt
		Dim ret, bufcnt, i, itemoption, optLimit, optlimityn, isUsing, optsellyn, optNameDiff, forceExpired, ooptCd
		ret = ""
		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_gsshop '"&CMALLNAME&"'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic

		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		tmpCnt = 0
		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				itemoption			= ArrRows(1,i)
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				optNameDiff			= (ArrRows(12,i)=1)
				forceExpired		= (ArrRows(13,i)=1)
				ooptCd				= ArrRows(15,i)

				If LEN(ooptCd) > 2 Then
					ret = ret & "&attrPrdListSupAttrPrdCd="&itemoption							'Null�̶���� Null�� �����ϸ� �� ��'(SYS)���»�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
					ret = ret & "&attrPrdListAttrPrdCd="&ooptCd									'(*)(SYS)GS�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
					If ((forceExpired) or (optNameDiff) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5)) Then
						ret = ret & "&attrPrdListSaleEndDtm="&FormatDate(now(), "00000000000000")	'(*)�Ǹ������Ͻ�
						tmpCnt = tmpCnt + 1
					Else
						ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)�Ǹ������Ͻ�
					End If
				End If
			Next

			If FOptionCnt = 1 AND tmpCnt = 1 AND i = 1 Then
				FOptNotMatch = "Y"
			ElseIf (FOptionCnt > 1) AND (i = tmpCnt) Then
				FOptNotMatch = "Y"
			End If
		End If
		getGSShopOptionEditParam = bufcnt&"|_|"&ret
	End Function

	'// ��ǰ �ɼ� �߰� �� ���� ���� �Ķ���� ����
	Public Function getGSShopOptParameter()
		'################################ �Ӽ�(�ɼ�) �׸� ���� ȣ�� ###############################
		Dim CallOpt, COptyn, CntOpt, NmOpt
		CallOpt = getGSShopOptionParam()
		COptyn = Split(CallOpt, "|_|")(0)
		CntOpt = Split(CallOpt, "|_|")(1)
		NmOpt = Split(CallOpt, "|_|")(2)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=SA"														'(*)�������� SA : �Ӽ��߰� �� �ֹ����ɼ�������
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&attrPrdListCnt="&CntOpt											'(*)�Ӽ�[�ɼ�]����Ʈ�Ǽ�
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		strRst = strRst & "&prdTypCd="&CHKIIF(COptyn = "Y","S","P")							'(*)��ǰ�����ڵ� | ��ǰ�� �Ӽ�(�ɼ�)�� ������ �Է��մϴ�. P : �Ϲ� (�Ӽ������� ���� ���) S : �Ӽ� (����/������/����/����� �ִ� ���) | P�� ����� �Ŀ� S�κ����ϸ� �ɼ��߰�����//S->P�� �Ϲݻ�ǰ ��ȯ�� �� ��
		strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)�������»��ڵ� | �������»�� �������� �ʴ� ��� ���»��ڵ�� �����ϰ� �Է����ּž� �մϴ�.
		'��ǰ�⺻-�Ӽ�
		strRst = strRst & NmOpt
		getGSShopOptParameter = strRst
	End Function

	'// ��ǰ �ɼ� ���� ���� �Ķ���� ����
	Public Function getGSShopOptSellParameter()
		'################################ �Ӽ�(�ɼ�) �׸� ���� ȣ�� ###############################
		Dim CallOptSell, COptyn, CntOptSell, NmOptSell
		CallOptSell	= getGSShopOptionEditParam()
		CntOptSell	= Split(CallOptSell, "|_|")(0)
		NmOptSell	= Split(CallOptSell, "|_|")(1)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=SS"														'(*)�������� SS : �Ӽ��Ǹ�����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&attrPrdListCnt="&CntOptSell										'(*)�Ӽ�[�ɼ�]����Ʈ�Ǽ�
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��ǰ�⺻-�Ӽ�
		strRst = strRst & NmOptSell
		getGSShopOptSellParameter = strRst
	End Function

End Class

Class CGSShop
	Public FOneItem
	Public FItemList()

	Public FTotalCount
	Public FResultCount
	Public FCurrPage
	Public FTotalPage
	Public FPageSize
	Public FScrollCount
	Public FPageCount

	Public FRectMakerid
	Public FRectItemID
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectGSShopgoodno
	Public FRectMatchCate
	Public FRectPrdDivMatch
	Public FRectIsMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectGSShopYes10x10No
	Public FRectGSShopNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectFailCntExists
	Public FRectReqEdit

    ''���ļ���
    Public FRectOrdType

	'�귣�� ����
	Public FRectIsMaeip
	Public FRectIsDeliMapping
	Public FRectIsbrandcd
	Public FRectCatekey

	'��ǰ�з�
	Public FInfodiv
	Public FCateName
	Public FsearchName

	'ī�װ�
	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectdisptpcd
	Public FRectDspNo

	Private Sub Class_Initialize()
		Redim  FItemList(0)
		FCurrPage =1
		FPageSize = 30
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	End Function

	Public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	End Function

	Public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	End Function

	'// �̵�� ��ǰ ���(��Ͽ�)
	Public Sub getGSShopNotRegOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' �ɼ� �߰��ݾ� �ִ°�� ��� �Ұ�. //�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & " select itemid from ("
            addSql = addSql & "     select itemid"
            addSql = addSql & " 	,count(*) as optCNT"
            addSql = addSql & " 	,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
            addSql = addSql & " 	,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
            addSql = addSql & " 	from db_item.dbo.tbl_item_option"
            addSql = addSql & " 	where itemid in (" & FRectItemID & ")"
            addSql = addSql & " 	and isusing='Y'"
            addSql = addSql & " 	group by itemid"
            addSql = addSql & " ) T"
            addSql = addSql & " where (optCnt-optNotSellCnt<1)"
'            addSql = addSql & " or optAddCNT>0"
            addSql = addSql & " )"

            ''' 2013/05/29 Ư��ǰ�� ��� �Ұ� (ȭ��ǰ, ��ǰ��)
            'addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
			addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18')"	''2022-06-17 �����Կ�û..��ǰ �Ǹ�
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent , isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(R.gsshopStatCD,-9) as gsshopStatCD, IsNull(R.GSShopPrice, 0) as GSShopPrice "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.dtlCd, '') as divcode, isnull(pm.safecode, '') as safecode, uc.socname_kor "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_regItem R on i.itemid=R.itemid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "						'�ö��/ȭ�����/�ؿ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
'		strSql = strSql & " and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
'		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_gsshop_regItem where gsshopStatCD>3) "	''gsshopStatCD>=3 ��ϿϷ��̻��� ��Ͼȵ�.
		strSql = strSql & "	and IsNull(R.GSShopGoodNo, '') = '' "									'��ϻ�ǰ ����
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGSShopItem
				FOneItem.FItemid			= rsget("itemid")
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
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FGSShopStatCD		= rsget("gsshopStatCD")
				FOneItem.FGsshopprice		= rsget("gsshopprice")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")
                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
'                FOneItem.FBrandcd			= rsget("brandcd")
                FOneItem.FDeliveryType		= rsget("deliveryType")
'                FOneItem.FDeliveryCd		= rsget("deliveryCd")
'                FOneItem.FDeliveryAddrCd	= rsget("deliveryAddrCd")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FSocname_kor 		= rsget("socname_kor")
		End If
		rsget.Close
	End Sub

	'// GSShop ��ǰ ���(������)
	Public Sub getGSShopEditOneItem
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
		strSql = strSql & "	, m.gsshopGoodNo, m.gsshopprice, m.gsshopSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt, isNULL(convert(char(1), m.regedOptCnt), 'Y') as isNulltoTimeout  "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.dtlcd, '') as divcode, isnull(pm.safecode, '') as safecode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_GSShop_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and (m.GSShopStatCd = 3 OR m.GSShopStatCd = 7)  "
		strSql = strSql & addSql
		strSql = strSql & " and m.gsshopGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGSShopItem
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FGsshopGoodNo		= rsget("gsshopGoodNo")
				FOneItem.FGsshopprice		= rsget("gsshopprice")
				FOneItem.FGsshopSellYn		= rsget("gsshopSellYn")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName		= rsget("regImageName")
	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv       = rsget("infoDiv")
	            FOneItem.Fsafetyyn      = rsget("safetyyn")
	            FOneItem.FsafetyDiv     = rsget("safetyDiv")
	            FOneItem.FsafetyNum     = rsget("safetyNum")
	            FOneItem.FmaySoldOut    = rsget("maySoldOut")
	            FOneItem.FIsNulltoTimeout    = rsget("isNulltoTimeout")

                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
				FOneItem.FAdultType 		= rsget("adulttype")
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
