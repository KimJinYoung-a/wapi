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

	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut

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
	Public FNewitemname
	Public FItemnameChange


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
	Public FRegedOptionname
	Public FRegedItemname
	Public FItemoption
	Public FOptisusing
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public FOptaddprice
	Public FRealSellprice
	Public FNewItemid
	Public FOptionname

	Public Fvatinclude
	Public FGSShopStatCd
	Public FAdultType

	Public Function isDiffName
		isDiffName = false
		If (Fitemname <> FRegedItemname) OR (FOptionname <> FRegedOptionname) Then
			isDiffName = True
		End If
	End Function

	Public Function getRealItemname
		If FitemnameChange = "" Then
			getRealItemname = FNewitemname
		Else
			getRealItemname = FItemnameChange
		End If
	End Function

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
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
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
		buf = "[�ٹ�����]"&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"&","��")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"+","%2B")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	Function getItemOptNameFormat()
		Dim buf
		buf = "[�ٹ�����]"&replace(getRealItemname,"'","")		'���� ��ǰ�� �տ� [�ٹ�����] �̶�� ����
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"&","��")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"+","%2B")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemOptNameFormat = buf
	End Function

	'��ǰ�з��� ��������
	Public Function getGSShopItemSafeInfoParam()
		Dim buf, strSql
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
			buf = buf & "&safeCertFileNm="		'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
		Else							'SafeCode�� 1(�ʼ�,����)�̶��..
			strSql = ""
			strSql = strSql & " SELECT TOP 1 safeCertGbnCd, safeCertOrgCd, safeCertModelNm, safeCertNo, safeCertDt " & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_gsshop_safeCode " & VBCRLF
			strSql = strSql & " WHERE itemid = "&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.EOF Then
				safeCertGbnCd	= rsget("safeCertGbnCd")
				safeCertOrgCd	= rsget("safeCertOrgCd")
				safeCertModelNm	= rsget("safeCertModelNm")
				safeCertNo		= rsget("safeCertNo")
				safeCertDt		= rsget("safeCertDt")
			End If
			rsget.Close

			If (Fsafetyyn) = "Y" AND (FSafecode = "1" OR FSafecode = "2") Then			'SafeCode�� 1(�ʼ�,����)�̰� �ٹ����ٿ� �����������ΰ� Y���
				buf = buf & "&safeCertGbnCd="&safeCertGbnCd								'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ, 4 : �����ǰ��������Ȯ��
				buf = buf & "&safeCertOrgCd="&safeCertOrgCd								'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertModelNm="&safeCertModelNm							'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertNo="&safeCertNo									'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertDt="&FormatDate(safeCertDt, "00000000000000")		'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertFileNm=Y"											'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			Else						'�� ���� ���� ���� �ش���� ó��
				buf = buf & "&safeCertGbnCd=0"		'(*)���������������� | 0 : �ش���׾���, 1 : �����������, 2 : ����ǰ��������, 3 : ����ǰ��������Ȯ�ι�ȣ
				buf = buf & "&safeCertOrgCd=0"		'(*)������� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertModelNm="		'�����𵨸� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertNo="			'������ȣ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertDt="			'������ | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
				buf = buf & "&safeCertFileNm="		'�����������ϸ� | �������� ���������ڵ尡 '0'�� ��� 0 �ƴҰ��� �Է�
			End If
		End If
		getGSShopItemSafeInfoParam = buf
	End Function

	Public Function getGSCateParam()
		Dim strSql, bufcnt, cateKey, BcateKey, buf
		buf = ""
		strSql = ""
		strSql = strSql & " SELECT top 100 c.CateKey, c.cateGbn "
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
				If rsget("cateGbn") = "B" Then
					BcateKey = rsget("CateKey")
				End If

			    cateKey  = rsget("CateKey")
				buf = buf & "&prdSectListSectid="&cateKey
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSCateParam = BcateKey&"|_|"&bufcnt&"|_|"&buf
	End Function

	'���»�������/�� | �⺻�� : �ǸŰ�*(1-0.13) // ����12��
    Function getGSShopSuplyPrice()
		getGSShopSuplyPrice = CLNG(FSellCash * (100-CGSSHOPMARGIN) / 100)
    End Function

	'���»�������/�� | �⺻�� : �ǸŰ�*(1-0.13) // ����12��
    Function getGSShopOptSuplyPrice()
		getGSShopOptSuplyPrice = CLNG(FRealSellprice * (100-CGSSHOPMARGIN) / 100)
    End Function

   ''�ֹ����� ����
    Public Function getzCostomMadeInd()
		Dim ordMnfcYn, ordMnfcTypCd, ordMnfcTermDdcnt, ordMnfcCntnt
		Dim buf
		If (Fitemdiv="06" or Fitemdiv="16") Then
			If Fitemdiv = "06" Then
				ordMnfcTypCd = "10"
				ordMnfcCntnt = "�ֹ����ۿ�û����"
			ElseIf Fitemdiv="16" Then
				ordMnfcTypCd = "20"
			End If

			If (FrequireMakeDay > 5) Then
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
		buf = buf & "&ordMnfcCntnt="&ordMnfcCntnt			'(*)�ֹ����۳��� | �ֹ����������� 10�� ���������� ��� �ʼ��Է��׸��Դϴ�.
		buf = buf & "&ordMnfcTermDdcnt="&ordMnfcTermDdcnt	'(*)�ֹ����۱Ⱓ�ϼ� | �ֹ����ۿ��ΰ� 'Y'�� ��� �ʼ��Է��׸��Դϴ�.('N'�� ���� NULL)
		buf = buf & "&ordMnfcTypCd="&ordMnfcTypCd			'(*)�ֹ����������ڵ� | �ֹ����ۿ��ΰ� 'Y'�� ��� �ʼ��Է��׸��Դϴ�.('N'�� ���� NULL) NULL : �ش����, 10 : ��������, 20 : �ֹ�������, 30 : �ֹ��ļ���
		buf = buf & "&ordMnfcYn="&ordMnfcYn					'(*)�ֹ����ۿ���
		getzCostomMadeInd = buf
    End Function

	'// ��ǰ��� �Ķ���� ����
	Public Function getGSShopItemNewRegParameter()
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

		DeliverCd = "CJ"															'CJ�ù�
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
		Dim CallCate, CntCate, NmCate, ZunCateKey
		CallCate = getGSCateParam()
		ZunCateKey = Split(CallCate, "|_|")(0)
		CntCate = Split(CallCate, "|_|")(1)
		NmCate = Split(CallCate, "|_|")(2)
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
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
		'strRst = strRst & "&operMdId="&Mdid												'(*)�mdid
		strRst = strRst & "&operMdId=80055"													'(*)�mdid
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)��ǰ�з��ڵ�
		strRst = strRst & "&orgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)�������� | ��ǰ�� ���������� �Է��մϴ�. ��)�̱�,�ѱ�,�߱� ��
		strRst = strRst & "&prdNm="&DDotFormat(getItemOptNameFormat, 15)					'(*)��ǰ��(����) | ����忡 �ԷµǴ� ��ǰ���Դϴ�.
		strRst = strRst & "&regChanlGrpCd=GE"												'(*)���ä�α׷��ڵ� | �Ǹ��� ��ǰ�� ä�α׷��ڵ��Դϴ�. GE : ���ͳݻ�ǰ
		strRst = strRst & "&ordPrdTypCd=02"													'(*)�ֹ���ǰ�����ڵ� | �Ӽ��� �ֹ����ɼ���(���)�� �����ϴ� �����ڵ��Դϴ�.02 : ��ǰ�Ӽ����ֹ��������� 01 : ��ǰ�� �ֹ���������
		'strRst = strRst & "&chrDlvYn="&CHKIIF(FRealSellprice>=30000, "N", "Y")				'(*)�����ۿ���
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
		strRst = strRst & "&openAftRtpNoadmtYn="&CHKIIF(Fitemdiv="06" OR Fitemdiv="16" ,"Y","N")	'(*)�����Ĺ�ǰ�Ұ����� | �⺻�� : Y,N	(�ֹ������� Y // �ƴѰ� N)
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
		strRst = strRst & "&prdBaseCmposCntnt="&Trim(chrbyte(getItemOptNameFormat,56,"Y"))	'(*)��ǰ�⺻�������� | ��ǰ��� �����ϰ� �Է�
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
		strRst = strRst & "&prdPrcSalePrc="&Clng(GetRaiseValue(FRealSellprice/10)*10)		'(*)�ǸŰ���
'		strRst = strRst & "&prdPrcPrchPrc="													'(SYS)���԰��� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
		strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)���»�������/���ڵ� | 01 : ��
		strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopOptSuplyPrice()					'(*)���»�������/�� | �⺻�� : �ǸŰ�*(1-0.12)
		'�����ǰ��(prdNmChg)
		strRst = strRst & "&prdNmChgValidStrDtm="&FormatDate(now(), "00000000000000")		'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdNmChgValidEndDtm=29991231235959"								'(*)��ȿ�����Ͻ�
		strRst = strRst & "&prdNmChgExposPrdNm=" & Trim(chrbyte(getItemOptNameFormat,56,"Y"))	'(*)�����ǰ�� | GSShop�����ǰ��
		'��ǰ�̹���(prdCntntList)
		strRst = strRst & NmImage
		'��ǰ�̹��������(prdDescdHtml)
		strRst = strRst & getGSShopItemContParam()
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
		Dim strSql, bufcnt, buf
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		buf = ""
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , " & vbcrlf
		strSql = strSql & "		CASE " & vbcrlf
        strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN IC.safetyNum " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00001') THEN '��ǰ��������' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00002') THEN '������������' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00003') THEN '�ֿ��������' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00004') THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00005') THEN '������ǰ' " & vbcrlf
		strSql = strSql & "			WHEN (M.infoCd='00006') THEN '�ǰ���ɽ�ǰ' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '�ش����' " & vbcrlf
		strSql = strSql & "			WHEN c.infotype='P' AND c.infoCd <> '22009' THEN '�ٹ����� ���ູ���� 1644-6035' " & vbcrlf
		strSql = strSql & "			WHEN LEN(F.infocontent) <= 1 THEN F.infocontent + ' ����' " & vbcrlf
		strSql = strSql & "		ELSE convert(varchar(500),F.infocontent) " & vbcrlf
		strSql = strSql & " END AS infocontent " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' and IC.itemid='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
				infocd		= rsget("infocd")
				mallinfodiv = rsget("mallinfodiv")
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
		strRst = strRst & getItemOptNameFormat
		strRst = strRst & Server.URLEncode("</p>")
		strRst = strRst & Server.URLEncode("</td>")
		strRst = strRst & Server.URLEncode("</tr>")
		strRst = strRst & Server.URLEncode("</table>")
		strRst = strRst & Server.URLEncode("</div>")

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
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
		strRst = strRst & getItemOptNameFormat

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
		Dim strSql, itemSu, outmallOptCode
		Dim ret, bufcnt, optyn, i
		ret = ""
		strSql = ""
		strSql = strSql & " SELECT TOP 1 outmallOptCode FROM "
		strSql = strSql & " db_item.[dbo].tbl_outmall_regedoption "
		strSql = strSql & " WHERE mallid = '"&CMALLNAME&"' "
		strSql = strSql & " and itemid = "&Fitemid
		strSql = strSql & " and itemoption = '"&FItemoption&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			outmallOptCode = rsget("outmallOptCode")
		End If
		rsget.Close

		itemSu	= getOptionLimitNo()
		optyn	= "N"
		bufcnt	= 1

		ret = ret & "&attrPrdListSupAttrPrdCd="&FItemoption							'Null�̶���� Null�� �����ϸ� �� ��'(SYS)���»�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
'		ret = ret & "&attrPrdListAttrPrdCd="&Chkiif(outmallOptCode <> "", outmallOptCode, "")	'(*)(SYS)GS�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
		ret = ret & "&attrPrdListAttrValCd1=00000"									'(*)�Ӽ����ڵ�1 | �⺻�� : 00000
		ret = ret & "&attrPrdListAttrValCd2=00000"									'(*)�Ӽ����ڵ�2 | �⺻�� : 00000
		ret = ret & "&attrPrdListAttrValCd3=00000"									'(*)�Ӽ����ڵ�3 | �⺻�� : 00000
		ret = ret & "&attrPrdListAttrValCd4=00000"									'(*)�Ӽ����ڵ�4 | �⺻�� : 00000
		ret = ret & "&attrPrdListSaleStrDtm="&FormatDate(now(), "00000000000000")	'(*)�ǸŽ����Ͻ�
		ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)�Ǹ������Ͻ�
		ret = ret & "&attrPrdListModelNo="											'�𵨹�ȣ
		ret = ret & "&attrPrdListAttrVal1=����"										'(*)�Ӽ���1 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ���� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
		ret = ret & "&attrPrdListAttrVal2=����"										'(*)�Ӽ���2 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ����� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
		ret = ret & "&attrPrdListAttrVal3=����"										'(*)�Ӽ���3 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ��Ÿ�ϰ� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��
		ret = ret & "&attrPrdListAttrVal4=����"										'(*)�Ӽ���4 | ��ǰ�⺻�� ��ǰ�����ڵ尡 P�� ��� : '����' ���� ������ �Ӽ����� 1��, S�� ��� : ����ǰ�� ������ 'None', ������ ���Է��ϰ� �Ӽ������� n��, (��ǰ�� �������ؼ� �ִ� ����ǰ)
'		ret = ret & "&attrPrdListArsAttrVal1="										'(*)�ڵ��ֹ��Ӽ���1 | �⺻�� : NULL
'		ret = ret & "&attrPrdListArsAttrVal2="										'(*)�ڵ��ֹ��Ӽ���2 | �⺻�� : NULL
'		ret = ret & "&attrPrdListArsAttrVal3="										'(*)�ڵ��ֹ��Ӽ���3 | �⺻�� : NULL
'		ret = ret & "&attrPrdListArsAttrVal4="										'(*)�ڵ��ֹ��Ӽ���4 | �⺻�� : NULL
'		ret = ret & "&attrPrdListAttrPkgCnt="										'(*)�Ӽ����尳�� | �⺻�� : NULL
		ret = ret & "&attrPrdListAttrCmposCntnt="									'(*)�Ӽ��������� | �⺻�� : NULL
		ret = ret & "&attrPrdListOrgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"��ǰ���� ����",Fsourcearea)	'(*)��������
		ret = ret & "&attrPrdListMnfcCoNm="&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)	'(*)�������
		ret = ret & "&attrPrdListSafeStockQty=5"									'(*)���������� | ����������Ϸ� ������ �������� ���MD���� �˸��� ��
		ret = ret & "&attrPrdListTempoutYn=N"										'(*)�Ͻ�ǰ������ | �⺻�� : N
'		ret = ret & "&attrPrdListTempoutDtm="										'�Ͻ�ǰ���Ͻ�
		ret = ret & "&attrPrdListChanlGrpCd=AZ"										'(*)ä�α׷��ڵ� | AZ : DM��(DM�� ������ ������ ä��)
		ret = ret & "&attrPrdListOrdPsblQty="&itemSu								'(*)�ֹ����ɼ���
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		'��ǰ�̹���(prdCntntList)
		strRst = strRst & NmImage
		getGSShopImageEditParameter = strRst
	End Function

	Public Function getGSShopItemEditParameter()
		Dim strRst
		Dim DeliverCd, DeliverAddrCd
		DeliverCd = "CJ"															'CJ�ù�
		DeliverAddrCd = "0001"														'0001�� ��� ���� �Ϸ�(������ ����)

		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)��ϱ��� U : ����
		strRst = strRst & "&modGbn=A"														'(*)�������� A: ��ǰ����
		strRst = strRst & "&regId="&COurRedId												'(*)�����
		strRst = strRst & "&regSubjCd=SUP"													'(*)�����ü�ڵ� | ���� ������ ��� : MD, ���»簡 ������ ��� : SUP
		'��ǰ�⺻(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)���»��ڵ�
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)��ǰ�з��ڵ�
		'���ΰ���׸�(prdGovPublsItmList)
		strRst = strRst & NmInfoCd
		getGSShopInfodivEditParameter = strRst
	End Function

	'�ɼ� �ǸŻ��� ����
	Public Function getGSShopOptionEditParam()
		Dim strSql, arrRows, regedOptname, regedOptCode, oSellOK
		Dim ret, bufcnt
		strSql = ""
		strSql = strSql & " SELECT TOP 1 outmallOptName, outmallOptCode FROM db_item.dbo.tbl_outmall_regedoption WHERE itemid = '"&FItemid&"' and itemoption = '"&FItemoption&"' and mallid = 'gsshop' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			regedOptname	= rsget("outmallOptName")
			regedOptCode	= rsget("outmallOptCode")
		End If
		rsget.close

		If (FOptisusing <> "Y") or (FOptsellyn <> "Y") or (FLimityn = "Y" and FOptlimitno - FOptlimitsold < 5) or (regedOptname <> FOptionname) Then
			oSellOK = "N"
		Else
			oSellOK = "Y"
		End If

		ret = ""
		ret = ret & "&attrPrdListSupAttrPrdCd="&FItemoption							'Null�̶���� Null�� �����ϸ� �� ��'(SYS)���»�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
		ret = ret & "&attrPrdListAttrPrdCd="&regedOptCode							'(*)(SYS)GS�Ӽ���ǰ�ڵ� | (SYS�� �����ʿ��� �ڵ����� �������ִ� �ڵ� �� ���� ���մϴ�. Null�� �����ֽø� �˴ϴ�.)
		If oSellOK = "N" Then
			ret = ret & "&attrPrdListSaleEndDtm="&FormatDate(now(), "00000000000000")	'(*)�Ǹ������Ͻ�
		Else
			ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)�Ǹ������Ͻ�
		End If
		getGSShopOptionEditParam = "1|_|"&ret
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
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
		strRst = strRst & "&supPrdCd="&FNewItemid											'(*)���»��ǰ�ڵ�
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

	Public FRectIdx
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
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, o.optaddprice, o.optionname "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent , isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(R.gsshopStatCD,-9) as gsshopStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & "	, isnull(pm.divcode, '') as divcode, isnull(pm.safecode, '') as safecode "
		strSql = strSql & "	, isnull(dm.brandcd, '') as brandcd, isnull(dm.deliveryCd, '') as deliveryCd, isnull(dm.deliveryAddrCd, '') as deliveryAddrCd "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'gsshop' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as R on R.midx = M.idx "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_brandDelivery_mapping as dm on i.makerid = dm.makerid "
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7')"
		IF (CUPJODLVVALID) then
		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash + o.optaddprice >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'�ö��/ȭ�����/�ؿ����� ��ǰ ����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
'		strSql = strSql & " and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_gsshop_regItem where gsshopStatCD>3) "	''gsshopStatCD>=3 ��ϿϷ��̻��� ��Ͼȵ�.										'�Ե���ϻ�ǰ ����
		strSql = strSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
		strSql = strSql & " and o.optsellyn = 'Y' "
		strSql = strSql & " and (o.optlimityn = 'N' or ((o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold >="&CMAXLIMITSELL&"))) "
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FGSShopStatCD		= rsget("gsshopStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")
                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
                FOneItem.FBrandcd			= rsget("brandcd")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FDeliveryCd		= rsget("deliveryCd")
                FOneItem.FDeliveryAddrCd	= rsget("deliveryAddrCd")
                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")

                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
                FOneItem.FOptionname		= rsget("optionname")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	'// GSShop ��ǰ ���(������)
	Public Sub getGSShopEditOneItem
		Dim strSql, addSql, i
        ''//���� ���ܻ�ǰ
        addSql = addSql & " and i.itemid not in ("
        addSql = addSql & "     select itemid from db_item.dbo.tbl_OutMall_etcLink"
        addSql = addSql & "     where stDt < getdate()"
        addSql = addSql & "     and edDt > getdate()"
        addSql = addSql & "     and mallid='"&CMALLNAME&"'"
        addSql = addSql & "     and linkgbn='donotEdit'"
        addSql = addSql & " )"

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, o.itemoption, o.isusing as optisusing, o.optsellyn, o.optlimitno, o.optlimitsold, isnull(o.optaddprice, 0) as optaddprice, o.optionname "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, R.gsshopGoodNo, R.gsshopprice, R.gsshopSellYn "
		strSql = strSql & "	, R.accFailCNT, R.lastErrStr "
		strSql = strSql & "	, isnull(M.newitemname, '') as newitemname, isnull(M.itemnameChange, '') as itemnameChange "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, M.optionname as regedOptionname, M.itemname as regedItemname  "
		strSql = strSql & "	, isnull(pm.divcode, '') as divcode, isnull(pm.safecode, '') as safecode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash + o.optaddprice < 10000))"
		strSql = strSql & "		or (i.deliveryType = 7) "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "
		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_option as o on i.itemid = o.itemid "
		strSql = strSql & " JOIN db_etcmall.[dbo].[tbl_Outmall_option_Manager] as M on M.itemid = o.itemid and M.itemoption = o.itemoption and M.mallid = 'gsshop' and M.idx = '"&FRectIdx&"' "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_gsshopAddoption_regitem as R on R.midx = M.idx "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_prdDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small and c.infodiv = pm.infodiv "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and (R.GSShopStatCd = 3 OR R.GSShopStatCd = 7)  "
		strSql = strSql & addSql
		strSql = strSql & " and R.gsshopGoodNo is Not Null "									'#��� ��ǰ��
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

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv       = rsget("infoDiv")
	            FOneItem.Fsafetyyn      = rsget("safetyyn")
	            FOneItem.FsafetyDiv     = rsget("safetyDiv")
	            FOneItem.FsafetyNum     = rsget("safetyNum")
                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
	            FOneItem.FRegedOptionname	= rsget("regedOptionname")
	            FOneItem.FRegedItemname		= rsget("regedItemname")

	            FOneItem.FmaySoldOut    = rsget("maySoldOut")
                FOneItem.FNewitemname		= rsget("newitemname")
                FOneItem.FItemnameChange	= rsget("itemnameChange")
                FOneItem.FItemoption		= rsget("itemoption")
                FOneItem.FOptisusing		= rsget("optisusing")
                FOneItem.FOptsellyn			= rsget("optsellyn")
                FOneItem.FOptlimitno		= rsget("optlimitno")
                FOneItem.FOptlimitsold		= rsget("optlimitsold")
                FOneItem.FOptaddprice		= rsget("optaddprice")
                FOneItem.FRealSellprice		= rsget("sellcash") + rsget("optaddprice")
                FOneItem.FNewItemid			= CStr(rsget("itemid")) & CStr(rsget("itemoption"))
                FOneItem.FOptionname		= rsget("optionname")
				FOneItem.FAdultType 		= rsget("adulttype")
		End If
		rsget.Close
	End Sub

	'�귣�� ����
	Public Sub getTengsshopBrandDeliverList
		If FRectMakerid <> "" Then
			addSql = addSql & " and C.userid = '"&FRectMakerid&"' "
		End If

		If FRectIsDeliMapping = "Y" Then
			addSql = addSql & " and M.deliveryCd is Not null and M.deliveryAddrCd is NOT null "
		ElseIf FRectIsDeliMapping = "N" Then
			addSql = addSql & " and (M.deliveryCd is null OR M.deliveryAddrCd is null) "
		End if

		If FRectIsbrandcd = "Y" Then
			addSql = addSql & " and M.brandcd is Not null "
		ElseIf FRectIsbrandcd = "N" Then
			addSql = addSql & " and M.brandcd is null "
		End if

		If FRectIsMaeip = "Y" Then
			addSql = addSql & " and c.maeipdiv <> 'U' "
		ElseIf FRectIsMaeip = "N" Then
			addSql = addSql & " and c.maeipdiv = 'U' "
		End if

		Dim sqlStr, i, addsql
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE c.isExtUsing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and p.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and c.isusing = 'Y' " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.userid, c.socname, c.socname_kor, p.defaultsongjangdiv, p.deliver_name, p.return_zipcode, p.return_address, p.return_address2, c.maeipdiv, isnull(m.deliveryCd, '') as deliveryCd, isnull(m.deliveryAddrCd, '') as deliveryAddrCd, isnull(m.brandcd, '') as brandcd, s.divname " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN db_partner.dbo.tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_order.dbo.tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE c.isExtUsing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and p.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " and c.isusing = 'Y' " & addSql
		sqlStr = sqlStr & " ORDER BY c.userid ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FUserid			= rsget("userid")
					FItemList(i).FSocname			= rsget("socname")
					FItemList(i).FSocname_kor		= rsget("socname_kor")
					FItemList(i).FDeliver_name		= rsget("deliver_name")
					FItemList(i).FReturn_zipcode	= Trim(rsget("return_zipcode"))
					FItemList(i).FReturn_address	= Trim(rsget("return_address"))
					FItemList(i).FReturn_address2	= Trim(rsget("return_address2"))
					FItemList(i).FMaeipdiv			= rsget("maeipdiv")
					FItemList(i).FDeliveryCd		= rsget("deliveryCd")
					FItemList(i).FDeliveryAddrCd	= rsget("deliveryAddrCd")
					FItemList(i).FBrandcd			= rsget("brandcd")
					FItemList(i).FDivname			= rsget("divname")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	Public Function getTengsshopOneBrandDeliver
		Dim sqlStr, addSql, addsql2

		If FRectMakerid <> "" Then
			addSql = addSql & " and C.userid='" & FRectMakerid & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 c.userid, C.socname, C.socname_kor, p.deliver_name, p.return_zipcode, p.return_address, p.return_address2, c.maeipdiv, m.deliveryCd, m.deliveryAddrCd, m.brandcd, s.divname " & VBCRLF
		sqlStr = sqlStr & " FROM db_user.dbo.tbl_user_c as c " & VBCRLF
		sqlStr = sqlStr & " JOIN [db_partner].[dbo].tbl_partner as p on c.userid = p.id " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN [db_order].[dbo].tbl_songjang_div as s on p.defaultsongjangdiv = s.divcd and s.isusing = 'Y' " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_gsshop_brandDelivery_mapping as m on c.userid = m.makerid " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new CGSShopItem
				FItemList(0).FUserid			= rsget("userid")
				FItemList(0).FSocname			= rsget("socname")
				FItemList(0).FSocname_kor		= rsget("socname_kor")
				FItemList(0).FDeliver_name		= rsget("deliver_name")
				FItemList(0).FReturn_zipcode	= rsget("return_zipcode")
				FItemList(0).FReturn_address	= rsget("return_address")
				FItemList(0).FReturn_address2	= rsget("return_address2")
				FItemList(0).FMaeipdiv			= rsget("maeipdiv")
				FItemList(0).FDeliveryCd		= rsget("deliveryCd")
				FItemList(0).FDeliveryAddrCd	= rsget("deliveryAddrCd")
				FItemList(0).FBrandcd			= rsget("brandcd")
				FItemList(0).FDivname			= rsget("divname")
		End If
		rsget.Close
	End Function

	'��ǰ�з�
	Public Function getTengsshopOneprdDiv
		Dim sqlStr, addSql, addsql2

		If FRectCDL<>"" Then
			addSql = addSql & " and v.cdlarge='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and v.cdmid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and v.cdsmall='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql2 = addSql2 & " and p.infodiv='" & Finfodiv & "'"
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 p.divcode, p.infodiv, p.tenCateLarge, p.tenCateMid, p.tenCateSmall, v.nmlarge, v.nmmid, v.nmsmall, T.cdd_NAME " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.vw_category as v " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_gsshop_prdDiv_mapping p on p.tenCateLarge = v.cdlarge and p.tenCateMid = v.cdmid and p.tenCateSmall = v.cdsmall " & addsql2
		sqlStr = sqlStr & " LEFT JOIN db_etcmall.dbo.tbl_gsshop_prdDiv as T on p.divcode = T.divcode " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addsql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		If not rsget.EOF Then
			Set FItemList(0) = new CGSShopItem
				FItemList(0).Finfodiv		= rsget("infodiv")
				FItemList(0).FtenCateLarge	= rsget("tenCateLarge")
				FItemList(0).FtenCateMid	= rsget("tenCateMid")
				FItemList(0).FtenCateSmall	= rsget("tenCateSmall")
				FItemList(0).FtenCDLName	= rsget("nmlarge")
				FItemList(0).FtenCDMName	= rsget("nmmid")
				FItemList(0).FtenCDSName	= rsget("nmsmall")
				FItemList(0).FDivcode		= rsget("divcode")
				FItemList(0).Fcdd_Name		= rsget("cdd_NAME")
		End If
		rsget.Close
	End Function

	Public Sub getgsshopPrdDivList
		Dim sqlStr, addSql, i

		If FsearchName <> "" Then
			addSql = addSql & " and (cdl_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cdm_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cds_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " or cdd_NAME like '%" & FsearchName & "%'"
			addSql = addSql & " )"
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_gsshop_prdDiv " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " * " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_gsshop_prdDiv " & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & addSql
		sqlStr = sqlStr & " ORDER BY divcode ASC"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FDivcode			= rsget("divcode")
					FItemList(i).FCdl_Name			= db2html(rsget("cdl_Name"))
					FItemList(i).FCdm_Name			= db2html(rsget("cdm_Name"))
					FItemList(i).FCds_Name			= db2html(rsget("cds_Name"))
					FItemList(i).FCdd_Name			= db2html(rsget("cdd_Name"))
					FItemList(i).FSafecode			= rsget("safecode")
					FItemList(i).FSafecode_NAME		= rsget("safecode_NAME")
					FItemList(i).FIsvat				= rsget("isvat")
					FItemList(i).FIsvat_NAME		= rsget("isvat_NAME")
					FItemList(i).FInfodiv1			= rsget("infodiv1")
					FItemList(i).FInfodiv2			= rsget("infodiv2")
					FItemList(i).FInfodiv3			= rsget("infodiv3")
					FItemList(i).FInfodiv4			= rsget("infodiv4")
					FItemList(i).FInfodiv5			= rsget("infodiv5")
					FItemList(i).FInfodiv6			= rsget("infodiv6")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �ٹ�����-gsshop ��ǰ�з� ����Ʈ
	Public Sub getTenGsshopprdDivList
		Dim sqlStr, addSql, i
		If FRectCDL<>"" Then
			addSql = addSql & " and i.cate_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and i.cate_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and i.cate_small='" & FRectCDS & "'"
		End if

		If Finfodiv <> "" Then
			addSql = addSql & " and c.infodiv='" & Finfodiv & "'"
		End if

		If FRectIsMapping <> "" Then
			If FRectIsMapping = "Y" Then
				addSql = addSql & " and isnull(P.divcode, '') <> '' "
			ElseIf FRectIsMapping = "N" Then
				addSql = addSql & " and isnull(P.divcode, '') = '' "
			End If
		End if

		If FCateName <> "" AND FsearchName <> "" Then
			Select Case FCateName
				Case "cdlnm"
					addSql = addSql & " and v.nmlarge like '%" & FsearchName & "%'"
				Case "cdmnm"
					addSql = addSql & " and v.nmmid like '%" & FsearchName & "%'"
				Case "cdsnm"
					addSql = addSql & " and v.nmsmall like '%" & FsearchName & "%'"
			End Select
		End if
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM  ( " & VBCRLF
		sqlStr = sqlStr & " 	SELECT c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " 	, v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & "		,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " 	INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " 	LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "		LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 		SELECT dm.divcode, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 		FROM db_etcmall.dbo.tbl_gsshop_prdDiv_mapping as dm "  & VBCRLF
		sqlStr = sqlStr & " 		JOIN db_etcmall.dbo.tbl_gsshop_prdDiv as pv on dm.divcode = pv.divcode "  & VBCRLF
		sqlStr = sqlStr & " 	) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv   "  & VBCRLF
		sqlStr = sqlStr & " 	WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " 	GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ) as T " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " c.infodiv, i.cate_large, i.cate_mid, i.cate_small " & VBCRLF
		sqlStr = sqlStr & " , v.nmlarge, v.nmmid, v.nmsmall , count(*) as icnt " & VBCRLF
		sqlStr = sqlStr & " ,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv as Pinfodiv "  & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item i " & VBCRLF
		sqlStr = sqlStr & " INNER JOIN db_item.dbo.tbl_item_contents c on i.itemid = C.itemid " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN db_item.dbo.vw_category v	on i.cate_large = v.cdlarge and i.cate_mid = v.cdmid and i.cate_small = v.cdsmall " & VBCRLF
		sqlStr = sqlStr & "	LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT dm.divcode, dm.tenCateLarge,dm.tenCateMid, dm.tenCateSmall, pv.cdd_Name, pv.cdl_Name, pv.cdm_Name, pv.cds_Name, dm.infodiv "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_gsshop_prdDiv_mapping as dm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_gsshop_prdDiv as pv on dm.divcode = pv.divcode "  & VBCRLF
		sqlStr = sqlStr & " ) P on P.tenCateLarge=i.cate_large and P.tenCateMid=i.cate_mid and P.tenCateSmall=i.cate_small and P.infodiv = c.infodiv  "  & VBCRLF
		sqlStr = sqlStr & " WHERE i.sellyn = 'Y' and v.nmlarge is not null and isNULL(c.infodiv,'')<>'' "&addsql&" " & VBCRLF
		sqlStr = sqlStr & " GROUP BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small, v.nmlarge, v.nmmid, v.nmsmall,P.divcode ,P.cdd_Name, P.cdl_Name, P.cdm_Name, P.cds_Name, P.infodiv " & VBCRLF
		sqlStr = sqlStr & " ORDER BY c.infodiv, i.cate_large, i.cate_mid, i.cate_small "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).Finfodiv		= rsget("infodiv")
					FItemList(i).FtenCateLarge	= rsget("cate_large")
					FItemList(i).FtenCateMid	= rsget("cate_mid")
					FItemList(i).FtenCateSmall	= rsget("cate_small")
					FItemList(i).FtenCDLName	= rsget("nmlarge")
					FItemList(i).FtenCDMName	= rsget("nmmid")
					FItemList(i).FtenCDSName	= rsget("nmsmall")
					FItemList(i).FIcnt			= rsget("icnt")
					FItemList(i).FDivcode		= rsget("divcode")
					FItemList(i).Fcdd_Name		= rsget("cdd_Name")
					FItemList(i).Fcdl_Name		= rsget("cdl_Name")
					FItemList(i).Fcdm_Name		= rsget("cdm_Name")
					FItemList(i).Fcds_Name		= rsget("cds_Name")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// �ٹ�����-gsshop ī�װ� ����Ʈ
	Public Sub getTengsshopCateList
		Dim sqlStr, addSql, i

		If FRectCDL<>"" Then
			addSql = addSql & " and s.code_large='" & FRectCDL & "'"
		End if

		If FRectCDM<>"" Then
			addSql = addSql & " and s.code_mid='" & FRectCDM & "'"
		End if

		If FRectCDS<>"" Then
			addSql = addSql & " and s.code_small='" & FRectCDS & "'"
		End if

		If FRectIsMapping = "Y" Then
			addSql = addSql & " and T.CateKey is Not null "
		ElseIf FRectIsMapping = "N" Then
			addSql = addSql & " and T.CateKey is null "
		End if

		If FRectKeyword<>"" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop �����ڵ� �˻�
					addSql = addSql & " and T.CateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���(�ٹ����� �Һз���)
					addSql = addSql & " and s.code_nm like '%" & FRectKeyword & "%'"
			End Select
		End if

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt, CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s  "  & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_gsshop_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_gsshop_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP " & CStr(FPageSize*FCurrPage) & VBCRLF
		sqlStr = sqlStr & " 	s.code_large,s.code_mid,s.code_small " & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_large Where code_large=s.code_large) as large_nm  "  & VBCRLF
		sqlStr = sqlStr & " ,(Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) as mid_nm "  & VBCRLF
		sqlStr = sqlStr & " ,code_nm as small_nm "  & VBCRLF
		sqlStr = sqlStr & " ,T.CateKey as DispNo , T.L_Name as DispLrgNm, T.M_Name as DispMidNm, isnull(T.S_Name, '') as DispSmlNm, isnull(T.D_Name, '') as D_Name, T.IsUsing as CateIsUsing,T.cateGbn as disptpcd, "  & VBCRLF
		sqlStr = sqlStr & " Case When (isnull(T.S_Name, '') = '') AND (isnull(T.D_Name, '') = '') Then T.M_Name "
		sqlStr = sqlStr & " 	 When (isnull(T.S_Name, '') <> '') AND (isnull(T.D_Name, '') = '') Then T.S_Name "
		sqlStr = sqlStr & " Else T.D_Name END as DispNm "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_cate_small as s " & VBCRLF
		sqlStr = sqlStr & " LEFT JOIN (  "  & VBCRLF
		sqlStr = sqlStr & " 	SELECT cm.CateKey, cm.tenCateLarge,cm.tenCateMid, cm.tenCateSmall,cc.D_Name,cc.L_Name,cc.M_Name,cc.S_Name, cc.isusing, cc.CateGbn  "  & VBCRLF
		sqlStr = sqlStr & " 	FROM db_etcmall.dbo.tbl_gsshop_cate_mapping as cm  "  & VBCRLF
		sqlStr = sqlStr & " 	JOIN db_etcmall.dbo.tbl_gsshop_category as cc on cc.CateKey = cm.CateKey  "  & VBCRLF
		If FRectdisptpcd <> "" Then
            sqlStr = sqlStr & " and cc.CateGbn='"&FRectdisptpcd&"'"
        End If
		sqlStr = sqlStr & " ) T on T.tenCateLarge=s.code_large and T.tenCateMid=s.code_mid and T.tenCateSmall=s.code_small  "  & VBCRLF
		sqlStr = sqlStr & " WHERE 1 = 1 " & VBCRLF
		sqlStr = sqlStr & " and (Select code_nm from db_item.dbo.tbl_cate_mid Where code_large=s.code_large and code_mid=s.code_mid) is not null  " & addSql
		sqlStr = sqlStr & " ORDER BY s.code_large,s.code_mid,s.code_small, T.CateGbn  ASC "
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.EOF
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FtenCateLarge		= rsget("code_large")
					FItemList(i).FtenCateMid		= rsget("code_mid")
					FItemList(i).FtenCateSmall		= rsget("code_small")
					FItemList(i).FtenCDLName		= db2html(rsget("large_nm"))
					FItemList(i).FtenCDMName		= db2html(rsget("mid_nm"))
					FItemList(i).FtenCDSName		= db2html(rsget("small_nm"))
					FItemList(i).FDispNo			= rsget("DispNo")
					FItemList(i).FDispNm			= rsget("DispNm")
					FItemList(i).FDispLrgNm			= rsget("DispLrgNm")
					FItemList(i).FDispMidNm			= rsget("DispMidNm")
					FItemList(i).FDispSmlNm			= rsget("DispSmlNm")
					FItemList(i).Fdisptpcd			= rsget("disptpcd")
	                FItemList(i).FCateIsUsing		= rsget("CateIsUsing")
	                FItemList(i).FD_NAME			= rsget("D_NAME")
				i = i + 1
				rsget.moveNext
			Loop
		End If
		rsget.Close
	End Sub

	'// gsshop ī�װ�
	Public Sub getgsshopCategoryList
		Dim sqlStr, addSql, i

		If FRectDspNo <> "" Then
			addSql = addSql & " and c.cateKey = " & FRectDspNo
		End If

		If FRectKeyword <> "" Then
			Select Case FRectSDiv
				Case "CCD"	'gsshop �����ڵ� �˻�
					addSql = addSql & " and c.cateKey='" & FRectKeyword & "'"
				Case "CNM"	'ī�װ���
					addSql = addSql & " and (c.D_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.S_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.M_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " or c.L_Name like '%" & FRectKeyword & "%'"
					addSql = addSql & " )"
			End Select
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(c.cateKey) as cnt, CEILING(CAST(Count(c.cateKey) AS FLOAT)/" & FPageSize & ") as totPg " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_gsshop_category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 " & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'������������ ��ü ���������� Ŭ �� �Լ�����
		If Cint(FCurrPage) > Cint(FTotalPage) Then
			FResultCount = 0
			Exit Sub
		End If

		sqlStr = ""
		sqlStr = sqlStr & " SELECT DISTINCT TOP " & CStr(FPageSize*FCurrPage) & " c.* " & VBCRLF
		sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_gsshop_category as c " & VBCRLF
		sqlStr = sqlStr & " WHERE 1=1 " & addSql
		sqlStr = sqlStr & " ORDER BY c.L_CODE, c.M_CODE, c.S_CODE, c.D_CODE ASC"
		rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))
		Redim preserve FItemList(FResultCount)
		i = 0
		If not rsget.EOF Then
			rsget.absolutepage = FCurrPage
			Do until rsget.eof
				Set FItemList(i) = new CGSShopItem
					FItemList(i).FDispNo		= rsget("cateKey")
					FItemList(i).FDispNm		= db2html(rsget("D_Name"))
					FItemList(i).FDispLrgNm		= db2html(rsget("L_Name"))
					FItemList(i).FDispMidNm		= db2html(rsget("M_Name"))
					FItemList(i).FDispSmlNm		= db2html(rsget("S_Name"))
					FItemList(i).FDispThnNm		= db2html(rsget("D_Name"))
					FItemList(i).FisUsing		= rsget("isUsing")
				i = i + 1
				rsget.moveNext
			Loop
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