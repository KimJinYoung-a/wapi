<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "coupang"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CCoupangItem
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
	Public FCoupangRegdate
	Public FCoupangLastUpdate
	Public FCoupangGoodNo
	Public FCoupangPrice
	Public FCoupangSellYn
	Public FregUserid
	Public FCoupangStatCd
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
	Public ForderComment
	Public FbasicImage
	Public FbasicimageNm
	Public FmainImage
	Public FmainImage2
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
				If (FLimitYN <> "Y") Then optLimit = 9999

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

	Public Function IsAllOptionChange
		Dim sqlStr, tenOptCnt, regedCoupangOptCnt
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
		sqlStr = sqlStr & " db_etcmall.dbo.tbl_coupang_regedoption "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and outmallOptName <> '���ϻ�ǰ' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			regedCoupangOptCnt = rsget("cnt")
		rsget.Close

		If tenOptCnt > 0 AND regedCoupangOptCnt = 0 Then			'��ǰ -> �ɼ�
			IsAllOptionChange = "Y"
		ElseIf tenOptCnt = 0 AND regedCoupangOptCnt > 0 Then		'�ɼ� -> ��ǰ
			IsAllOptionChange = "Y"
		Else
			IsAllOptionChange = "N"
		End If
	End Function

	Public Function getCoupangInfoDiv(infoDivName)
		Select Case infoDivName
			Case "�Ƿ�"								getCoupangInfoDiv =  "01"
			Case "����/�Ź�"							getCoupangInfoDiv =  "02"
			Case "����"								getCoupangInfoDiv =  "03"
			Case "�м���ȭ(����/��Ʈ/�׼�����)"			getCoupangInfoDiv =  "04"
			Case "ħ����/Ŀư"						getCoupangInfoDiv =  "05"
			Case "����"								getCoupangInfoDiv =  "06"
			Case "������(TV��)"						getCoupangInfoDiv =  "07"
			Case "������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ����� ��)"		getCoupangInfoDiv =  "08"
			Case "��������(������/��ǳ�� ��)"			getCoupangInfoDiv =  "09"
			Case "�繫����(��ǻ��/��Ʈ��/������ ��)"	getCoupangInfoDiv =  "10"
			Case "���б��(������ī�޶�/ķ�ڴ� ��)"		getCoupangInfoDiv =  "11"
			Case "�޴���"							getCoupangInfoDiv =  "13"
			Case "������̼�"							getCoupangInfoDiv =  "14"
			Case "�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)"		getCoupangInfoDiv =  "15"
			Case "�Ƿ���"							getCoupangInfoDiv =  "16"
			Case "�ֹ��ǰ"							getCoupangInfoDiv =  "17"
			Case "ȭ��ǰ"							getCoupangInfoDiv =  "18"
			Case "�ͱݼ�/����/�ð��"					getCoupangInfoDiv =  "19"
			Case "��ǰ(������깰)"					getCoupangInfoDiv =  "20"
			Case "������ǰ"							getCoupangInfoDiv =  "21"
			Case "�ǰ���ɽ�ǰ"						getCoupangInfoDiv =  "22"
			Case "�����ƿ�ǰ"							getCoupangInfoDiv =  "23"
			Case "�Ǳ�"								getCoupangInfoDiv =  "24"
			Case "��������ǰ"							getCoupangInfoDiv =  "25"
			Case "����"								getCoupangInfoDiv =  "26"
			Case "��ǰ�뿩 ����(������, ��, ����û���� ��)"						getCoupangInfoDiv =  "31"
			Case "������ ������(����, ����, ���ͳݰ��� ��)"							getCoupangInfoDiv =  "33"
			Case "��Ÿ ��ȭ"							getCoupangInfoDiv =  "35"
		End Select

		If Instr(infoDivName, "��������(MP") > 0 Then
			getCoupangInfoDiv =  "12"
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
		dim ret : ret = (FLimitno-FLimitSold)
		if (ret<1) then ret=0
		getLimitEa = ret
	End Function

	Public Function getCoupangStatName
	    If IsNULL(FCoupangStatCd) then FCoupangStatCd=-1
		Select Case FCoupangStatCd
			CASE -9 : getCoupangStatName = "�̵��"
			CASE -1 : getCoupangStatName = "��Ͻ���"
			CASE 0 : getCoupangStatName = "<font color=blue>��Ͽ���</font>"
			CASE 1 : getCoupangStatName = "���۽õ�"
			CASE 2 : getCoupangStatName = "�ݷ�"
			CASE 3 : getCoupangStatName = "������"
			CASE 7 : getCoupangStatName = ""
			CASE ELSE : getCoupangStatName = FCoupangStatCd
		End Select
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice
		GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
		If (GetTenTenMargin < CMAXMARGIN) Then
			tmpPrice = Forgprice
		Else
			tmpPrice = FSellCash
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'// Coupang �Ǹſ��� ��ȯ
	Public Function getCoupangSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getCoupangSellYn = "Y"
			Else
				getCoupangSellYn = "N"
			End If
		Else
			getCoupangSellYn = "N"
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CCoupang
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
	Public FRectCoupangGoodNo
	Public FRectMatchCate
	Public FRectMatchShipping
	Public FRectGosiEqual
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectCoupangNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectCoupangYes10x10No
	Public FRectCoupangNo10x10Yes
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


	Public Sub getCoupangNotRegOneItem
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
            addSql = addSql & " WHERE optCnt - optNotSellCnt < 1 "
            addSql = addSql & " )"

			'2019-06-03 17:37 ������ �ּ�ó��..ǰ���� �� �¾Ƶ� ����� �� �ְ� sp ����
			' addSql = addSql & " and c.infodiv in ( "
			' addSql = addSql & " SELECT  "
			' addSql = addSql & " 	CASE WHEN noticeCategoryName = '�Ƿ�' THEN '01' "
			' addSql = addSql & "  	WHEN noticeCategoryName = '����/�Ź�' THEN '02' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '����' THEN '03' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�м���ȭ(����/��Ʈ/�׼�����)' THEN '04' "
			' addSql = addSql & " 	WHEN noticeCategoryName = 'ħ����/Ŀư' THEN '05' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '����' THEN '06' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '������(TV��)' THEN '07' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '������ ������ǰ(�����/��Ź��/�ı⼼ô��/���ڷ����� ��)' THEN '08' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '��������(������/��ǳ�� ��)' THEN '09' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�繫����(��ǻ��/��Ʈ��/������ ��)' THEN '10' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '���б��(������ī�޶�/ķ�ڴ� ��)' THEN '11' "
			' addSql = addSql & " 	WHEN left(noticeCategoryName, 4) = '��������' THEN '12' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�޴���' THEN '13' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '������̼�' THEN '14' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�ڵ�����ǰ(�ڵ�����ǰ/��Ÿ �ڵ�����ǰ)' THEN '15' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�Ƿ���' THEN '16' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�ֹ��ǰ' THEN '17' "
			' addSql = addSql & " 	WHEN noticeCategoryName = 'ȭ��ǰ' THEN '18' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�ͱݼ�/����/�ð��' THEN '19' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '������ǰ' THEN '20' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '��ǰ(������깰)' THEN '21' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�ǰ���ɽ�ǰ' THEN '22' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '�����ƿ�ǰ' THEN '23' "
			' addSql = addSql & "		WHEN noticeCategoryName = '�Ǳ�' THEN '24' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '��������ǰ' THEN '25' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '����' THEN '26' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '��ǰ�뿩 ����(������, ��, ����û���� ��)' THEN '31' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '������ ������(����, ����, ���ͳݰ��� ��)' THEN '33' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '��Ÿ ��ȭ' THEN 35 END "
			' addSql = addSql & " FROM db_etcmall.dbo.tbl_coupang_categorynoti as si "
			' addSql = addSql & " WHERE si.CateKey = am.CateKey "
			' addSql = addSql & " ) "
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.coupangStatCD,-9) as coupangStatCD "
		strSql = strSql & "	, UC.socname_kor, am.CateKey "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_coupang_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_category as tm on am.CateKey = tm.CateKey "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_coupang_branddelivery_mapping as bm on i.makerid = bm.makerid "
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
'		strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and 1 = ( "
		strSql = strSql & " 	CASE WHEN CHARINDEX('������', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('����', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('��ź����', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('����', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('�Ͻ�', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('�ó�', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('������', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('���ڴ��', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 	ELSE 1 END "
		strSql = strSql & " ) "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'�ö��/ȭ�����/�ؿ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		'strSql = strSql & "	and isnull(R.CoupangStatCD,0) < 3  "
		'strSql = strSql & " and R.itemid is NULL"  '' 2018/11/23
		strSql = strSql & " and isNull(R.CoupangGoodNo, '') = ''"  '' 2019/04/12 �� R.itemid is null �ּ�
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & " and i.itemdiv not in ('06') "	''�ֹ����۹��� ��ǰ ����
		strSql = strSql & " and bm.outboundShippingPlaceCode is Not Null "
		strSql = strSql & "		"	& addSql											'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCoupangItem
				FOneItem.Fitemname			= db2html(rsget("itemname"))
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FsellYn			= rsget("sellYn")
				FOneItem.FisUsing			= rsget("isusing")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FoneItem.FoptionCnt			= rsget("optioncnt")  '' 2018/11/23
				FoneItem.FITemID 			= rsget("itemid")  '' 2018/11/23
		End If
		rsget.Close
	End Sub

	Public Sub getCoupangEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,m.coupangGoodNo, m.coupangSellyn "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & " 	or i.itemdiv in ('06') "		''�ֹ����۹��� ��ǰ ǰ��ó��
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_coupang_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.coupangGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCoupangItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")

				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FcoupangGoodNo		= rsget("coupangGoodNo")
				FOneItem.FCoupangSellYn		= rsget("coupangSellYn")

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
%>