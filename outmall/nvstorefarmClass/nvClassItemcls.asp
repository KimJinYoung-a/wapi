<%
CONST CMAXMARGIN = 10
CONST CMALLNAME = "nvstorefarmclass"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 0									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CDEFALUT_STOCK = 9999

Class CNvClassItem
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
	Public Fvatinclude
	Public ForderComment
	Public FoptionCnt
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FNvClassGoodNo
	Public FNvClassprice
	Public FNvClassSellyn
	Public FregedOptCnt
	Public FAccFailCNT
	Public FMaySoldOut
	Public Fregitemname
	Public FLastErrStr
	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FNvClassStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FSocname_kor
	Public FAPIaddImg
	Public FbasicimageNm
	Public FRegImageName
	Public FCateKey
	Public FNeedCert

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function MustPrice()
		MustPrice = FSellCash
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	'// ������� �Ǹſ��� ��ȯ
	Public Function getNvClassSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FSellYn="Y" and FIsUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getNvClassSellYn = "Y"
			Else
				getNvClassSellYn = "N"
			End If
		Else
			getNvClassSellYn = "N"
		End If
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


	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function getLimitNvClassEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold
			If ret > 10000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimitNvClassEa = ret
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FRegImageName
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

				If i = 0 Then		'0�� �ɼ��� ��� 0�� ���ϸ� ǰ��
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

	Function getItemNameFormat()
		Dim buf
		'buf = "[�ٹ����� Ŭ����] "&replace(FItemName,"'","")		'���� ��ǰ�� �տ� [�ٹ����� Ŭ����] �̶�� ����
		buf = replace(FItemName,"'","")
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","����")
		buf = replace(buf,"[������]","")
		buf = replace(buf,"[���� ���]","")
		getItemNameFormat = buf
	End Function

	Public Function getModelName
		Dim strSql, modelName, isRegCert, safetyDiv, safetyId
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN '121' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN '72' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN '1042' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN '51' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN '1020' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN '58' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN '1040' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN '1041' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN '1042' end as safetyId "
		strSql = strSql & " ,f.modelName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			modelName		= rsget("modelName")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		If isRegCert = "Y" and safetyDiv = "70" OR safetyDiv = "80" OR safetyDiv = "90" Then
			getModelName = "					<shop:ModelName>"&modelName&"</shop:ModelName>"
		Else
			getModelName = ""
		End If
	End Function

	'�ֹ� ���� ����
    Public Function getzCostomMadeInd()
		Dim buf, CustomMade, EstimatedDeliveryTime
		buf = "				<shop:CustomMade>N</shop:CustomMade>"		'# �ֹ� ���� ��ǰ ���� Y or N | Y: EstimatedDeliveryTime�Է� �ʼ�, "N": EstimatedDeliveryTime �Է� �Ұ�
		getzCostomMadeInd = buf
    End Function

	'������ ����
	Public Function getOriginAreaType
		Dim buf
		buf = ""
		buf = buf & "				<shop:OriginArea>"													'#������ ����
		If Fsourcearea = "�ѱ�" OR Fsourcearea = "���ѹα�" OR Fsourcearea = "����" Then
			buf = buf & "					<shop:Code>00</shop:Code>"									'#������ �� ���� | 00 : ����, 01 : �����, 02 : ���Ի�, 03 : �󼼼��� ǥ��, 04 : �����Է�
'			buf = buf & "					<shop:Importer></shop:Importer>"							'���Ի�� | ���Ի��� ��� �ʼ�
			buf = buf & "					<shop:Plural>N</shop:Plural>"								'���� ������ | Y or N
'			buf = buf & "					<shop:Content></shop:Content>"								'������ ǥ�� ���� | Code�� "��Ÿ:���� �Է�"�� ��� �ʼ�
		Else
			buf = buf & "					<shop:Code>04</shop:Code>"									'#������ �� ���� | 00 : ����, 01 : �����, 02 : ���Ի�, 03 : �󼼼��� ǥ��, 04 : �����Է�
'			buf = buf & "					<shop:Importer></shop:Importer>"							'���Ի�� | ���Ի��� ��� �ʼ�
			buf = buf & "					<shop:Plural>N</shop:Plural>"								'���� ������ | Y or N
			buf = buf & "					<shop:Content><![CDATA["&Fsourcearea&"]]></shop:Content>"	'������ ǥ�� ���� | Code�� "��Ÿ:���� �Է�"�� ��� �ʼ�
		End If
		buf = buf & "				</shop:OriginArea>"
		getOriginAreaType = buf
	End Function

	'// ��ǰ���: ��ǰ�߰��̹��� �Ķ���� ����
	Public Function getImageType()
		Dim buf, strSql, arrRows, i, basicimgStr, addimgStr
		addimgStr	= ""
		basicimgStr	= ""
		strSql = ""
		strSql = strSql & " SELECT TOP 10 imgType, storefarmURL FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_Image] WHERE itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			arrRows = rsget.getRows()
		End If
		rsget.Close

		If isArray(arrRows) then
			For i = 0 To UBound(arrRows, 2)
				If arrRows(0, i) = "1" Then
					basicimgStr = arrRows(1,i)																		'��ǥ �̹���
				Else
					addimgStr = addimgStr & "						<shop:Optional>"								'�߰� �̹���
					addimgStr = addimgStr & "							<shop:URL>"&arrRows(1,i)&"</shop:URL>"
					addimgStr = addimgStr & "						</shop:Optional>"
				End If
			Next
		End If

		buf = ""
		buf = buf & "				<shop:Image>"
		buf = buf & "					<shop:Representative>"
		buf = buf & "						<shop:URL>"&basicimgStr&"</shop:URL>"
		buf = buf & "					</shop:Representative>"
		If addimgStr <> "" Then
		buf = buf & "					<shop:OptionalList>"
		buf = buf & addimgStr
		buf = buf & "					</shop:OptionalList>"
		End If
		buf = buf & "				</shop:Image>"
		getImageType = buf
	End Function

	'// ��������
	Public Function getECouponType()
		Dim buf, strSql, arrRows, i, UsePlaceContents, ContactInformationContents
		strSql = ""
		strSql = strSql & " SELECT TOP 1 p.tPAddress, p.tPTel "
		strSql = strSql & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSql = strSql & " WHERE k.itemid = '"& FItemID &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			UsePlaceContents = rsget("tPAddress")
			ContactInformationContents = rsget("tPTel")
		End If
		rsget.Close

		If UsePlaceContents = "" Then
			UsePlaceContents = "����� ���α� ���з�12�� 31 �������� 2��"
		End If

		If ContactInformationContents = "" Then
			ContactInformationContents = "1644-6030"
		End If

		buf = ""
		buf = buf & "				<shop:AfterServiceTelephoneNumber><![CDATA["&ContactInformationContents&"]]></shop:AfterServiceTelephoneNumber>"		'#A/S ��ȭ��ȣ
		buf = buf & "				<shop:AfterServiceGuideContent><![CDATA[A/S ������ "&FSocname_kor&" ������� ���� ������ �ֽñ� �ٶ��ϴ�.]]></shop:AfterServiceGuideContent>"	'#A/S �ȳ�
		buf = buf & "				<shop:ECoupon>"											'ECOUPON | ������ ī�װ� ��ǰ�� ��� �ʼ�
		buf = buf & "					<shop:PeriodType>FB</shop:PeriodType>"				'#e���� ��ȿ�Ⱓ ���� | FX : Ư���Ⱓ, FB : �ڵ��Ⱓ
'		buf = buf & "					<shop:ValidStartDate></shop:ValidStartDate>"		'e���� ��ȿ�Ⱓ ������..YYYY-MM-DD����, e���� ��ȿ�Ⱓ ���� Ÿ��(PeriodType)�� Ư���Ⱓ�� ��� �ʼ�
'		buf = buf & "					<shop:ValidEndDate></shop:ValidEndDate>"			'e���� ��ȿ�Ⱓ ������..YYYY-MM-DD����, e���� ��ȿ�Ⱓ ���� Ÿ��(PeriodType)�� Ư���Ⱓ�� ��� �ʼ�
		buf = buf & "					<shop:PeriodDays>30</shop:PeriodDays>"				'e���� ��ȿ�Ⱓ ����..e���� ��ȿ�Ⱓ ���� Ÿ��(PeriodType)�� �ڵ� �Ⱓ�� ��� �ʼ�
		buf = buf & "					<shop:PublicInformationContents><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></shop:PublicInformationContents>"		'e���� ����ó
		buf = buf & "					<shop:ContactInformationContents><![CDATA["&ContactInformationContents&"]]></shop:ContactInformationContents>"	'e���� ����ó
		buf = buf & "					<shop:UsePlaceType>PLACE</shop:UsePlaceType>"		'e���� ��� ��� Ÿ�� | PLACE : ���, URL : URL
		buf = buf & "					<shop:UsePlaceContents><![CDATA["& UsePlaceContents &"]]></shop:UsePlaceContents>"	'e���� ��� ���
		buf = buf & "					<shop:RestrictCart>Y</shop:RestrictCart>"			'e���� ��ٱ��� ���� | Y or N
		buf = buf & "				</shop:ECoupon>"
		getECouponType = buf
	End Function

	Public Function getNvClassItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
'		strRst = strRst & ("<p><center><a href=""http://storefarm.naver.com/tenbytenclass"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_nvClass.jpg""></a></center></p><br>")

'		If ForderComment <> "" Then
'			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
'		End If

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#��� ���ǻ���
'		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_nvClass.jpg"">")
		strRst = strRst & ("</div>")

		Dim ticketPlaceName, tPAddress, tPTel, parkingGuide
		Dim strticketPlaceName, strtPAddress, strtPTel, strparkingGuide
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 isNull(p.ticketPlaceName, '') as ticketPlaceName, isNull(p.tPAddress, '') as tPAddress, isNull(p.tPTel, '') as tPTel, isNull(p.parkingGuide, '') as parkingGuide "
		strSQL = strSQL & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSQL = strSQL & " WHERE k.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			ticketPlaceName	= rsget("ticketPlaceName")
			tPAddress		= rsget("tPAddress")
			tPTel			= rsget("tPTel")
			parkingGuide	= rsget("parkingGuide")
		End If
		rsget.Close

		If ticketPlaceName <> "" Then
			strticketPlaceName = "<strong>[��Ҹ�]</strong><br />" & ticketPlaceName & "<br />"
		End If

		If tPAddress <> "" Then
			strtPAddress = "<strong>[�ּ�]</strong><br />" & tPAddress & "<br />"
		End If

		If tPTel <> "" Then
			strtPTel = "<strong>[��ȭ]</strong><br />" & tPTel & "<br />"
		End If

		If parkingGuide <> "" Then
			strparkingGuide = "<strong>[���� ����]</strong><br />" & nl2br(parkingGuide)
		End If

		If (ticketPlaceName <> "") OR (tPAddress <> "") OR (tPTel <> "") OR (parkingGuide <> "") Then
			strRst = strRst & "<div class=""alignCt"" style=""background-color:#f8f8f8; margin-top:100px; padding:57px 0px; width:100%"">"
			strRst = strRst & "<p style=""margin-bottom:0px; margin-left:0px; margin-right:0px; margin-top:0px; padding:0px 8%; text-align:center""><span style=""font-family:malgun gothic,&quot;���� ���&quot;,sans-serif""><span style=""color:#000000""><span style=""font-size:22px; font-weight:600; line-height:1.2"">��ġ ����</span></span></span></p>"
			strRst = strRst & "	<p style=""margin-bottom:0px; margin-left:0px; margin-right:0px; margin-top:0px; padding:0px 8%; text-align:center"">&nbsp;</p>"
			strRst = strRst & "	<div style=""padding:11px 8% 0px; text-align:left"">"
			strRst = strRst & "		<span style=""font-family:malgun gothic,&quot;���� ���&quot;,sans-serif"">"
			strRst = strRst & "			<span style=""font-size:16px"">"
			strRst = strRst & "				<span style=""color:#000000"">"
			strRst = strRst & "					<span style=""line-height:1.6"">"
			strRst = strRst & "						"& strticketPlaceName &" "
			strRst = strRst & "						"& strtPAddress &" "
			strRst = strRst & "						"& strtPTel &" "
			strRst = strRst & "						"& strparkingGuide &" "
			strRst = strRst & "					</span>"
			strRst = strRst & "				</span>"
			strRst = strRst & "			</span>"
			strRst = strRst & "		</span>"
			strRst = strRst & "		</p>"
			strRst = strRst & "	</div>"
			strRst = strRst & "</div>"
		End If
		getNvClassItemContParamToReg = strRst
	End Function

	Public Function getSellerComment
		Dim buf, icomment
		icomment = Fordercomment
		icomment = replace(icomment,"\","")
		icomment = replace(icomment,"*","")
		icomment = replace(icomment,"?","")
		icomment = replace(icomment,"""","")
		icomment = replace(icomment,"<","")
		icomment = replace(icomment,">","")
		icomment = replace(icomment,"&#160;"," ")	'2018-12-27 �̻��� �ƽ�Ű�� ġȯ..maybw �������� �ɾ�µ�
		buf = ""

		If len(icomment) > 1300 Then
			icomment = DDotFormat(icomment,1290)
		End If

		If len(icomment) = 2 AND instr(icomment, chr(13)) Then
			icomment = ""
		End If

		If IsNULL(icomment) OR Trim(icomment) = "" Then
			buf = buf & "				<shop:SellerCommentUsable>N</shop:SellerCommentUsable>"			'�Ǹ��� Ư�̻��� ��� ���� | Y or N..Y�Է½� SellerCommentContent �ʼ�, N �Է½� Ư�� ���� �������� ����Ǹ� SellerCommentContent �ʵ� ����..��ǰ ������ SellerCommentUsable ��Ҹ� �����ϰ� �����ϸ� ������ ����� ���� ������� �ʴ´�.
'			buf = buf & "				<shop:SellerCommentContent></shop:SellerCommentContent>"		'�Ǹ��� Ư�̻��� ���� �Է� �� | SellerCommentUsable�� Y�� �� ����
		Else
			buf = buf & "				<shop:SellerCommentUsable>Y</shop:SellerCommentUsable>"			'�Ǹ��� Ư�̻��� ��� ���� | Y or N..Y�Է½� SellerCommentContent �ʼ�, N �Է½� Ư�� ���� �������� ����Ǹ� SellerCommentContent �ʵ� ����..��ǰ ������ SellerCommentUsable ��Ҹ� �����ϰ� �����ϸ� ������ ����� ���� ������� �ʴ´�.
			buf = buf & "				<shop:SellerCommentContent><![CDATA["&icomment&"]]></shop:SellerCommentContent>"		'�Ǹ��� Ư�̻��� ���� �Է� �� | SellerCommentUsable�� Y�� �� ����
		End If
'		buf = buf & "				<shop:SellerCustomCode1></shop:SellerCustomCode1>"				'�Ǹ��ڰ� ���ο��� ����ϴ� �ڵ�
'		buf = buf & "				<shop:SellerCustomCode2></shop:SellerCustomCode2>"				'�Ǹ��ڰ� ���ο��� ����ϴ� �ڵ�
		getSellerComment = buf
	End Function

	Public Function getNvClassItemInfoCdToReg
		Dim buf, strSQL, mallinfoCd, infoContent, mallinfodiv, mallinfoName
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN isNull(TR.certNum, IC.safetyNum) " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn<> 'Y') THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '�ش����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '���ù� �� �Һ��ں����ذ���ؿ� ����' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'N') THEN 'N' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'Y') THEN 'Y' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) < 2 THEN i.itemname "
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) >= 2 THEN F.infocontent "
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN K.tPTel " & vbcrlf
		strSQL = strSQL & " 	 WHEN LEN(isnull(F.infocontent, '')) < 2 THEN '�������� ����' " & vbcrlf
		strSQL = strSQL & " ELSE isnull(F.infocontent, '') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.[dbo].[tbl_safetycert_tenReg] as TR ON i.itemid = TR.itemid "
		strSQL = strSQL & " LEFT JOIN ( "
		strSQL = strSQL & " 	SELECT TOP 1 k.itemid, isNull(p.tPTel, '�ٹ����� 1644-6030') as tPTel "
		strSQL = strSQL & " 	FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSQL = strSQL & " 	LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSQL = strSQL & " 	WHERE k.itemid = '"& FItemID &"' "
		strSQL = strSQL & " ) as K on K.itemid = IC.itemid "
		strSQL = strSQL & " WHERE M.mallid = 'nvstorefarm' and IC.itemid='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY infocd ASC " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = rsget("mallinfodiv")
			Select Case mallinfodiv
				Case "01"		mallinfoName = "Wear"
				Case "02"		mallinfoName = "Shoes"
				Case "03"		mallinfoName = "Bag"
				Case "04"		mallinfoName = "FashionItems"
				Case "05"		mallinfoName = "SleepingGear"
				Case "06"		mallinfoName = "Furniture"
				Case "07"		mallinfoName = "ImageAppliances"
				Case "08"		mallinfoName = "HomeAppliances"
				Case "09"		mallinfoName = "SeasonAppliances"
				Case "10"		mallinfoName = "OfficeAppliances"
				Case "11"		mallinfoName = "OpticsAppliances"
				Case "12"		mallinfoName = "MicroElectronics"
				Case "13"		mallinfoName = "Cellphone"
				Case "14"		mallinfoName = "Navigation"
				Case "15"		mallinfoName = "CarArticles"
				Case "16"		mallinfoName = "MedicalAppliances"
				Case "17"		mallinfoName = "KitchenUtensils"
				Case "18"		mallinfoName = "Cosmetic"
				Case "19"		mallinfoName = "Jewellery"
				Case "20"		mallinfoName = "Food"
				Case "21"		mallinfoName = "GeneralFood"
				Case "22"		mallinfoName = "DietFood"
				Case "23"		mallinfoName = "Kids"
				Case "24"		mallinfoName = "MusicalInstrument"
				Case "25"		mallinfoName = "SportsEquipment"
				Case "26"		mallinfoName = "Books"
				Case "27"		mallinfoName = "LodgmentReservation"
				Case "28"		mallinfoName = "TravelPackage"
				Case "30"		mallinfoName = "RentCar"
				Case "31"		mallinfoName = "RentalHa"
				Case "32"		mallinfoName = "RentalEtc"
				Case "33"		mallinfoName = "DigitalContents"
				Case "35"		mallinfoName = "Etc"
				Case "47"		mallinfoName = "Biochemistry"
				Case "48"		mallinfoName = "Biocidal"
			End Select

			buf = ""
			buf = buf & "				<shop:"&mallinfoName&">"
			buf = buf & "					<shop:NoRefundReason><![CDATA[�������� ����]]></shop:NoRefundReason>"
			buf = buf & "					<shop:ReturnCostReason><![CDATA[�������� ����]]></shop:ReturnCostReason>"
			buf = buf & "					<shop:QualityAssuranceStandard><![CDATA[�������� ����]]></shop:QualityAssuranceStandard>"
			buf = buf & "					<shop:CompensationProcedure><![CDATA[�������� ����]]></shop:CompensationProcedure>"
			buf = buf & "					<shop:TroubleShootingContents><![CDATA[�������� ����]]></shop:TroubleShootingContents>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
'			    If mallinfoCd = "Size" Then
				If infoContent <> "" Then
			    	infoContent = replace(infoContent, "*", "x")
			    End If
'			    End If
				buf = buf & "					<shop:"&mallinfoCd&"><![CDATA["&infoContent&"]]></shop:"&mallinfoCd&">"
				rsget.MoveNext
			Loop
			buf = buf & "				</shop:"&mallinfoName&">"
		End If
		rsget.Close
		getNvClassItemInfoCdToReg = buf
	End Function

	Public Function getNvClassItemInfoCdToRegOnlyMobile
		Dim buf, strSql, arrRows, i, UsePlaceContents, ContactInformationContents
		strSql = ""
		strSql = strSql & " SELECT TOP 1 p.tPAddress, p.tPTel "
		strSql = strSql & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSql = strSql & " WHERE k.itemid = '"& FItemID &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			UsePlaceContents = rsget("tPAddress")
			ContactInformationContents = rsget("tPTel")
		End If
		rsget.Close

		If UsePlaceContents = "" Then
			UsePlaceContents = "����� ���α� ���з�12�� 31 �������� 2��"
		End If

		If ContactInformationContents = "" Then
			ContactInformationContents = "1644-6030"
		End If

		buf = ""
		buf = buf & "				<shop:MobileCoupon>"
		buf = buf & "					<NoRefundReason><![CDATA[�������� ����]]></NoRefundReason>"							'��ǰ���ڰ� �ƴ� �Һ����� �ܼ�����, �������� �� ���� û��öȸ ���� �Ұ����� ��� �� �� ü�� ������ �ٰ�
		buf = buf & "					<ReturnCostReason><![CDATA[�������� ����]]></ReturnCostReason>"						'��ǰ����?����� � ���� û��öȸ ���� �� �� û��öȸ ���� �� �� �ִ� �Ⱓ �� ����� �ž��ڰ� �δ��ϴ� ��ǰ��� � ���� ����
		buf = buf & "					<QualityAssuranceStandard><![CDATA[�������� ����]]></QualityAssuranceStandard>"		'��ȭ ���� ��ȯ����ǰ������ ���� �� ǰ�� ���� ����
		buf = buf & "					<CompensationProcedure><![CDATA[�������� ����]]></CompensationProcedure>"				'����� ȯ�ҹޱ� ���� ����� ȯ���� ������ ��� ������ ���� ������ ���޹��� �� �ִ� �� ��� �� ���� ������ ��ü�� ���� �� �� ��
		buf = buf & "					<TroubleShootingContents><![CDATA[�������� ����]]></TroubleShootingContents>"			'�Һ��� ���� ������ ó��, ��ȭ � ���� �� �� ó�� �� �Һ��ڿ� ����� ������ ����ó�� �� ���� ����
		buf = buf & "					<Issuer><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></Issuer>"		'������
		buf = buf & "					<UsableCondition><![CDATA[�����Ϸκ��� 30��]]></UsableCondition>"						'��ȿ�Ⱓ, �̿� ����
		buf = buf & "					<UsableStore><![CDATA["& UsePlaceContents &"]]></UsableStore>"							'�̿� ���� ����
		buf = buf & "					<CancelationPolicy><![CDATA[�������� ����]]></CancelationPolicy>"						'ȯ�� ���� �� ���
		buf = buf & "					<CustomerServicePhoneNumber><![CDATA["&ContactInformationContents&"]]></CustomerServicePhoneNumber>"	'�Һ��� ��� ���� ��ȭ��ȣ
		buf = buf & "				</shop:MobileCoupon>"
		getNvClassItemInfoCdToRegOnlyMobile = buf
	End Function

	'// ���ε� �̹��� XML ����
	Public Function getNvClassImageRegXML(oServ, oOper)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature, strSQL, i
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1np6kl_01"
		End If
		Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, oOper)

		strRst = ""
		strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
		strRst = strRst & "	<soapenv:Header/>"
		strRst = strRst & "	<soapenv:Body>"
		strRst = strRst & "		<shop:UploadImageRequest>"
		strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
		strRst = strRst & "			<shop:AccessCredentials>"
		strRst = strRst & "				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
		strRst = strRst & "				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
		strRst = strRst & "				<shop:Signature>"&osignature&"</shop:Signature>"
		strRst = strRst & "			</shop:AccessCredentials>"
		strRst = strRst & "			<shop:Version>2.0</shop:Version>"
		strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
		strRst = strRst & "			<ImageURLList>"
		If (application("Svr_Info") = "Dev") Then
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/basic/146/B001469141.jpg</shop:URL>"
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/add1/146/A001469141_01.jpg</shop:URL>"
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/add2/146/A001469141_02.jpg</shop:URL>"
		Else
			strRst = strRst & "				<shop:URL>"&FbasicImage&"</shop:URL>"
			strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType=adOpenStatic
			rsget.Locktype=adLockReadOnly
			rsget.Open strSQL, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				For i=1 to rsget.RecordCount
					If rsget("imgType") = "0" Then
						strRst = strRst & "				<shop:URL>"&"http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"</shop:URL>"
					End If
					rsget.MoveNext
					If i >= 4 Then Exit For
				Next
			End If
			rsget.Close
		End If
		strRst = strRst & "			</ImageURLList>"
		strRst = strRst & "		</shop:UploadImageRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		getNvClassImageRegXML = strRst
	End Function

	'// ��ǰ��� XML ����
	Public Function getNvClassItemRegXML(oServ, oOper, isEdit)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1np6kl_01"
		End If
		Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, oOper)

		strRst = ""
		strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
		strRst = strRst & "	<soapenv:Header/>"
   		strRst = strRst & "	<soapenv:Body>"
		strRst = strRst & "		<shop:ManageProductRequest>"
		strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
		strRst = strRst & "			<shop:AccessCredentials>"
		strRst = strRst & "				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
		strRst = strRst & "				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
		strRst = strRst & "				<shop:Signature>"&osignature&"</shop:Signature>"
		strRst = strRst & "			</shop:AccessCredentials>"
		strRst = strRst & "			<shop:Version>2.0</shop:Version>"
		strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
		strRst = strRst & "			<Product>"
		If isEdit = "Y" Then
			strRst = strRst & "			<shop:ProductId>"&FNvClassGoodNo&"</shop:ProductId>"		'��ǰID | ������ ���, ������ ����
		End If
		strRst = strRst & "				<shop:StatusType>SALE</shop:StatusType>"			'# ��ǰ�ǸŻ��� | ����� SALE(�Ǹ���)�� �Է�, ������ SALE, SUSP(�Ǹ� ����)�� �Է�, StockQuantity�� 0 �̸� OSTK(ǰ��)�� �����
		strRst = strRst & "				<shop:SaleType>NEW</shop:SaleType>"					'��ǰ �Ǹ� ����..���Է½� NEW�� ����
		strRst = strRst & getzCostomMadeInd													'#�ֹ� ���� ��ǰ ����

		''test�Դϴ� #######################################################################
		If FItemid = "2525634" Then
			strRst = strRst & "				<shop:CategoryId>50007215</shop:CategoryId>"		'#Leaf ī�װ� | ID ��ǰ��Ͻ� �ʼ� | ModelType�� �𵨸�ID�� �Էµ� ��� �ش� �𵨸� ID�� ���ε�  Leaf ī�װ� ID�� �����ϸ� ��û���� ���޵� CategoryId�� ���õȴ�
		Else
			strRst = strRst & "				<shop:CategoryId>50007332</shop:CategoryId>"		'#Leaf ī�װ� | ID ��ǰ��Ͻ� �ʼ� | ModelType�� �𵨸�ID�� �Էµ� ��� �ش� �𵨸� ID�� ���ε�  Leaf ī�װ� ID�� �����ϸ� ��û���� ���޵� CategoryId�� ���õȴ�
		End If

'		strRst = strRst & "				<shop:LayoutType></shop:LayoutType>"				'��ǰ �� ���̾ƿ� Ÿ�� �ڵ� | ���� �ڵ� ��ǰ �� ���̾ƿ� Ÿ�� : �ڵ� ���Է� �� �������� (BASIC)���� ����ȴ�
		strRst = strRst & "				<shop:Name><![CDATA["&getItemNameFormat&"]]></shop:Name>"			'#��ǰ��
'		strRst = strRst & "				<shop:PublicityPhraseContent></shop:PublicityPhraseContent>"		'ȫ�� ����
'		strRst = strRst & "				<shop:PublicityPhraseStartDate></shop:PublicityPhraseStartDate>"	'ȫ�� ���� ���� ������
'		strRst = strRst & "				<shop:PublicityPhraseEndDate></shop:PublicityPhraseEndDate>"		'ȫ�� ���� ���� ������
		strRst = strRst & "				<shop:SellerManagementCode>"&FItemid&"</shop:SellerManagementCode>"	'�Ǹ��� ��ǰ �ڵ�
'		strRst = strRst & "				<shop:SellerBarCode></shop:SellerBarCode>"							'�Ǹ��� ���ڵ�
		strRst = strRst & "				<shop:Model>"	'�� ����| �� ID ������ ���� ��� �귣���, ������� ���� ����..���������� "�����ű����� ��������/���յ��/�������� �����ǰ ��������/����Ȯ��/���������ռ�Ȯ�� �� ��� �ʼ�, �������(ManufacturerName), �귣���(BrandName),�𵨸�(ModelName)�� �ʼ��� �Է�
'		strRst = strRst & "					<shop:Id></shop:Id>"									'�� ID
		strRst = strRst & "					<shop:ManufacturerName><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"��ǰ���� ����",Fmakername)&"]]></shop:ManufacturerName>"		'�������
		strRst = strRst & "					<shop:BrandName><![CDATA["&chkIIF(trim(FSocname_kor)="" or isNull(FSocname_kor),"��ǰ���� ����",FSocname_kor)&"]]></shop:BrandName>"				'�귣���
		strRst = strRst & "				</shop:Model>"
'		strRst = strRst & "				<shop:AttributeValueList></shop:AttributeValueList>"		' ,�� �и��� �Ӽ��� ��� | ����� ������� ������ ���� ��� ����
		strRst = strRst & getOriginAreaType															'#������ ����
'		strRst = strRst & "				<shop:ManufactureDate></shop:ManufactureDate>"				'���� ���� | YYYY-MM-DD ����
'		strRst = strRst & "				<shop:ValidDate></shop:ValidDate>"							'��ȿ ���� | YYYY-MM-DD ����
		strRst = strRst & "				<shop:TaxType>"&CHKIIF(FVatInclude="N","DUTYFREE","TAX")&"</shop:TaxType>"	'#�ΰ��� | ���� : TAX, �鼼 : DUTYFREE, ���� : SMALL
		strRst = strRst & "				<shop:MinorPurchasable>Y</shop:MinorPurchasable>"			'#�̼����� ���� ���� ���� Y or N
		strRst = strRst & getImageType																'#�̹��� ����
		strRst = strRst & "				<shop:DetailContent><![CDATA["&getNvClassItemContParamToReg&"]]></shop:DetailContent>"		'#��ǰ �� ����
'		strRst = strRst & "				<shop:SellerNoticeId></shop:SellerNoticeId>"										'�������� ��ȣ
'		strRst = strRst & "				<shop:PurchaseReviewExposure></shop:PurchaseReviewExposure>"						'������ ���� ���� | Y or N, ������ ���� ���� ���� ī�װ��� ��쿡�� ��ȿ�ϸ� �� �ܿ��� Y�� �����ȴ�. ���Է� �� Y�� �����
'		strRst = strRst & "				<shop:RegularCustomerExclusiveProduct></shop:RegularCustomerExclusiveProduct>"		'�ܰ� ȸ�� ���� ��ǰ ���� | Y or N ���Է½� N���� �����
'		strRst = strRst & "				<shop:KnowledgeShoppingProductRegistration></shop:KnowledgeShoppingProductRegistration>"	'���̹� ���� ��� | Y or N ���̹� �����ְ� �ƴ� ��� N���� �����
'		strRst = strRst & "				<shop:GalleryId></shop:GalleryId>"							'������ ��ȣ
'		strRst = strRst & "				<shop:SaleStartDate></shop:SaleStartDate>"					'�Ǹ� ������ | YYYY-MM-DD ����..��¥������ �Է��ϴ� ��� �ڵ����� 0��0���� �ٿ��� �����.�Žð� 00�����θ� ���� ����
'		strRst = strRst & "				<shop:SaleEndDate></shop:SaleEndDate>"						'�Ǹ� ������ | YYYY-MM-DD HH:mm����..��¥������ �Է��ϴ� ��� �ڵ����� 23�� 59���� �ٿ��� �����.�Žð� 59�����θ� ���� ����
		strRst = strRst & "				<shop:SalePrice>"&Clng(GetRaiseValue(MustPrice/10)*10)&"</shop:SalePrice>"		'#�ǸŰ�
		If (isEdit = "Y")  Then
			If (Foptioncnt = 0) Then
				strRst = strRst & "				<shop:StockQuantity>"&getLimitNvClassEa&"</shop:StockQuantity>"		'#��� ���� | ��ǰ��Ͻ� �ʼ�, ��ǰ ������ ��� ������ �Է����� ������ ������� DB�� ����� ���� ����� ������ �ʴ´�. ������ ��� ���� 0���� �ԷµǸ� StatusType���� ���޵� �׸��� ���õǸ� ��ǰ ���´� OSTK(ǰ��)�� �����
			End If
		Else
			strRst = strRst & "				<shop:StockQuantity>"&getLimitNvClassEa&"</shop:StockQuantity>"		'#��� ���� | ��ǰ��Ͻ� �ʼ�, ��ǰ ������ ��� ������ �Է����� ������ ������� DB�� ����� ���� ����� ������ �ʴ´�. ������ ��� ���� 0���� �ԷµǸ� StatusType���� ���޵� �׸��� ���õǸ� ��ǰ ���´� OSTK(ǰ��)�� �����
		End If
'		strRst = strRst & "				<shop:MinPurchaseQuantity></shop:MinPurchaseQuantity>"					'�ּ� ���� ����
'		strRst = strRst & "				<shop:MaxPurchaseQuantityPerId></shop:MaxPurchaseQuantityPerId>"		'1�� �ִ� ���� ����
'		strRst = strRst & "				<shop:MaxPurchaseQuantityPerOrder></shop:MaxPurchaseQuantityPerOrder>"	'1ȸ �ִ� ���� ����
'		strRst = strRst & "				<shop:SellerDiscount>"									'�Ǹ��� ��� ���� | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#PC ��� ���ξ�/������ | PC���θ� �����Ϸ��� MobileAmount���� 0�� �Է�..������(%, ����)�� ���� ������ ���е�..ex)���� 10%�̸� ������, 1000�̸� ���ξ��� ��Ÿ����
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'PC ��� ���� ������ | YYYY-MM-DD HH:mm ����..��¥������ �Է��ϴ� ��� �ڵ����� 0��0���� �ٿ��� �����.�Žð� 00, 10, 20, 30, 40, 50�����θ� ���� ����
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'PC ��� ���� ������ | YYYY-MM-DD HH:mm ����..��¥������ �Է��ϴ� ��� 23�� 59���� �ٿ��� �����..�Žð� 09, 19, 29, 39, 49, 59�����θ� ���� ����
'		strRst = strRst & "					<shop:MobileAmount></shop:MobileAmount>"			'#����� ��� ���ξ�/������ | ����� ���θ� �����Ϸ��� Amount�� 0�� �Է�..������(%, ����)�� ���� ������ ���е�..ex)���� 10%�̸� ������, 1000�̸� ���ξ��� ��Ÿ����
'		strRst = strRst & "					<shop:MobileStartDate></shop:MobileStartDate>"		'����� ��� ���� ������ | YYYY-MM-DD HH:mm ����..��¥������ �Է��ϴ� ��� �ڵ����� 0��0���� �ٿ��� �����.�Žð� 00, 10, 20, 30, 40, 50�����θ� ���� ����
'		strRst = strRst & "					<shop:MobileEndDate></shop:MobileEndDate>"			'����� ��� ���� ������ | YYYY-MM-DD HH:mm ����..��¥������ �Է��ϴ� ��� 23�� 59���� �ٿ��� �����..�Žð� 09, 19, 29, 39, 49, 59�����θ� ���� ����
'		strRst = strRst & "				</shop:SellerDiscount>"
'		strRst = strRst & "				<shop:MultiPurchaseDiscount>"							'���� ���� ���� | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#���� ���� ���ξ�/������ | ������(%, ����)�� ���� ������ ���е�..ex)���� 10%�̸� ������, 1000�̸� ���ξ��� ��Ÿ����
'		strRst = strRst & "					<shop:OrderAmount></shop:OrderAmount>"				'#���� ���� ���� ���� �ݾ�/���� | ������(��, ����)�� ���� ���� ����..ex)���� 10���̸� ����, 1000�̸� �ݾ��� ��Ÿ����
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'���� ���� ���� ������ | YYYY-MM-DD ����
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'���� ���� ���� ������ | YYYY-MM-DD ����..�������� �Է��� ��� �ʼ�
'		strRst = strRst & "				</shop:MultiPurchaseDiscount>"
'		strRst = strRst & "				<shop:Mileage>"											'��ǰ ���Ž� �����Ǵ� ���̹����� ����Ʈ | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#���̹����� ����Ʈ ������/������ | ������(%, ����)�� ���� ������ ���е�..ex)���� 10%�̸� ������, 1000�̸� ���ξ��� ��Ÿ����
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'���̹����� ����Ʈ ��ȿ �Ⱓ ������..YYYY-MM-DD ����
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'���̹����� ����Ʈ ��ȿ �Ⱓ ������..YYYY-MM-DD ����, �������� �Է��� ��� �ʼ�
'		strRst = strRst & "				</shop:Mileage>"
'		strRst = strRst & "				<shop:ReviewPoint>"												'������ �ۼ� �� �����Ǵ� ���̹����� ����Ʈ | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:PurchaseReviewPoint></shop:PurchaseReviewPoint>"		'������ �ۼ� �� �����Ǵ� ���̹����� ����Ʈ | ������, �����̾� ������ �� �� �ϳ��� �ʼ� �Է�
'		strRst = strRst & "					<shop:PremiumReviewPoint></shop:PremiumReviewPoint>"		'�����̾� ������ �ۼ� �� �����Ǵ� ���̹����� ����Ʈ | ������, �����̾� ������ �� �� �ϳ��� �ʼ� �Է�
'		strRst = strRst & "					<shop:RegularCustomerPoint></shop:RegularCustomerPoint>"	'�ܰ� ȸ���� �������̳� �����̾� ������ �ۼ� �� �߰� �����Ǵ� ���̹����� ����Ʈ | �������̳� �����̾� �������� �ִ� ��쿡�� �Է�
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"							'���̹����� ����Ʈ ��ȿ �Ⱓ ������ | YYYY-MM-DD ����
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"								'���̹����� ����Ʈ ��ȿ �Ⱓ ������ | YYYY-MM-DD ����, �������� �Է��� ��� �ʼ�
'		strRst = strRst & "				</shop:ReviewPoint>"
'		strRst = strRst & "				<shop:FreeInterest>"								'������ �Һ� | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:Month></shop:Month>"						'#������ �Һ� ���� ��
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"				'������ �Һ� ������ | YYYY-MM-DD ����
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"					'������ �Һ� ������ | YYYY-MM-DD ����, �������� �Է��� ��� �ʼ�
'		strRst = strRst & "				</shop:FreeInterest>"
'		strRst = strRst & "				<shop:Gift>"										'����ǰ | �����̳�, �Է��� ��� �ϴ� #�� �ʼ�
'		strRst = strRst & "					<shop:Name></shop:Name>"						'#����ǰ
'		strRst = strRst & "				</shop:Gift>"

		''test�Դϴ� #######################################################################
		If Fitemid = "2525634" Then
		Else
			strRst = strRst & getECouponType
		End If
'		strRst = strRst & "				<shop:PurchaseApplicationUrl></shop:PurchaseApplicationUrl>"	'�޴��� ���Ž�û�� URL | �޴��� ī�װ� ��ǰ�� ��� �ʼ�
'		strRst = strRst & "				<shop:CellPhonePrice></shop:CellPhonePrice>"					'���δ� �޴��� �ܸ��� ��� | �޴��� ī�װ� ��ǰ�� ��� �ʼ�
'		strRst = strRst & "				<shop:WifiOnly></shop:WifiOnly>"		'Wifi ���� ��ǰ ���� | Y or N..�º� ī�װ� ��ǰ�� ��� �ʼ�..Y �Է½� PurchaseApplicationUrl, CellPhonePrice �ԷºҰ�..N �Է½� PurchaseApplicationUrl, CellPhonePrice �Է� �ʼ�
		strRst = strRst & "				<shop:ProductSummary>"					'��ǰ ��� ���� | ��ǰ ��Ͻ� �ʼ�, ��ǰ ���� �ÿ��� ������ ��ǰ ��� ������ �Էµ� ��쿡�� ������ �� �ִ�. �� ��� ������ ����� ��ǰ ��� ���� ���� �����ȴ�.

		''test�Դϴ� #######################################################################
		If Fitemid = "2525634" Then
			strRst = strRst & getNvClassItemInfoCdToRegOnlyMobile
		Else
			strRst = strRst & getNvClassItemInfoCdToReg
		End If
		strRst = strRst & "				</shop:ProductSummary>"
		strRst = strRst & getSellerComment
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</shop:ManageProductRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		''test�Դϴ� #######################################################################
		If Fitemid = "2525634" Then
			response.write strRst
		End If
'response.end
		getNvClassItemRegXML = strRst
	End Function
End Class

Class CNvClass
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
	Public FRectGubun

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

	Public Sub getNvClassNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'�ɼ� ��ü ǰ���� ��� ��� �Ұ�.
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
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.nvClassStatCD,-9) as nvClassStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		'strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName, isnull(ca.needCert, '') as needCert "
		strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName "
		strSql = strSql & "	, isnull(R.APIaddImg, '') as APIaddImg "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & " and i.isusing='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.itemdiv in ('08') "	'Ƽ��/Ŭ���� ��ǰ
		strSql = strSql & " and (i.cate_large = '035' and i.cate_mid = '022' and i.cate_small = '010') " '����/��� > ���/���� > Ŭ���� �� ��ϵǾ� ��
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstorefarmclass') "
		If FRectGubun <> "IMG" Then
			strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] where nvClassStatCD > 3) "
		End If
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvClassItem
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
                FOneItem.FNvClassStatCD		= rsget("nvClassStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FSocname_kor		= rsget("socname_kor")
                FOneItem.FAPIaddImg			= rsget("APIaddImg")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
                FOneItem.FRegImageName 		= rsget("regImageName")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                'FOneItem.FNeedCert 			= rsget("needCert")
		End If
		rsget.Close
	End Sub

	Public Sub getNvClassEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(m.nvClassGoodNo, '') as nvClassGoodNo, m.nvClassprice, m.nvClassSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, isnull(m.APIaddImg, '') as APIaddImg "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
    	strSql = strSql & "	,(CASE WHEN i.isusing = 'N' "
		strSql = strSql & "		or i.sellyn <> 'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv <> '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & " 	or exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstorefarmclass') "
		strSql = strSql & " 	or i.cate_large + i.cate_mid + i.cate_small <> '035022010' "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIaddImg = 'Y' "
		strSql = strSql & " and m.nvClassStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.nvClassGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvClassItem
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FNvClassGoodNo		= rsget("nvClassGoodNo")
				FOneItem.FNvClassprice		= rsget("nvClassprice")
				FOneItem.FNvClassSellyn		= rsget("nvClassSellYn")

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv			= rsget("infoDiv")
	            FOneItem.Fsafetyyn			= rsget("safetyyn")
	            FOneItem.FsafetyDiv			= rsget("safetyDiv")
	            FOneItem.FsafetyNum			= rsget("safetyNum")
	            FOneItem.FmaySoldOut		= rsget("maySoldOut")
	            FOneItem.Fregitemname		= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
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

Function getNvClassGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 nvClassGoodNo FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getNvClassGoodNo = rsget("nvClassGoodNo")
	rsget.Close
End Function

%>
