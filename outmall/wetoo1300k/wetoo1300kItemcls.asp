<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "wetoo1300k"
CONST CUPJODLVVALID = TRUE			''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5				'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST CDEFALUT_STOCK = 9999

Class CWetoo1300kItem
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
	Public FbasicImage600
	Public FbasicImage600str
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
	Public FWetoo1300kStatCD
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
	Public FWetoo1300kGoodNo
	Public FWetoo1300kPrice
	Public FBrandCode
	Public FWetoo1300kSellYn
	Public FLarge_category
	Public FMiddle_category
	Public FSmall_category
	Public FDetail_category
	Public FAdultType
	Public FLastStatCheckDate
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999999" Then
			getOrderMaxNum = 999999
		End If
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
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin
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

		If specialPrice <> "" Then
			tmpPrice = specialPrice
		Else
			If outmallstandardMargin = "" Then
				outmallstandardMargin	= FOutmallstandardMargin
			End If
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)

			If FWetoo1300kPrice = 0 Then
				If (GetTenTenMargin < outmallstandardMargin) Then
					tmpPrice = Forgprice
				Else
					tmpPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < outmallstandardMargin Then
					If (Forgprice < Round(FWetoo1300kPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FWetoo1300kPrice * 0.35, 0)/10)*10)
					ElseIf Clng(Forgprice) > Clng(Round(FWetoo1300kPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FWetoo1300kPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = Forgprice
					End If
				Else
					If (FSellCash < Round(FWetoo1300kPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FWetoo1300kPrice * 0.35, 0)/10)*10)
					ElseIf Clng(FSellCash) > Clng(Round(FWetoo1300kPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FWetoo1300kPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'// Wetoo1300k �Ǹſ��� ��ȯ
	Public Function getWetoo1300kSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getWetoo1300kSellYn = "Y"
			Else
				getWetoo1300kSellYn = "N"
			End If
		Else
			getWetoo1300kSellYn = "N"
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
        buf = LeftB(buf, 100)
        getItemNameFormat = buf
    end function

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

	Public Function getKeywords()
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
			retKeyword = LeftB(arrRst(0), 20) &";"&LeftB(arrRst(1), 20) &";"& LeftB(arrRst(2), 20) &";"& LeftB(arrRst(3), 20) &";"& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &";"
			Next
			If Right(keyword1,1) = ";" Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			retKeyword = keyword1
		End If
		getKeywords = retKeyword
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

	Public Function getMadeParam()
		If FItemdiv = "06" OR FItemdiv = "16" Then
			getMadeParam = "13"
		Else
			getMadeParam = "0"
		End If
	End Function

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

    Public Function getCertOrganName(isafetyDiv, icertOrganName)
		Select Case isafetyDiv
			Case "11"
				If Instr(icertOrganName, "�ѱ�ȭ�н��迬����") > 0 Then
					getCertOrganName = "11"
				ElseIf Instr(icertOrganName, "�ѱ��Ƿ����迬����") > 0 Then
					getCertOrganName = "12"
				ElseIf Instr(icertOrganName, "�ѱ�����������ڽ��迬����") > 0 Then
					getCertOrganName = "13"
				ElseIf Instr(icertOrganName, "�ѱ��Ǽ���Ȱ���迬����") > 0 Then
					getCertOrganName = "14"
				ElseIf Instr(icertOrganName, "�ѱ�FITI���迬����") > 0 Then
					getCertOrganName = "15"
				ElseIf Instr(icertOrganName, "�ѱ����������迬����") > 0 Then
					getCertOrganName = "16"
				ElseIf Instr(icertOrganName, "KOTITI���迬����") > 0 Then
					getCertOrganName = "17"
				Else
					getCertOrganName = "11"
				End If
			Case "12"
				If Instr(icertOrganName, "�ѱ�ȭ�����ս��迬����") > 0 Then
					getCertOrganName = "11"
				ElseIf Instr(icertOrganName, "�ѱ��Ƿ����迬����") > 0 Then
					getCertOrganName = "12"
				ElseIf Instr(icertOrganName, "�ѱ�����������ڽ��迬����") > 0 Then
					getCertOrganName = "13"
				ElseIf Instr(icertOrganName, "�ѱ��Ǽ���Ȱ���迬����") > 0 Then
					getCertOrganName = "14"
				ElseIf Instr(icertOrganName, "�ѱ�FITI���迬����") > 0 Then
					getCertOrganName = "15"
				ElseIf Instr(icertOrganName, "�ѱ����������迬����") > 0 Then
					getCertOrganName = "16"
				ElseIf Instr(icertOrganName, "KOTITI���迬����") > 0 Then
					getCertOrganName = "17"
				Else
					getCertOrganName = "11"
				End If
			Case "13"
				If Instr(icertOrganName, "�ѱ�����������ڽ��迬����") > 0 Then
					getCertOrganName = "11"
				ElseIf Instr(icertOrganName, "�ѱ�������������") > 0 Then
					getCertOrganName = "13"
				ElseIf Instr(icertOrganName, "�ѱ�ȭ�����ս��迬����") > 0 Then
					getCertOrganName = "14"
				Else
					getCertOrganName = "11"
				End If
			Case "14"
				If Instr(icertOrganName, "�ѱ�����������ڽ��迬����") > 0 Then
					getCertOrganName = "11"
				ElseIf Instr(icertOrganName, "�ѱ�ȭ�����ս��迬����") > 0 Then
					getCertOrganName = "12"
				ElseIf Instr(icertOrganName, "�ѱ�������������") > 0 Then
					getCertOrganName = "13"
				Else
					getCertOrganName = "11"
				End If
		End Select
    End function
	
	'��ǰ �̹��� �Ķ���� ����
	Public Function getWetoo1300kImageParamToReg(obj)
		Dim strSQL, fImage, addImgUrl
		If NOT(isnull(FbasicImage600)) and NOT(FbasicImage600 = "") Then
			fImage = FbasicImage600str
		Else
			fImage = FbasicImage
		End If

		Set obj("product")("image") = jsObject()
			obj("product")("image")("image_url1") = fImage		'�̹��� URL | 640*640 JPG
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				addImgUrl = ""
				If (NOT(IsNULL(rsget("addimage_600")))) AND (rsget("addimage_600") <> "") AND (Len(rsget("addimage_600"))) > 0 Then
					addImgUrl = "add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")
				Else
					addImgUrl = "add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				End If

				If rsget("imgType") = "0" Then
					obj("product")("image")("image_url"&i+1&"") = "http://webimage.10x10.co.kr/image/"& addImgUrl	'�̹��� URL | 640*640 JPG
				End If
				rsget.MoveNext
				If i>=5 Then Exit For
			Next
		End If
		rsget.Close
	End Function

	'��ǰ �������� �Ķ���� ����
	Public Function getWetoo1300kSafetyParamToReg(obj)
		Dim strSql, safetyDiv, certNum, certOrganName, modelName, certDate, isRegCert
		Dim safefytypecode, safetyCertNo

		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isNull(t.certNum, '') as certNum, isNull(f.modelName, '') as modelName, isNull(f.certDate, '') as certDate, isNull(f.certOrganName, '') as certOrganName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
			certOrganName = rsget("certOrganName")
			modelName	= rsget("modelName")
			certDate	= rsget("certDate")
			isRegCert	= "Y"
		Else
			isRegCert	= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			Select Case safetyDiv
				Case "10"
					safefytypecode			= "13"
					safetyCertNo			= certNum
				Case "20"
					safefytypecode			= "14"
					safetyCertNo			= certNum
				Case "40"
					safefytypecode			= "11"
					safetyCertNo			= certNum
				Case "50"
					safefytypecode			= "12"
					safetyCertNo			= certNum
				Case "70"	'������� �Ⱥ���, ��Ȱ�� ��ü
					safefytypecode			= "11"
					safetyCertNo			= certNum
				Case "80"	'������� �Ⱥ���, ��Ȱ�� ��ü
					safefytypecode			= "12"
					safetyCertNo			= certNum
			End Select

			If len(certDate) = 8 Then
				certDate = Left(certDate,4)&"-"&Mid(certDate,5,2)&"-"&Mid(certDate,7,2)
			Else
				certDate = ""
			End If
			certOrganName = getCertOrganName(safefytypecode, certOrganName)

			Set obj("product")("safefy") = jsObject()
				obj("product")("safefy")("safefy_type_code") = safefytypecode	'������ | �������� ������
				obj("product")("safefy")("safefy_center_code") = certOrganName	'������� | �������� ������
				obj("product")("safefy")("safefy_no") = safetyCertNo			'�����ڵ�
				obj("product")("safefy")("safefy_model") = modelName			'������
				obj("product")("safefy")("safefy_date") = certDate				'�������� | YYYY-MM-DD
				obj("product")("safefy")("safefy_memo") = ""					'�޸�
		End If
	End Function

	'��ǰ �ɼ� �Ķ���� ����
	Public Function getWetoo1300kOptParamToReg(obj)
		Dim buf, isOptSoldout, lp
		Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, optDc1, optDc2
		Dim optNm, optDc, optLimit, itemoption, MultiYN
		Dim arrOptValue, arrOptmixlist
		obj("product")("option_use") = "Y"												'�ɼǻ�뿩�� | Y:�ɼǻ�� N:�ɼ� �̻��
		Set obj("product")("option") = jsObject()
			obj("product")("option")("option_mix") = "Y"
			obj("product")("option")("option_level") = "1"
			If FOptionCnt = 0 Then			'��ǰ
				obj("product")("option")("option_title1") = "�ɼ�"										'�ɼǸ�^��뿩��^��Ī|�ɼǸ�^��뿩��^��Ī
				obj("product")("option")("option_value1") = "���ϻ�ǰ^Y^0000"							'�ɼǸ�^��뿩��^��Ī|�ɼǸ�^��뿩��^��Ī
				obj("product")("option")("option_mix_list") = "���ϻ�ǰ^Y^0000^0^"&getLimitEa()&"^T"	'�ɼǸ�^��뿩��^��Ī^����^���^�Ǹ�����|�ɼǸ�^��뿩��^��Ī^����^���^�Ǹ�����
			Else
				strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
				strSql = strSql & " From [db_item].[dbo].tbl_item_option "
				strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				arrOptValue = ""
				arrOptmixlist = ""
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						lp = lp + 1
						optLimit = rsget("optLimit")
						optLimit = optLimit-5
						If (optLimit < 1) Then optLimit = 0
						If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
						itemoption	= rsget("itemoption")
						optDc		= db2Html(rsget("optionname"))
						optIsusing	= rsget("isusing")
						optSellYn	= rsget("optsellyn")
						optaddprice	= rsget("optaddprice")
'						optaddprice = MustPrice() + rsget("optaddprice")

						If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
							isOptSoldout = "N"
						Else
							isOptSoldout = "Y"
						End If
						optDc = Replace(optDc, ":", "")
						optDc = Replace(optDc, "|", "")
						optDc = Replace(optDc, "^", "")

						arrOptValue = arrOptValue & optDc & "^" & isOptSoldout & "^" & itemoption & "|"
						arrOptmixlist = arrOptmixlist & optDc & "^" & isOptSoldout & "^" & itemoption & "^" & optaddprice & "^" & optLimit & "^" & "T" & "|"

						rsget.MoveNext
					Loop
				End If
				rsget.Close

				If Right(arrOptValue,1) = "|" Then
					arrOptValue = Left(arrOptValue, Len(arrOptValue) - 1)
				End If

				If Right(arrOptmixlist,1) = "|" Then
					arrOptmixlist = Left(arrOptmixlist, Len(arrOptmixlist) - 1)
				End If
				obj("product")("option")("option_title1") = "�ɼ�"
				obj("product")("option")("option_value1") = arrOptValue				'�ɼǸ�^��뿩��^��Ī|�ɼǸ�^��뿩��^��Ī
				obj("product")("option")("option_mix_list") = arrOptmixlist			'�ɼǸ�^��뿩��^��Ī^����^���^�Ǹ�����|�ɼǸ�^��뿩��^��Ī^����^���^�Ǹ�����
				' obj("product")("option")("option_title2") = ""
				' obj("product")("option")("option_value2") = ""
			End If
	End Function

	'��ǰ ������� �Ķ���� ����
	Public Function getWetoo1300kInfoCdParameter(obj)
		Dim strSql, i, mallinfoCd, infoContent
		strSql = ""
		strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Wetoo1300k_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			Set obj("product")("noti") = jsObject()
				obj("product")("noti")("noti_group") = rsget("mallinfoDiv")
				Set obj("product")("noti")("noti_info") = jsArray()
			i = 0
			Do until rsget.EOF
				mallinfoCd  = rsget("mallinfoCd")
				infoContent = rsget("infoContent")
				If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
					infoContent = replaceRst(replace(infoContent, chr(31), ""))
				End If

				Set obj("product")("noti")("noti_info")(i) = jsObject()
					obj("product")("noti")("noti_info")(i)("noti_code") = mallinfoCd
					obj("product")("noti")("noti_info")(i)("noti_value") = infoContent
				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
	End Function

	'��ǰ���� �Ķ���� ����
	Public Function getWetoo1300kContParamToReg()
		Dim strRst, strSQL, retContents, retOrderComment
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_wetoo1300k.jpg></p><br />")
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
		strRst = strRst & ("<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_wetoo1300k.jpg>")
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
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_wetoo1300k.jpg></p><br />" & strRst & "<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_wetoo1300k.jpg></div>"
			retContents = strRst
		End If
		rsget.Close

		getWetoo1300kContParamToReg = retContents
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

	'��ǰ��� Json
	Public Function getWetoo1300kItemRegParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj
		If application("Svr_Info") = "Dev" Then
			FBrandCode = "001"
		End If
'http://localhost:11117/outmall/wetoo1300k/wetoo1300kActProc.asp?act=REG&itemid=2937363
		Set obj = jsObject()
			Set obj("header") = jsObject()
				obj("header")("company_code") = company_code									'��ü�ڵ�
				obj("header")("company_auth") = company_auth									'�����ڵ�
				Set obj("product") = jsObject()
					obj("product")("product_name") = getItemNameFormat()						'#��ǰ��
					obj("product")("prefix") = ""												'��ǰ �Ӹ���
					obj("product")("category") = FLarge_category & "-" & FMiddle_category & "-" & FSmall_category & "-" & FDetail_category	'#ī�װ� | ��-��-��-��
					obj("product")("product_desc") = getWetoo1300kContParamToReg()				'#��ǰ�� | ��ǰ��������
					obj("product")("product_type") = getMadeParam()								'#��ǰ���� | 0:�Ϲݻ�ǰ 13:�ֹ����ۻ�ǰ
					obj("product")("company_product_code") = FItemid							'#��ü��ǰ�ڵ�
					obj("product")("company_code") = company_code								'#���۴��ڵ�
					obj("product")("brand_code") = FBrandCode									'#�귣���ڵ�
					obj("product")("origin_place") = GetSourcearea()							'#������
					obj("product")("maker") = CStr(FMakerName)									'#������
					obj("product")("model") = ""												'#�𵨸�
					obj("product")("standard") = CStr(Fitemsize)								'#�԰�
					obj("product")("meterial") = CStr(Fitemsource)								'#���
					obj("product")("color") = "000"												'�÷��ڵ� | ������ : 000
					obj("product")("keyword") = getKeywords()									'Ű���� | ��)���̾;�����췯;����  @�귣���, ��ǰ���� �ڵ����� Ű���� ����
					obj("product")("sale_price") = Forgprice									'#�Ǹűݾ�
					obj("product")("dc_price") = Forgprice - MustPrice()						'#���αݾ� | �ǸŰ� = sale_price-dc_price, ���ΰ������� 0���� ����
					obj("product")("supply_price") = Clng(MustPrice()*0.88)						'#���ް�
					obj("product")("supply_type") = "0"											'#���޹�� | 1:���� 0:��Ź
					obj("product")("display") = "Y"												'�������� | Y:���� N:�����ߴ� (MD ���� ������ ������ N���� ����ǰ� ��. ���� �� ���ʹ� ��������)
					obj("product")("tax_yn") = CHKIIF(FVatInclude="N","0","1")					'�ΰ������� | 1: ������ǰ 0:�鼼��ǰ
					obj("product")("sale_type") = "Y"											'#�Ǹ����� | Y:�ǸŰ��� N:��������ǰ  P:�Ǹ�����(������ display�� �����ǰ�  �ǸŴ� ��������)
					obj("product")("soldout_type") = ""											'1:�Ͻ�ǰ�� 2:�Ǹ�����  @sale_type ���� N�� ��쿡 �ش�
					obj("product")("stock_cnt") = getLimitEa()									'��� | ������
					obj("product")("delivery_type") = "X"										'#������� | Y:��ü������ X:��ü������ W:��ü���ҹ��
					obj("product")("delivery_agency_code") = "D020"								'#�ù���ڵ� | �ϴ� ����(D020)���� ����
					obj("product")("adult_check") = IsAdultItem()								'���������ʿ� | Y:�������� �ʿ� N:���ʿ�
					Call getWetoo1300kImageParamToReg(obj)
					Call getWetoo1300kSafetyParamToReg(obj)
					Call getWetoo1300kOptParamToReg(obj)
					Call getWetoo1300kInfoCdParameter(obj)
				getWetoo1300kItemRegParameter = obj.jsString
		Set obj = nothing
	End Function

	'��ǰ��� Json
	Public Function getWetoo1300kItemEditParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj
		If application("Svr_Info") = "Dev" Then
			FBrandCode = "001"
		End If
'http://localhost:11117/outmall/wetoo1300k/wetoo1300kActProc.asp?act=EDIT&itemid=2937363
		Set obj = jsObject()
			Set obj("header") = jsObject()
				obj("header")("company_code") = company_code									'��ü�ڵ�
				obj("header")("company_auth") = company_auth									'�����ڵ�
				Set obj("product") = jsObject()
					obj("product")("product_code") = FWetoo1300kGoodNo							'#��ǰ�ڵ�
					obj("product")("product_name") = getItemNameFormat()						'#��ǰ��
					obj("product")("prefix") = ""												'��ǰ �Ӹ���
					obj("product")("category") = FLarge_category & "-" & FMiddle_category & "-" & FSmall_category & "-" & FDetail_category	'#ī�װ� | ��-��-��-��
					obj("product")("product_desc") = getWetoo1300kContParamToReg()				'#��ǰ�� | ��ǰ��������
					obj("product")("product_type") = getMadeParam()								'#��ǰ���� | 0:�Ϲݻ�ǰ 13:�ֹ����ۻ�ǰ
					obj("product")("company_product_code") = FItemid							'#��ü��ǰ�ڵ�
					obj("product")("origin_place") = GetSourcearea()							'#������
					obj("product")("maker") = CStr(FMakerName)									'#������
					obj("product")("model") = ""												'#�𵨸�
					obj("product")("standard") = CStr(Fitemsize)								'#�԰�
					obj("product")("meterial") = CStr(Fitemsource)								'#���
					obj("product")("color") = "000"												'�÷��ڵ� | ������ : 000
					obj("product")("keyword") = getKeywords()									'Ű���� | ��)���̾;�����췯;����  @�귣���, ��ǰ���� �ڵ����� Ű���� ����
					obj("product")("sale_price") = Forgprice									'#�Ǹűݾ�
					obj("product")("dc_price") = Forgprice - MustPrice()						'#���αݾ� | �ǸŰ� = sale_price-dc_price, ���ΰ������� 0���� ����
					obj("product")("supply_price") = Clng(MustPrice()*0.88)						'#���ް�
					obj("product")("supply_type") = "0"											'#���޹�� | 1:���� 0:��Ź
					obj("product")("display") = "Y"												'�������� | Y:���� N:�����ߴ� (MD ���� ������ ������ N���� ����ǰ� ��. ���� �� ���ʹ� ��������)
					obj("product")("tax_yn") = CHKIIF(FVatInclude="N","0","1")					'�ΰ������� | 1: ������ǰ 0:�鼼��ǰ
					obj("product")("sale_type") = "Y"											'#�Ǹ����� | Y:�ǸŰ��� N:��������ǰ  P:�Ǹ�����(������ display�� �����ǰ�  �ǸŴ� ��������)
					obj("product")("soldout_type") = ""											'1:�Ͻ�ǰ�� 2:�Ǹ�����  @sale_type ���� N�� ��쿡 �ش�
					obj("product")("stock_cnt") = getLimitEa()									'��� | ������
					obj("product")("delivery_type") = "X"										'#������� | Y:��ü������ X:��ü������ W:��ü���ҹ��
					obj("product")("delivery_agency_code") = "D020"								'#�ù���ڵ� | �ϴ� ����(D020)���� ����
					obj("product")("adult_check") = IsAdultItem()								'���������ʿ� | Y:�������� �ʿ� N:���ʿ�
					Call getWetoo1300kImageParamToReg(obj)
					Call getWetoo1300kSafetyParamToReg(obj)
					Call getWetoo1300kOptParamToReg(obj)
					Call getWetoo1300kInfoCdParameter(obj)
				getWetoo1300kItemEditParameter = obj.jsString
		Set obj = nothing
	End Function

	'��ǰ �ǸŻ��� ���� Json
	Public Function getWetoo1300kPriceSellynParameter(ichgSellYn)
		Dim strRst
		Dim obj, sale_type
		Select Case ichgSellYn
			Case "Y"	sale_type = "Y"		'�Ǹ���
			Case "N"	sale_type = "P"		'�Ǹ�����
			Case "X"	sale_type = "N"		'�Ǹ�����
		End Select

		Set obj = jsObject()
			Set obj("header") = jsObject()
				obj("header")("company_code") = company_code						'��ü�ڵ�
				obj("header")("company_auth") = company_auth						'�����ڵ�
				Set obj("product") = jsObject()
					obj("product")("product_code") = FWetoo1300kGoodNo				'��ǰ�ڵ�
					If application("Svr_Info") <> "Dev" Then
						obj("product")("display") = "Y"									'�������� | Y:���� N:�����ߴ� (MD ���� ������ ������ ��N�� ���� ����ǰ� ��. ���� �� ���ʹ� ��������)
					Else
						obj("product")("display") = "N"									'�������� | Y:���� N:�����ߴ� (MD ���� ������ ������ ��N�� ���� ����ǰ� ��. ���� �� ���ʹ� ��������)
					End If
					obj("product")("sale_type") = sale_type							'�Ǹ����� | Y:�ǸŰ��� N:��������ǰ  P:�Ǹ�����(������ display�� �����ǰ�  �ǸŴ� ��������)
					obj("product")("soldout_type") = CHKIIF(sale_type="N", "2", "")	'1:�Ͻ�ǰ�� 2:�Ǹ�����  @sale_type ���� N�� ��쿡 �ش�
					obj("product")("sale_price") = Forgprice						'�Ǹűݾ� | ������ �ǸŰ�
					obj("product")("dc_price") = Forgprice - MustPrice()			'���αݾ�
					obj("product")("supply_price") = Clng(MustPrice()*0.88)			'���ް�
		getWetoo1300kPriceSellynParameter = obj.jsString
	End Function
End Class

Class CWetoo1300k
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
	Public Sub getWetoo1300kNotRegOneItem
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
		strSql = strSql & "	, isNULL(R.wetoo1300kStatCD,-9) as wetoo1300kStatCD, isnull(R.wetoo1300kPrice, 0) as wetoo1300kPrice "
		strSql = strSql & "	, UC.socname_kor, am.large_category, am.middle_category, am.small_category, am.detail_category, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & "	, b.brandCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_wetoo1300k_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_category as tm on am.large_category = tm.large_category and am.middle_category = tm.middle_category and am.small_category = tm.small_category and am.detail_category = tm.detail_category "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_wetoo1300k_brandcode] as b on UC.userid = b.makerid "
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
		strSql = strSql & " and i.itemdiv in ('01', '16', '07') "		'01 : �Ϲ�, 16 : �ֹ�����, 07 : ��������
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
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'������� ī�װ�
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_wetoo1300k_regItem WHERE wetoo1300kStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CWetoo1300kItem
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
				FOneItem.FbasicImage600		= rsget("basicimage600")
				FOneItem.FbasicImage600str	= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage600")
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
				FOneItem.FWetoo1300kStatCD	= rsget("wetoo1300kStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FLarge_category	= rsget("large_category")
				FOneItem.FMiddle_category	= rsget("middle_category")
				FOneItem.FSmall_category	= rsget("small_category")
				FOneItem.FDetail_category	= rsget("detail_category")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin = rsget("outmallstandardMargin")
				FOneItem.FWetoo1300kPrice	= rsget("wetoo1300kPrice")
				FOneItem.FBrandCode			= rsget("brandCode")
		End If
		rsget.Close
	End Sub

	Public Sub getWetoo1300kNotEditOneItem
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
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, m.wetoo1300kGoodNo, m.wetoo1300kprice, m.wetoo1300kSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.large_category, am.middle_category, am.small_category, am.detail_category, isNULL(m.wetoo1300kStatCD,-9) as wetoo1300kStatCD, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
        strSql = strSql & "	, b.brandCode "
		strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & " 	or i.itemdiv not in ('01', '16', '07') "		'01 : �Ϲ�, 16 : �ֹ�����, 07 : ��������
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_wetoo1300k_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_wetoo1300k_brandcode] as b on UC.userid = b.makerid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_wetoo1300k_category as tm on am.large_category = tm.large_category and am.middle_category = tm.middle_category and am.small_category = tm.small_category and am.detail_category = tm.detail_category "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.wetoo1300kGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CWetoo1300kItem
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
				FOneItem.FbasicImage600		= rsget("basicimage600")
				FOneItem.FbasicImage600str	= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage600")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FWetoo1300kGoodNo	= rsget("wetoo1300kGoodNo")
				FOneItem.FWetoo1300kprice	= rsget("wetoo1300kprice")
				FOneItem.FWetoo1300kSellYn	= rsget("wetoo1300kSellYn")

				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If
				FOneItem.FItemsize 			= rsget("itemsize")
				FOneItem.FItemsource 		= rsget("itemsource")
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

				FOneItem.FLarge_category	= rsget("large_category")
				FOneItem.FMiddle_category	= rsget("middle_category")
				FOneItem.FSmall_category	= rsget("small_category")
				FOneItem.FDetail_category	= rsget("detail_category")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FWetoo1300kStatCD	= rsget("wetoo1300kStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
				FOneItem.FBrandCode			= rsget("brandCode")
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

function wetoo1300kAPIURL()
	If application("Svr_Info") = "Dev" Then
		wetoo1300kAPIURL = "https://ts.1300k.com"
	Else
		wetoo1300kAPIURL = "http://api.1300k.com"
	End If
end function

function company_auth()
	If application("Svr_Info") = "Dev" Then
		company_auth = "1ac6e7cd04fc587cc26722b1cbaaa75c"
	Else
		company_auth = "f91f60a59e32425e4f22c3d20cf4f7b7"
	End If
end function

function company_code()
	If application("Svr_Info") = "Dev" Then
		company_code = "C927"
	Else
		company_code = "C927"
	End If
end function
%>