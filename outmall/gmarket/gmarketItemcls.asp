<%
CONST CMAXMARGIN = 18
CONST CMALLNAME = "gmarket1010"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.
CONST gmarketAPIURL = "http://tpl.gmarket.co.kr"
CONST gmarketSSLAPIURL = "https://tpl.gmarket.co.kr"
CONST gmarketTicket = "0A2799EE6A1B65CC78DA96AA52C7546B2181855E48A0A31EDD4F3A77C3C61015856FE3DE5D7828B129A31AAD5914D7060556616D3AB7F2A84008A600C89F5953A0362065429900D0EB25CEBEA0E1CAF9E784FBC4F36E86608F2CF44B40113ADF"
CONST CDEFALUT_STOCK = 999
CONST CRETURNFEE = 3000
CONST MAKERNO = "100005224"	'��Ÿ // ���� ���� ���ɼ� ����

Class CGmarketItem
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
	Public FListimage
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FGmarketStatCD
	Public Fdeliverfixday
	Public Fdeliverytype
	Public FrequireMakeDay
	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FmaySoldOut
	Public FDisplayDate
	Public Fregitemname
	Public FregImageName
	Public FbasicImageNm
	Public FBrandCode
	Public Fsocname_kor
	Public FDepthCode
	Public FDepth4Code
	Public FReturnShippingFee
	Public Fcdmkey
	Public Fcddkey
	Public FGmarketGoodNo
	Public FG9GoodNo
	Public FGmarketprice
	Public FGmarketSellYn
	Public FAPIadditem
	Public FAPIaddopt

	Public FNotinCate
	Public FSafeAuthType
	Public FAuthItemTypeCode
	Public FIsChildrenCate
	Public FOverlap
	Public FAdultType
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 999
		End If
	End Function

	'// ǰ������
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
	end function

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice
		Dim ownItemCnt, outmallstandardMargin
		sqlStr = ""
		sqlStr = sqlStr & " SELECT mustPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as m "
		sqlStr = sqlStr & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " WHERE m.mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and m.itemid = '"& Fitemid &"' "
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
			If outmallstandardMargin = "" Then
				outmallstandardMargin	= FOutmallstandardMargin
			End If

			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < outmallstandardMargin Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
		End If
	End Function

	Public Function getFiftyUpDown()
		Dim strSql, zoptaddprice, tmpPrice
		If FOptionCnt = 0 Then
			getFiftyUpDown = "N"
		Else
			strSql = ""
			strSql = strSql &" SELECT Max(optaddprice) optaddprice "
			strSql = strSql &" FROM db_item.dbo.tbl_item_option "
			strSql = strSql &" WHERE itemid = '"&FItemid&"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				zoptaddprice = rsget("optaddprice")
			End If
			rsget.Close

			If zoptaddprice = 0 Then
				getFiftyUpDown = "N"
			Else
				tmpPrice = Clng(MustPrice / 2)
				If tmpPrice > zoptaddprice Then
					getFiftyUpDown = "N"
				Else
					getFiftyUpDown = "Y"
				End If
			End If
		End If
	End Function

	'// ������ �Ǹſ��� ��ȯ
	Public Function getGmarketSellYn()
		'�ǸŻ��� (10:�Ǹ�����, 20:ǰ��)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getGmarketSellYn = "Y"
			Else
				getGmarketSellYn = "N"
			End If
		Else
			getGmarketSellYn = "N"
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
			FItemName = "TEST��ǰ "&FItemName
		End If

		If Date() >="2017-03-10" and Date() <= "2017-03-12" Then
			Select Case FItemid
				Case "1625309"		FItemName = FItemName & " ��Ű Ŀ�ǵ帮��"
				Case "1569915"		FItemName = FItemName & " ��Ű ���ܸ���Ŀ"
				Case "1565223"		FItemName = FItemName & " �ٸ��� ����"
				Case "1523844"		FItemName = FItemName & " ���ۺ� �Ŀ�ġ"
				Case "1523843"		FItemName = FItemName & " ���ۺ� �Ŀ�ġ"
				Case "1523841"		FItemName = FItemName & " ���ۺ� ����� �Ŀ�ġ"
				Case "1523840"		FItemName = FItemName & " ���ۺ� ��Ƽ �Ŀ�ġ"
				Case "1523839"		FItemName = FItemName & " ���ۺ� ����ũ�� �Ŀ�ġ"
				Case "1523838"		FItemName = FItemName & " ���ۺ� �ȴ�"
				Case "1523836"		FItemName = FItemName & " ���ۺ� ���ĵ�"
				Case "1523835"		FItemName = FItemName & " ũ�� �۷���"
				Case "1523833"		FItemName = FItemName & " ���ۺ� ��ġŸ��"
				Case "1520151"		FItemName = FItemName & " ���ۺ� ����"
				Case "1520149"		FItemName = FItemName & " ���ۺ� ����"
				Case "1509355"		FItemName = FItemName & " �ٸ��� �ڽ���"
				Case "1488156"		FItemName = FItemName & " ��Ű ��ǻ��"
				Case "1488140"		FItemName = FItemName & " Ǫ�� ��ǻ��"
				Case "1473441"		FItemName = FItemName & " ũ�� �۷���"
				Case "1471075"		FItemName = FItemName & " �ٸ��� ����"
				Case "1471073"		FItemName = FItemName & " �ٸ��� ������ ũ����"
				Case "1422085"		FItemName = FItemName & " �ٸ��� ����"
				Case "1407891"		FItemName = FItemName & " �ٸ��� ����"
				Case "1405564"		FItemName = FItemName & " �ٸ��� Ƽ��Ǭ��ũ"
				Case "1405559"		FItemName = FItemName & " �ٸ��� Ƽ��"
			End Select
		End If
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","����")
        buf = replace(buf,"&","��")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
'        buf = LeftB(buf, 94)
		If fnStrLength(buf) >= 80 Then
			buf = chrbyte(buf,76,"")
		End If
        getItemNameFormat = buf
    end function

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

	Public Function checkItemContent()
		Dim strSql, chkRst, etcLinkStr, isVal
		isVal = "N"
		strSql = ""
		strSql = strSql & " SELECT itemid, mallid, linkgbn, textVal, 'Y' as isVal " & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSql = strSql & " where mallid in ('','gmarket1010') and linkgbn = 'contents' and itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			etcLinkStr	= rsget("textVal")
			isVal		= rsget("isVal")
		End If
		rsget.Close

		If Instr(LCase(etcLinkStr), "<iframe") > 0 Then
			checkItemContent = "Y"
		ElseIf isVal <> "Y" AND Instr(LCase(FItemcontent), "<iframe") > 0 Then
			checkItemContent = "Y"
		Else
			checkItemContent = "N"
		End If
	End Function

	'// ��ǰ���: ��ǰ���� �Ķ���� ����(��ǰ��Ͽ�)
	Public Function getGmarketItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 ������ ž �̹��� �߰�
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_gmarket.jpg></p>&#xA;")

		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & Fordercomment & "<br>"
		End If

		'#�⺻ ��ǰ����
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "&#xA;")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%"">&#xA;")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#�⺻ ��ǰ �����̹���
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%"">&#xA;")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%"">&#xA;")

		'#��� ���ǻ���
		strRst = strRst & ("&#xA;<img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_gmarket.jpg>")

		strRst = strRst & ("</div>")
		getGmarketItemContParamToReg = strRst

		''2013-06-10 ������ �߰�(�Ե�����ó�� ��ǰ�̹����� ��� ���ڳ����� ����)
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','gmarket1010') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = nl2br(rsget("textVal")) & "&#xA;"
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_gmarket.jpg></p>&#xA;" & strRst & "&#xA;<img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_gmarket.jpg>"
			getGmarketItemContParamToReg = strRst
		End If
		rsget.Close
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

	Public Function getGmarketAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		strRst = strRst & "				<ItemImage "
		strRst = strRst & "					DefaultImage="""&FbasicImage&""""			'#��ǰ �⺻ �̹��� URL | 600 �� 600 �̹��� ����( jpg �̹��� )
		strRst = strRst & "					LargeImage="""&FbasicImage&""""				'��ǰ ū �̹��� URL | 600 �� 600
		strRst = strRst & "					SmallImage="""&FListImage&""""				'��ǰ ���� �̹��� URL | 100 �� 100

		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "		AddImage"&i+1&"="""&"http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """"
				End If
				rsget.MoveNext
				If i>=2 Then Exit For
			Next
		End If
		rsget.Close

		strRst = strRst & "					 xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		getGmarketAddImageParam = strRst
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

	Function getiszeroWonSoldOut(iitemid)
		Dim sqlStr, i, goptlimitno, goptlimitsold, cnt
		i = 0
		If Flimityn = "Y" Then
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
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getGmarketShippingParam()
		Dim strRst, NewGroupYn, GroupCode
		NewGroupYn = True
		If Not(NewGroupYn) Then
			strRst = strRst & "				<Shipping "
			strRst = strRst & "					SetType=""New"""			'#��ۺ� ���� | New : ��ۺ� �׷��ȣ �ű� ����, Use : ���� ��ۺ� �׷��ȣ ���
			strRst = strRst & "					BundleNo=""0"""				'������ȣ | ��ۺ� �׷��ڵ带 ����ó �������� ���� ��ۺ�� ���� �� ��� ��� AddAddressBook�� BundleNO�� RefundAddrNum�� ���� ����
			strRst = strRst & "					GroupCode="""""				'��ۺ� �׷��ڵ� | SetType: Use �� ���
			strRst = strRst & "					RefundAddrNum=""740092"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"	'��ǰ����� ��ȣ | AddAddressBook�� AddressCode
			strRst = strRst & "					<NewItemShipping "
			strRst = strRst & "						FeeCondition=""ConditionalFee"" "	'��ǰ�� ��ۺ� ���� | SetType�� New �� ��� Free : ����, ConditionalFee : ���Ǻι���, FixedFee : ����, PrepayableOnDelivery : ���Ҽ�����, PayOnDelivery : ����
			strRst = strRst & "						FeeBasePrice=""30000"""				'��ǰ�� ��ۺ� ���� | ���Ǻι����� ���
			strRst = strRst & "						Fee=""2500"""						'���Ǻι����̰ų� ������ ���
			strRst = strRst & "					/>"
			strRst = strRst & "				</Shipping>"
		Else
			'GroupCode = "389827401"
			GroupCode = "856237774"		'5���� �̸� 3õ�� ��ۺ��ڵ� 2020-01-10 ������ ����

			strRst = strRst & "				<Shipping "
			strRst = strRst & "					SetType=""Use"""				'#��ۺ� ���� | New : ��ۺ� �׷��ȣ �ű� ����, Use : ���� ��ۺ� �׷��ȣ ���
			strRst = strRst & "					BundleNo=""0"""					'������ȣ | ��ۺ� �׷��ڵ带 ����ó �������� ���� ��ۺ�� ���� �� ��� ��� AddAddressBook�� BundleNO�� RefundAddrNum�� ���� ����
			strRst = strRst & "					GroupCode="""&GroupCode&""""	'��ۺ� �׷��ڵ� | SetType: Use �� ���
			strRst = strRst & "					RefundAddrNum=""740092"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"	'��ǰ����� ��ȣ | AddAddressBook�� AddressCode
'			strRst = strRst & "					<NewItemShipping "
'			strRst = strRst & "						FeeCondition=""Free or ConditionalFee or PayOnDelivery or PrepayableOnDelivery or FixedFee"" "
'			strRst = strRst & "						FeeBasePrice=""decimal"""
'			strRst = strRst & "						Fee=""decimal"""
'			strRst = strRst & "					/>"
			strRst = strRst & "				</Shipping>"
		End If
		getGmarketShippingParam = strRst
	End Function

	'�⺻���� Gmarket ��� soap XML
	Public Function getGmarketItemRegParameter(isReg)
		Dim strRst, tt, isMadeInKorea
		If Fsourcearea = "�ѱ�" OR Fsourcearea = "���ѹα�" Then
			isMadeInKorea = "Domestic"		'����
		Else
			isMadeInKorea = "Foreign"		'����
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
  		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
		strRst = strRst & "		</EncTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<AddItem xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<AddItem "
		strRst = strRst & "				OutItemNo="""&FItemid&""""							'#�ܺλ�ǰ��ȣ | ���޻� ��ǰ ��ȣ
		strRst = strRst & "				CategoryCode="""&FDepthCode&""""					'#�Һз��ڵ�
		If isReg Then
			strRst = strRst & "			GmktItemNo="""&FGmarketGoodNo&""""					'G���� ��ǰ��ȣ | ��ǰ���� ������
		End If
		strRst = strRst & "				ItemName="""&getItemNameFormat&""""					'#��ǰ��
'		strRst = strRst & "				ItemEngName=""string"""								'������ǰ��
		strRst = strRst & "				ItemDescription="""""								'#��ǰ������
		strRst = strRst & "				GdHtml="""&replaceRst(getGmarketItemContParamToReg)&""""		'New ��ǰ ������ - ��ǰ����
'		strRst = strRst & "				GdHtml=""string"""									'New ��ǰ ������ - ��ǰ����
'		strRst = strRst & "				GdAddHtml=""string"""								'New ��ǰ ������ - �߰�����
'		strRst = strRst & "				GdPrmtHtml=""string"""								'New ��ǰ ������ - ����/ȫ��
		strRst = strRst & "				MakerNo="""&MAKERNO&""""							'#�������ȣ | �켱 ��Ÿ�� ���� ��������
'		strRst = strRst & "				BrandNo="""&FBrandCode&""""							'�귣���ȣ
		strRst = strRst & "				BrandNo=""100356"""									'�귣���ȣ  2019-02-21 16:16 ������ ����(�ٹ����� �ڵ�(100356)�� �Ƚ�)
'		strRst = strRst & "				ModelName=""string"""								'�𵨸�
		strRst = strRst & "				IsAdult="""&Chkiif(IsAdultItem() = "Y", "true", "false")&""""	'#���ο�ǰ ���� | true, false
		strRst = strRst & "				Tax="""&CHKIIF(FVatInclude="N","Free","VAT")&""""	'#�ΰ��� �鼼���� | VAT, Free
'		strRst = strRst & "				MadeDate=""date"""									'����(����)�����
'		strRst = strRst & "				AppearedDate=""date"""								'��ó��
		strRst = strRst & "				ExpirationDate=""2078-12-31"""						'#��ȿ�� | ex. 2011-01-02 | 1/1/1900 12:00:00 AM and 6/6/2079 11:59:59 PM.
'		strRst = strRst & "				FreeGift=""string"""								'����ǰ
		strRst = strRst & "				ItemKind=""Shipping"""								'#��ǰ���� | Shipping: ��ۻ�ǰ / Ecoupon: ��������ǰ
'		strRst = strRst & "				InventoryNo=""string"""								'�Ǹ��ڰ����ڵ� | ��ǰ ������ ����� ����(��ǰ���� , ��������)�� ��ϵ� code�� �ֹ������� �����Ͽ� ����
'		strRst = strRst & "				ItemWeight=""double"""								'��ǰ ����
		strRst = strRst & "				IsOverseaTransGoods=""false"""						'�ؿܹ�� ���� ���� | True : ��ü ���� ��� ����, False : ��ü ���� ��� �Ұ�
'		strRst = strRst & "				IsGift=""false"""									'�����ϱ� ��ǰ ���� | �����ϱ� ��ǰ ���� ���� true: �����ϱ� ���� false: �����ϱ� �Ұ��� * ���Է�/skip �� default = true �� ���
'		strRst = strRst & "				FreeDelFeeType=""int"""								'�����ۺ� Ÿ�� | 1 : ������ ���� ����, 2 : ��ġ ��ۺ�, 3 : ���� ���� ���� ��ǰ
'		strRst = strRst & "				IsGmktDiscount=""boolean"""							'G���� ���� ���� ���� | True : ���Է½� ���� ����, False�� ��� G���� �δ� ���� ���� �Ұ�
		strRst = strRst & "				>"
'		strRst = strRst & "				<ReferencePrice "
'		strRst = strRst & "					Kind=""Quotation or Department or HomeShopping"""					'������ ���� | Quotation : ���߰�, Department : ��ȭ����, HomeShopping : Ȩ���ΰ�
'		strRst = strRst & "					Price=""decimal"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'������
		strRst = strRst & "				<Refusal "
		strRst = strRst & "					IsPriceCompare=""false"""											'���ݺ� �������� | true, false ..2017-02-17 true->false�� ����
		strRst = strRst & "					IsNego=""true"""													'�����ϱ� �������� | true, false-
		strRst = strRst & "					IsJaehuDiscount=""true"""											'�������� ���� | true, false
		strRst = strRst & "					IsPack=""false"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'��ٱ��� �Ұ� | True �� ��� ��ٱ��� �� ���� ó���̸�, ���� false�� ���� �� ���� �̽��� ������, ����true�� ���� �� ��ǰ ������ false �Ǵ� null�� ���� ��� ��ٱ��� ���� ������ Ǯ��
		strRst = strRst & getGmarketAddImageParam()
		strRst = strRst & "				<As Telephone=""1644-6035"" Address=""Seller"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'#����ó, AS���� �ּ�/���� | Manufacturing_Seller : ������AS ���ͳ� �Ǹ��ڿ��� ����, Seller : �Ǹ��ڿ��� ����
		strRst = strRst & getGmarketShippingParam()
'		strRst = strRst & "				<BundleOrder BuyUnitCount=""int"" MinBuyCount=""int"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'BuyUnitCount : �ּұ��ż���, MinBuyCount : ���ż�������
		strRst = strRst & "				<OrderLimit OrderLimitCount="""&getOrderMaxNum&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'OrderLimitCount : �ִ뱸�Ű��ɼ���, OrderLimitPeriod : ���ż���Ȯ��
	If FDepth4Code <> "" Then
		strRst = strRst & "				<AttributeCode AttributeCode="""&FDepth4Code&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'�з��Ӽ��ڵ�
	End If
		strRst = strRst & "				<Origin Code="""&isMadeInKorea&""" Place="""&Fsourcearea&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'������ ���� | Domestic : ����, Foreign : ����, Etc : ��...��������
'		strRst = strRst & "				<Book ISBN=""string"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'���� ISBN �ڵ� | ISBN ��Ͻ� �ֹ� �ɼ� ��� �Ұ�
		strRst = strRst & "				<GoodsKind GoodsKind=""New"" GoodsStatus=""NotUsed"" GoodsTag=""Default"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'��ǰ���� | New : �Ż�ǰ, Used : �߰��ǰ...NotUsed : �̻��, AlmostNew : ���� ����, Fine : ��ȣ, Old : �ణ ����, ForCollect : ��� �Ұ�(������)
'		strRst = strRst & "				<GoodsKind GoodsKind=""Unknown or New or Stock or Used or Returned or Displayed or Refurbished"" GoodsStatus=""None or Under3Months or Under6Months or Under1Year or Over2Years or NotUsed or AlmostNew or Fine or Old or ForCollect or Sealed or Unsealed or UsedAfterUnsealed or DisplayedNotUsed or DisplayedAlmostNew or DisplayedFine or DisplayedOld or DisplayedForCollect"" GoodsTag=""Default or New or Hot or Sale or MDRecommend or InterestFree or Limited or Gift or LowestPrice or NoMargin or Donation or SpecialBargain or EyeCatch or PowerDealer or Premium2Days or Premium7Days or Premium2Weeks or Premium1Month or Premium2Month or Premium3Month or ImmediateDelivery or Patronage or PremiumPlus"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		strRst = strRst & "			</AddItem>"
		strRst = strRst & "		</AddItem>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
'response.write strRst
'response.end
		getGmarketItemRegParameter = strRst
	End Function

	'�ɼǵ�� Soap XML
	Public Function getGmarketItemOptRegParameter()
		Dim strSQL, strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
		strRst = strRst & "		</EncTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<AddItemOption xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<AddItemOption GmktItemNo="""&FGmarketGoodNo&""">"
		strRst = strRst & getGmarketOptParamtoReg()
	If FItemdiv = "06" Then
		strRst = strRst & "				<ItemTextList xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
		strRst = strRst & "					<ItemText Name=""�ؽ�Ʈ�� �Է��ϼ���"" />"
		strRst = strRst & "				</ItemTextList>"
	End If
		strRst = strRst & "			</AddItemOption>"
		strRst = strRst & "		</AddItemOption>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getGmarketItemOptRegParameter = strRst
	End Function

	Public Function getGmarketOptParamtoReg()
		Dim strRst, strSql, IsCombination, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, type3, optDc1, optDc2, optDc3
		Dim optNm, optDc, optLimit, itemoption, IsDisplayable, Remain
		MultiTypeCnt = 0
		IsCombination = "false"

		If FOptionCnt = 0 Then			'��ǰ
			strRst = "<ItemSelectionList IsInventory=""true"" IsCombination="""&IsCombination&""" OptionImageLevel=""0""  xmlns=""http://tpl.gmarket.co.kr/tpl.xsd""></ItemSelectionList>"
		Else							'�ɼ��ִ� ��ǰ
			strSql = ""
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
	        rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) Then
				IsCombination = "true"
				MultiTypeCnt = rsget.recordcount
				Do until rsget.EOF
					arrMultiTypeNm = arrMultiTypeNm & db2Html(rsget("optionTypeName"))&"^|^"
					rsget.MoveNext
				Loop
			End If
			rsget.Close

			'1. strRst ������ ����
			strRst = ""
			strRst = strRst & "				<ItemSelectionList "
			strRst = strRst & "					IsInventory=""true"""					'#����뿩�� | �ɼǺ� ��� ���� �ʼ�
			strRst = strRst & "					OptionSortType=""Register"""			'#���ļ��� | Register(��ϼ�), Price(���ݼ�), Name(�̸���)
			strRst = strRst & "					IsCombination="""&IsCombination&""""	'#������ ��� | �� �Է½� False�� ó��, True�� ��� ������ �ɼ� ���
			strRst = strRst & "					OptionImageLevel=""0"""					'New�ɼ� ����� True�� ��� �ʼ�, 0 : �̹��� ��Ī �� ���, 1 : �ɼǸ� 1�� ��Ī, 2 : �ɼǸ� 2�� ��Ī
			strRst = strRst & "					xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
			'1.strRst ������ ���� ��

			strSql = ""
			strSql = strSql & "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					optLimit = rsget("optLimit")
					optIsusing	= rsget("isusing")
					optSellYn	= rsget("optsellyn")
					optLimit = optLimit-5
					If (optLimit < 1) Then optLimit = 0
					If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					If (optIsusing <> "Y") OR (optSellYn <> "Y") Then optLimit = 0
					itemoption	= rsget("itemoption")
					optDc		= replaceRst(rsget("optionname"))
					optIsusing	= rsget("isusing")
					optSellYn	= rsget("optsellyn")
					optaddprice	= rsget("optaddprice")
					strRst = strRst & "					<ItemSelection "
					If IsCombination = "true" Then
						If Right(arrMultiTypeNm,3) = "^|^" Then
							arrMultiTypeNm = Left(arrMultiTypeNm, Len(arrMultiTypeNm) - 3)
						End If
						strRst = strRst & "					Name="""&arrMultiTypeNm&""""				'#������ | �׸���� �ִ� 5��,�� ���ü��� �ִ� 500��, �ɼǸ� �� �ִ� 25�� New�ɼ� ��� ���ΰ� True�� ��� ������ ��� ó��  ex) ����^|^������^|^����ǰ
					Else
						If db2Html(rsget("optionTypeName")) <> "" Then
							optNm = db2Html(rsget("optionTypeName"))
						Else
							optNm = "�ɼ�"
						End If
						strRst = strRst & "					Name="""&optNm&""""							'#������ | �׸���� �ִ� 5��,�� ���ü��� �ִ� 500��, �ɼǸ� �� �ִ� 25�� New�ɼ� ��� ���ΰ� True�� ��� ������ ��� ó��  ex) ����^|^������^|^����ǰ
					End If
					strRst = strRst & "						Code="""&itemoption&"""" 					'�ɼ� �Ǹ��� �ڵ�
					strRst = strRst & "						Value="""&Replace(replace(optDc, ",", "^|^"), "���þ���", "���þ���.")&"""" 	'#������ | ���� ���� �� �ִ� 10��, New�ɼ� ��� ���ΰ� True�� ��� ������ ��� ó�� ��ϵǴ� �ɼ� ���� ������ �ɼǸ��� ���а� ���� �Ͼ� ��, ex) ����^|^90^|^���
					'strRst = strRst & "					Value="""&replace(optDc, ",", "^|^")&"""" 	'#������ | ���� ���� �� �ִ� 10��, New�ɼ� ��� ���ΰ� True�� ��� ������ ��� ó�� ��ϵǴ� �ɼ� ���� ������ �ɼǸ��� ���а� ���� �Ͼ� ��, ex) ����^|^90^|^���
					strRst = strRst & "						Price="""&optaddprice&"""" 					'#���� | �׸񺰷� ������ 0���� ���� 1�� �̻� ����, �ǸŰ����� -50% ~ +100% �̳�. ������ 10�� ����, ��,�� �Է� �Ұ�
					strRst = strRst & "						Remain="""&optLimit&""""			 		'#������
					If IsCombination = "true" Then
						strRst = strRst & "					OptionImageUrl=""-"""						'�ɼ� �̹��� URL | ������ �ɼ� ��Ͻ� 0 : �̹��� ��Ī �̻���� ��쿡�� ��(����1��)�� �־� �ش� fieldȣ��
					End If
					strRst = strRst & "					/>"
					rsget.MoveNext
				Loop
			End If
			rsget.Close
			strRst = strRst & "				</ItemSelectionList>"
		End If
		getGmarketOptParamtoReg = strRst
	End Function

	'�̹��� ���� Soap XML
	Public Function getGmarketItemEditImgParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
		strRst = strRst & "		</EncTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<EditItemImage xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<EditItemImage GmktItemNo="""&FGmarketGoodNo&""">"
		strRst = strRst & getGmarketAddImageParam()
		strRst = strRst & "			</EditItemImage>"
		strRst = strRst & "		</EditItemImage>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getGmarketItemEditImgParameter = strRst
	End Function

	Public Function getGmarketAddPriceParameter(isReged, mustPrice, idisplayDate)		''���� XML������ �ʿ��ϴٸ�..incGmarketFunction�� getGmarketAddPriceParameter�� ���� ����
		Dim strSQL, strRst, GetTenTenMargin, iStockQty

		'���� �Ⱓ ����
		If FDisplayDate = "" or isnull(FDisplayDate) Then
			idisplayDate = DateAdd("yyyy", 1, Date())
		Else
			If DateDiff("m", FDisplayDate, Date()) <= 3 Then
				idisplayDate = DateAdd("yyyy", 1, Date())
			Else
				'idisplayDate = FDisplayDate
				idisplayDate = DateAdd("d", 1, Date())
			End If
		End If

		'��� ���� ����
		If isReged = "N" Then
			iStockQty = 0
		Else
			If FLimityn = "Y" Then
				iStockQty = Flimitno - Flimitsold - 5
				If iStockQty > 1000 Then
					iStockQty = CDEFALUT_STOCK
				End If
			Else
				iStockQty = CDEFALUT_STOCK
			End If
			If (iStockQty < 1) Then iStockQty = 0
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
		strRst = strRst & "		</EncTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<AddPrice xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<AddPrice "
		strRst = strRst & "				GmktItemNo="""&FGmarketGoodNo&""""			'#G���� ��ǰ��ȣ
		strRst = strRst & "				DisplayDate="""&idisplayDate&""""			'#�ֹ��Ⱓ | �ִ� 1��
		strRst = strRst & "				SellPrice="""&mustPrice&""""				'#�ǸŰ��� | �ּ� 100�� �̻� 1,000,000,000�� �̸� (100�� ����)
		strRst = strRst & "				StockQty="""&iStockQty&""""					'#������
		strRst = strRst & "				InventoryNo="""&FItemid&""" />"				'�Ǹ��ڰ����ڵ�
		strRst = strRst & "		</AddPrice>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getGmarketAddPriceParameter = strRst
	End Function

	'G9 ��� soap XML
	Public Function getG9ItemRegParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
		strRst = strRst & "	<soap:Header>"
		strRst = strRst & "		<EncTicket xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<encTicket>"&gmarketTicket&"</encTicket>"
		strRst = strRst & "		</EncTicket>"
		strRst = strRst & "	</soap:Header>"
		strRst = strRst & "	<soap:Body>"
		strRst = strRst & "		<AddG9Item xmlns=""http://tpl.gmarket.co.kr/"">"
		strRst = strRst & "			<AddG9Item "
		strRst = strRst & "				GmktItemNo="""&FGmarketGoodNo&""""			'#G���� ��ǰ��ȣ
		strRst = strRst & "				SellManageYn=""N"""							'#�ǰ�/��� ���� ���� | Y : G9 �Ǹſ� ����/��� ��� ���ϵǴ� ���� ��ǰ�ڵ�� ���� ���� ��� API�� ���� ���� (addprice), N : ���� ��� ��ǰ�� ����/��� ��� - ���� ��� ��ǰ ����/��� ����� ���� ��ǰ�� ����/��� �����ϰ� ����ȭ ó�� - �� �Է½� False�� ó��
		strRst = strRst & "				CostManageYn=""N"""							'#G9 �Ǹ��� ���� ���� | Y : G9 �Ǹſ� ���� ���� ��å ��� ���ϵǴ� ���� ��ǰ�ڵ�� ���� ���� API�� ���� ���� (AddPremiumItem), N : ���� ��� ��ǰ�� ���� ��å ���� ���� ��� ���� ��å ����� ���� ��ǰ�� ���� ��å�� �����ϰ� ����ȭ ó�� - �� �Է½� False�� ó��
		strRst = strRst & "				ItemManageYn=""N"" />"						'#�⺻��ǰ���� ���� ���� | Y : G9 �Ǹſ� ���� ��ǰ���� ��� ���ϵǴ� ���� ��ǰ�ڵ�� ���� ��ǰ���� ��ϼ��� API�� ���� ���� (AddItem), N : ���� ��� ��ǰ�� ��ǰ �⺻������ ��� - ���� ��� ��ǰ�� ��ǰ��, ī�װ�����, ��ۺ� ���� ���� ������ �����ϰ� ����ȭ ó��
		strRst = strRst & "		</AddG9Item>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getG9ItemRegParameter = strRst
	End Function
End Class

Class CGmarket
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
	Public Sub getGmarketNotRegOneItem
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.gmarketStatCD,-9) as gmarketStatCD "
		strSql = strSql & "	, UC.socname_kor, isnull(am.depthCode, '') as depthCode, am.depth4Code, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		' strSql = strSql & "	, isnull(bm.BrandCode, '') as BrandCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_gmarket_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_category as tm on am.depthCode = tm.depthCode and am.depth4Code = tm.depth4Code "
		' strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_brand_mapping as bm on bm.makerid = i.makerid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " Where i.isusing='Y' "
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
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'�ö��/ȭ�����/�ؿ�����
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''���� ǰ�� �� ��� ����.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.itemid in ('2594140', '2594138', '2594139', '2557733' , '2558483', '2549730', '2549728') THEN 'Y' "	'2019-12-02 ������..��ϸ��������̳� ��� ��û'
		strSql = strSql & " 				WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'������� �귣��
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'������� ��ǰ
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"' and i.mwdiv <> 'M')) "	'������� ī�װ�
		strSql = strSql & "	and isnull(R.APIadditem, '') <> 'Y' "									'�⺻���� ��ϵ������� ����ϸ� �� ��
		strSql = strSql & "	and isnull(R.GmarketGoodNo, '') = '' "
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
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
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FGmarketStatCD		= rsget("gmarketStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FDepth4Code		= rsget("depth4Code")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
				' FOneItem.FBrandCode 		= rsget("BrandCode")
		End If
		rsget.Close
	End Sub

	'// �̵�� �ɼ�(��Ͽ�)
	Public Sub getGmarketNotOptOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, J.GmarketGoodNo, isnull(J.APIadditem, 'N') as APIadditem, isnull(J.APIaddopt, 'N') as APIaddopt "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_gmarket_regItem as J on i.itemid = J.itemid"
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and J.itemid = '"&FRectItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FitemDiv			= rsget("itemdiv")
				FOneItem.FoptionCnt			= rsget("optionCnt")
				FOneItem.FLimitYn			= rsget("LimitYn")
				FOneItem.FLimitNo			= rsget("LimitNo")
				FOneItem.FLimitSold			= rsget("LimitSold")
				FOneItem.ForgPrice			= rsget("orgPrice")
				FOneItem.ForgSuplyCash		= rsget("orgSuplyCash")
				FOneItem.FSellCash			= rsget("sellcash")
				FOneItem.FBuyCash			= rsget("buycash")
				FOneItem.FGmarketGoodNo		= rsget("GmarketGoodNo")
				FOneItem.FAPIadditem		= rsget("APIadditem")
				FOneItem.FAPIaddopt			= rsget("APIaddopt")
		End If
		rsget.Close
	End Sub

	'// ������ �̹���
	Public Sub getGmarketEditImageOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.*, J.GmarketGoodNo "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_gmarket_regItem as J on i.itemid = J.itemid"
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and J.itemid = '"&FRectItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FGmarketGoodNo		= rsget("GmarketGoodNo")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.FListimage			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
		End If
		rsget.Close
	End Sub

	Public Sub getGmarketEditOneItem
		Dim strSql, addSql, i, infoContent1919807
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
		strSql = strSql & "	, m.GmarketGoodNo, m.Gmarketprice, m.GmarketSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, isnull(am.depthCode, '') as depthCode, am.depth4Code, isNull(m.returnShippingFee, 0) as returnShippingFee, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		' strSql = strSql & "	, isnull(bm.BrandCode, '') as BrandCode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_gmarket_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_etcmall.dbo.tbl_gmarket_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_category as tm on am.depthCode = tm.depthCode and am.depth4Code = tm.depth4Code "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		' strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_brand_mapping as bm on bm.makerid = i.makerid "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIadditem = 'Y' "
		strSql = strSql & addSql
		strSql = strSql & " and m.GmarketGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
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
				If FRectItemID = "1919807" Then
					infoContent1919807 = ""
					infoContent1919807 = infoContent1919807 & "<div align=""center"">"
					infoContent1919807 = infoContent1919807 & "	<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gmarket.jpg""></p>"
					infoContent1919807 = infoContent1919807 & "	<p style=""text-align: center;""><br> <img src=""http://gi.esmplus.com/blanktv/10x10/gong100/cleaner.jpg""></p> "
					infoContent1919807 = infoContent1919807 & "	<img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_gmarket.jpg"">"
					infoContent1919807 = infoContent1919807 & "</div>"
					FOneItem.Fitemcontent = infoContent1919807
				End If

				FOneItem.FGmarketGoodNo		= rsget("GmarketGoodNo")
				FOneItem.FGmarketprice		= rsget("Gmarketprice")
				FOneItem.FGmarketSellYn		= rsget("GmarketSellYn")

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
	            FOneItem.Fregitemname    	= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")

				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FDepth4Code		= rsget("depth4Code")
				FOneItem.FReturnShippingFee	= rsget("returnShippingFee")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
				' FOneItem.FBrandCode 		= rsget("BrandCode")
		End If
		rsget.Close
	End Sub

	Public Sub getGmarketEditPriceOptOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'���û�ǰ�� �ִٸ�
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.*, m.GmarketGoodNo, m.Gmarketprice, m.GmarketSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt, m.displayDate, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100) < "&CMAXMARGIN&" )) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"' and i.mwdiv <> 'M')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_gmarket_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIadditem = 'Y' "
		strSql = strSql & " and m.APIaddopt = 'Y' "
		strSql = strSql & " and m.GmarketStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.GmarketGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FMakerid			= rsget("makerid")
				FOneItem.FGmarketGoodNo		= rsget("GmarketGoodNo")
				FOneItem.FGmarketprice		= rsget("Gmarketprice")
				FOneItem.FGmarketSellYn		= rsget("GmarketSellYn")
	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
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
	            FOneItem.FmaySoldOut		= rsget("maySoldOut")
	            FOneItem.FDisplayDate		= rsget("displayDate")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	'//G9 �̵�� ��ǰ ���(��Ͽ�)
	Public Sub getG9NotRegOneItem
		Dim strSql, addSql, i
		strSql = " EXEC [db_etcmall].[dbo].[usp_API_G9_Reg_Get] " & FRectItemID
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGmarketItem
				FOneItem.Fitemid			= rsget("itemid")
				FOneItem.FGmarketGoodno		= rsget("GmarketGoodno")
				FOneItem.FG9GoodNo			= rsget("G9GoodNo")
		End If
		rsget.Close
	End Sub
End Class

'������ ��ǰ�ڵ� ���
Function getGmarketGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 Gmarketgoodno FROM db_etcmall.dbo.tbl_gmarket_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getGmarketGoodno = rsget("Gmarketgoodno")
	rsget.Close
End Function

'������ ��� ���� ī�װ� ���� Ȯ��
Function getGmarketChildrenCate(iitemid, byref isChildrenCate, byref isLifeCate, byref isElecCate)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.itemid, isChildrenCate, isLifeCate, isElecCate "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " INNER JOIN (  "
	strSql = strSql & " 	SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
	strSql = strSql & " 	FROM db_etcmall.dbo.tbl_gmarket_cate_mapping "
	strSql = strSql & " 	GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
	strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
	strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
	strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_gmarket_category as tm on am.depthCode = tm.depthCode and am.depth4Code = tm.depth4Code "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	strSql = strSql & " and (isChildrenCate = 'Y' OR isLifeCate = 'Y' OR isElecCate = 'Y') "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		isChildrenCate	= rsget("isChildrenCate")
		isLifeCate		= rsget("isLifeCate")
		isElecCate		= rsget("isElecCate")
'		getGmarketChildrenCate = rsget("isChildrenCate")
	End If
	rsget.Close
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
    v = replace(v, "&", "&amp;")
    v = replace(v, """", "&quot;")
	'v = Replace(v,"<br>","&#xA;")
	'v = Replace(v,"</br>","&#xA;")
	'v = Replace(v,"<br />","&#xA;")
	v = Replace(v,"<","&lt;")
	v = Replace(v,">","&gt;")
    replaceRst = v
end function

Function getAllRegChk(iitemid, iaddPrice)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT Count(*) as cnt " & VBCRLF
	sqlStr = sqlStr & " FROM db_etcmall.dbo.tbl_gmarket_regItem " & VBCRLF
	sqlStr = sqlStr & " WHERE itemid='"&iitemid&"'"
	sqlStr = sqlStr & " and APIadditem = 'Y' "
	sqlStr = sqlStr & " and APIaddgosi = 'Y' "
	If iaddPrice = "X" Then
		sqlStr = sqlStr & " and APIaddopt = 'Y' "
	End If
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If rsget("cnt") = 0 Then
		getAllRegChk = "N"
	Else
		getAllRegChk = "Y"
	End If
	rsget.Close
End Function

Function getAllRegChk2(iitemid, byref iGmarketGoodNo, byref ioptioncnt, byref iLimityn, ichk)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT TOP 1 R.itemid, i.optioncnt, R.GmarketGoodNo, i.limityn "
	sqlStr = sqlStr & " from db_etcmall.dbo.tbl_gmarket_regItem as R "
	sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item as i on R.itemid = i.itemid "
	If ichk = "Y" Then
		sqlStr = sqlStr & " and APIadditem = 'Y' "
		sqlStr = sqlStr & " and APIaddgosi = 'Y' "
		sqlStr = sqlStr & " and APIaddopt = 'Y' "
	End If
	sqlStr = sqlStr & " WHERE i.itemid = '"&iitemid&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ioptioncnt		= rsget("optioncnt")
		iGmarketGoodNo	= rsget("GmarketGoodNo")
		iLimityn		= rsget("limityn")
	End If
	rsget.Close
End Function

Function fnStrLength(str)
	Dim strLen, strByte, strCut, strRes, char, i
	strLen = 0
	strByte = 0
	strLen = Len(str)
	for i = 1 to strLen
		char = ""
		strCut = Mid(str, i, 1)
		char = len(hex(ascw(strCut)))

		'if Len(char) = 1 And char = "1" then
		if char = 2 then
			strByte = strByte + 1
		else
			strByte = strByte + 2
		end if
	next
	fnStrLength = strByte
End function
%>
