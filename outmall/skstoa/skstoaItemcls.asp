<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "skstoa"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CSkstoaItem
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
	Public FSocname_kor
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FSkstoaStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FdepthCode
	Public FbasicimageNm
	Public FReglevel
	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
    Public FsafetyNum
    Public FmaySoldOut

    Public Fregitemname
    Public FregImageName
	Public FOrderMaxNum
	Public FAdultType
	Public FLgroup
	Public FMgroup
	Public FSgroup
	Public FDgroup
	Public FTgroup
	Public FOutmallstandardMargin
	Public FSkstoaTmpGoodNo
	Public FSkstoaGoodNo
	Public FSkstoaprice
	Public FSkstoaSellYn

	Public Function getOrderMaxNum()
		getOrderMaxNum = Cint(FOrderMaxNum)
		If Cint(FOrderMaxNum) > 999 Then
			getOrderMaxNum = 999
		End If
	End Function

	Function RightCommaDel(ostr)
		Dim restr
		restr = ""
		If IsNULL(ostr) Then Exit Function
		restr = Trim(ostr)
		If (Right(restr,1)=",") Then restr = Left(restr,Len(restr)-1)
		RightCommaDel = restr
	End Function

	public Function getKeywords()
		Dim strRst
		strRst = "�ٹ�����," & FKeywords
		strRst = replace(strRst, "�α�", "")
		strRst = replace(strRst, "��ġ", "")
		strRst = replace(strRst, "�����ġ", "")
		strRst = RightCommaDel(strRst)
		If strRst = "" Then
			strRst = "�ٹ�����"
		End If

		If stringCount(strRst) > 100 Then
			strRst = chrbyte(strRst, 100, "N")
		End If
		getKeywords = URLEncodeUTF8Plus(strRst)
	End Function

	Function stringCount(strString)
		Dim intPos, chrTemp, intLength
		'���ڿ� ���� �ʱ�ȭ
		intLength = 0
		intPos = 1

		'���ڿ� ���̸�ŭ ����
		while ( intPos <= Len( strString ) )
			'���ڿ��� �ѹ��ھ� ���Ѵ�
			chrTemp = ASC(Mid( strString, intPos, 1))
			if chrTemp < 0 then '������(-)�� ������ �ѱ���
				intLength = intLength + 2 '�ѱ��� ��� 2����Ʈ�� ���Ѵ�
			else
				intLength = intLength + 1 '�ѱ��� �ƴҰ�� 1����Ʈ�� ���Ѵ�
			end If
			intPos = intPos + 1
		wend
		stringCount = intLength
	End function

	'// ǰ������
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
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

	Public Function MustBuyPrice(imustPrice)
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin, ownItemCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 (100 - m.margin) * 0.01 as fmargin, d.itemid "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_ssg_marginItem_master] as m  "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_ssg_marginItem_detail] as d on m.idx = d.midx  "
		sqlStr = sqlStr & " WHERE m.isusing = 'Y'  "
		sqlStr = sqlStr & " and convert(char(10), getdate(), 120) between m.startDate and m.enddate  "
		sqlStr = sqlStr & " and m.mallid = 'skstoa' "
		sqlStr = sqlStr & " and d.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			MustBuyPrice = Clng(imustPrice * rsget("fmargin"))
		Else
			MustBuyPrice = Clng(imustPrice * 0.88)
		End If
		rsget.Close
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, outmallstandardMargin, ownItemCnt
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

			If FSkstoaPrice = 0 Then
				If (GetTenTenMargin < outmallstandardMargin) Then
					tmpPrice = Forgprice
				Else
					tmpPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < outmallstandardMargin Then
					If (Forgprice < Round(FSkstoaPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FSkstoaPrice * 0.35, 0)/10)*10)
					ElseIf Clng(Forgprice) > Clng(Round(FSkstoaPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FSkstoaPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = Forgprice
					End If
				Else
					If (FSellCash < Round(FSkstoaPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FSkstoaPrice * 0.35, 0)/10)*10)
					ElseIf Clng(FSellCash) > Clng(Round(FSkstoaPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FSkstoaPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'// Skstoa �Ǹſ��� ��ȯ
	Public Function getskstoaSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getskstoaSellYn = "Y"
			Else
				getskstoaSellYn = "N"
			End If
		Else
			getskstoaSellYn = "N"
		End If
	End Function

	'// Skstoa ��������Ÿ��
	Public Function getSkstoaOfferType()
		Dim buf
		Select Case FinfoDiv
			Case "35"	buf = "38"
			Case "36"	buf = "35"
			Case "47"	buf = "39"
			Case "48"	buf = "40"
			Case Else	buf = FinfoDiv
		End Select
		getSkstoaOfferType = buf
	End Function

	Public Function getSkstoaContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_skstoa.jpg""></p><br>")
		ForderComment = replace(ForderComment,"&nbsp;"," ")
		ForderComment = replace(ForderComment,"&nbsp"," ")
		ForderComment = replace(ForderComment,"&"," ")
		ForderComment = replace(ForderComment,chr(13)," ")
		ForderComment = replace(ForderComment,chr(10)," ")
		ForderComment = replace(ForderComment,chr(9)," ")
		If ForderComment <> "" Then
			strRst = strRst & "- �ֹ��� ���ǻ��� :<br>" & URLEncodeUTF8(Fordercomment) & "<br>"
		End If

		'#�⺻ ��ǰ����
		Fitemcontent = replace(Fitemcontent,"&nbsp;"," ")
		Fitemcontent = replace(Fitemcontent,"&nbsp"," ")
		Fitemcontent = replace(Fitemcontent,"&"," ")
		Fitemcontent = replace(Fitemcontent,chr(13)," ")
		Fitemcontent = replace(Fitemcontent,chr(10)," ")
		Fitemcontent = replace(Fitemcontent,chr(9)," ")
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & URLEncodeUTF8(Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & URLEncodeUTF8(Fitemcontent & "<br>")
			Case Else
				strRst = strRst & URLEncodeUTF8(Fitemcontent & "<br>")
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
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_skstoa.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getSkstoaContParamToReg = strRst
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
'		If InStr(ibuf,"-") < 1 Then
'			isImageChanged = FALSE
'			Exit Function
'		End If
'		isImageChanged = ibuf <> FRegImageName
		If ibuf = FRegImageName Then
			isImageChanged = False
		Else
			isImageChanged = True
		End If
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	Public Function checkTenItemOptionValid2()
		Dim strSql, chkRst, optValid
		chkRst = true

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Skstoa_OptionValid_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			optValid = rsget("optValid")
		End If
		rsget.Close

		If optValid = "N" Then
			chkRst = false
		End If
		'//��� ��ȯ
		checkTenItemOptionValid2 = chkRst
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
'				strSql = strSql & " 	and optaddprice=0 "
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
 
    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST��ǰ] "&FItemName
		Else
			'buf = "[�ٹ�����] "&FItemName
			buf = FItemName		'2022-02-07 �������� ��û / ��ǰ��� �ٹ����� ����
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
		buf = replace(buf,"_","/")
        buf = replace(buf,"%","����")
		buf = replace(buf,"&","/")
        buf = replace(buf,"&amp;","")
        buf = replace(buf,"[������]","")
        buf = replace(buf,"[���� ���]","")
'        buf = LeftB(buf, 40)
        getItemNameFormat = URLEncodeUTF8Plus(buf)
    end function

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

	Public Function IsMakeItem()
		Select Case FItemdiv
			Case "06", "16"
				IsMakeItem = "Y"
			Case Else
				IsMakeItem = "N"
		End Select
	End Function

	Function getMakecoCode()
		Dim strSql
		strSql = strSql & " SELECT TOP 1 makeCompanyCode "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_skstoa_makeCompanyCode] "
		strSql = strSql & " WHERE makeCompanyName like '%"& html2db(Fmakername) &"%' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getMakecoCode = rsget("makeCompanyCode")
		Else
			getMakecoCode = "9999"	'��Ÿ
		End If
		rsget.Close
	End Function

	Function getOriginCode()
		Dim strSql
		strSql = strSql & " SELECT TOP 1 originCode "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_skstoa_originCode] "
		strSql = strSql & " WHERE originName like '%"& html2db(Fsourcearea) &"%' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getOriginCode = rsget("originCode")
		Else
			getOriginCode = "9999"	'��Ÿ
		End If
		rsget.Close
	End Function

	Function getBrandCode()
		Dim strSql
		strSql = strSql & " SELECT TOP 1 brandCode "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_skstoa_brandCode] "
		strSql = strSql & " WHERE brandName = '"& html2db(FSocname_kor) &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getBrandCode = rsget("brandCode")
		Else
			getBrandCode = "999999"	'��Ÿ
		End If
		rsget.Close
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(�ٹ�), 4,5(����)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'��ü����
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	'�ӽû�ǰ �������� ���							
	Public Function getskstoaItemRegParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId							'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsName=" & getItemNameFormat()				'#��ǰ��
		strRst = strRst & "&keyword=" & getKeywords							'�˻��� | [default:��ǰ��]
'		strRst = strRst & "&arsName=" & getItemNameFormat()					'ARS�� | [default:��ǰ��]
'		strRst = strRst & "&mobileGoodsName="								'����ϻ�ǰ�� | [default:��ǰ��]
		strRst = strRst & "&entpManSeq=" & skentpManSeq						'#��ü����� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&mdCode=" & skmdCode								'#MD | MD����Ʈ ����(IF_API_00_001), �ҽ̸�ü�� ���θ��� MD�� ��1
		strRst = strRst & "&taxYn=" & CHKIIF(FVatInclude="N","0","1")		'�ǸŰ������� | 0:�鼼, 1:����
		strRst = strRst & "&codeLgroup=" & FLgroup							'#��з� | ��ǰ ��з� ��ȸ ����(IF_API_00_002)
		strRst = strRst & "&codeMgroup=" & FMgroup							'#�ߺз� | ��ǰ �ߺз� ��ȸ ����(IF_API_00_003)
		strRst = strRst & "&codeSgroup=" & FSgroup							'#�Һз� | ��ǰ �Һз� ��ȸ ����(IF_API_00_004)
		strRst = strRst & "&codeDgroup=" & FDgroup							'#���з� | ��ǰ ���з� ��ȸ ����(IF_API_00_005)
		strRst = strRst & "&shipCostCode=" & skshipCostCode					'#���δ��ۺ񱸺� | ��ۺ���å ��ȸ ����(IF_API_00_025)
		strRst = strRst & "&delyBoxQty=1"									'#���Box���� | [default:1]
		strRst = strRst & "&installYn=0"									'#��ġ��ۿ��� | 0:N, 1:Y, [default:0]
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")	'#���λ�ǰ���� | 0:N, 1:Y
		strRst = strRst & "&makecoCode=" & getMakecoCode					'#������ü | ������ü��ȸ ����(IF_API_00_019)
		strRst = strRst & "&originCode=" & getOriginCode					'#������ | ��������ȸ ����(IF_API_00_018)
		strRst = strRst & "&brandCode="	& getBrandCode						'#�귣�� | �귣����ȸ ����(IF_API_00_015)		
		strRst = strRst & "&buyPrice=" & MustBuyPrice(MustPrice())			'#���԰� | �Ҽ��� �Ұ�
		strRst = strRst & "&salePrice=" & MustPrice()						'#�ǸŰ�
		strRst = strRst & "&custPrice=" & MustPrice()						'#�����ǸŰ� ???
		strRst = strRst & "&shipManSeq=" & skshipManSeq						'#������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&returnManSeq=" & skreturnManSeq					'#ȸ������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&offerType=" & getSkstoaOfferType				'#��������Ÿ�� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)
'		strRst = strRst & "&weight="										'���� | [default:0]
'		strRst = strRst & "&volumeDesc="									'���Ǽ���
		strRst = strRst & "&formCode=F999"									'#��ǰ���� | ��ǰ���±׷� ��ȸ ����(IF_API_00_008), F999(��Ÿ)
		strRst = strRst & "&sizeCode=S999"									'#ũ�� | ��ǰũ��׷� ��ȸ ����(IF_API_00_007), S999(��Ÿ)
		strRst = strRst & "&costTaxYn=" & CHKIIF(FVatInclude="N","0","1")	'���԰������� | "�̻�� �÷����� ��ȯ - �ǸŰ����� �Է¹޾Ƽ� �ǸŰ����� ������ �����͸� �ڵ����� ������.
		strRst = strRst & "&accountManSeq=" & skaccountManSeq				'#ȸ������ | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&islandsDelyOut=0"								'#�����갣��ۺҰ� | 0: ��۰���, 1:��ۺҰ�
		strRst = strRst & "&jejuDelyOut=0"									'#����������ۺҰ� | 0: ��۰���, 1:��ۺҰ�
		strRst = strRst & "&mdKind=" & skmdKind								'#MD�з��ڵ� | MD�з�����Ʈ ����(IF_API_00_026)
		strRst = strRst & "&collectYn=0"									'#���ҿ��� | 0:N, 1:Y, [default:0]
		strRst = strRst & "&returnNoYn=0"									'#��ȯ/��ǰ�Ұ����� | 0: ��ȯ/��ǰ����, 1: ��ȯ/��ǰ�Ұ�
		strRst = strRst & "&globalDelyYn=0"									'#�ؿ�����ۿ��� | 0: N, 1: Y
'		strRst = strRst & "&globalCode="									'�����Ǵ±����ڵ� | ��������ȸ�� ����(IF_API_00_018) �ؿ�������ϰ�� �ʼ� �Է�
'		strRst = strRst & "&modelNo="										'�𵨹�ȣ | �𵨸�api�� ���Ͽ� �˻� �� ���� modelId
'		strRst = strRst & "&modelName="										'�𵨸� | �𵨸�api�� ���Ͽ� �˻� �� ���� modelName
'		strRst = strRst & "&dawnYn="										'������ۻ�ǰ���� | 0: �Ϲ�, 1: �����������, 2: ��Ÿ�������
		strRst = strRst & "&orderCreateYn=" & Chkiif(IsMakeItem()="Y", "1", "0")	'�ֹ����ۿ��� | 0:N, 1:Y, [default:0]
'		strRst = strRst & "&bookYn="										'�������� | 0: N, 1: Y, 1�� ��� ���԰���,�ǸŰ����� 0���� ���� �ʼ�, ISBN �� ���� �ʼ�
'		strRst = strRst & "&isbn="											'����ISBN | bookYn�� 1�� ��� �ʼ�
		getskstoaItemRegParameter = strRst
'		response.end
	End Function

	'�ӽû�ǰ ����� ���
	Public Function getskstoaContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId							'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#�ӽû�ǰ�ڵ�
		strRst = strRst & "&descCode=998" 									'#����� �ڵ� | ������׸� ��ȸ ����, 200 : ��ǰ����, 201 : �������, 202 : ���/��ǰ/AS�ȳ�, 301 : TV��ǰ�� ����1, 302 : TV��ǰ�� ����2, 998 : ����ϼ���, 999 : WEB����
		strRst = strRst & "&descExt=" & getSkstoaContParamToReg()			'#����� ���� | CLOB type
		getskstoaContentParameter = strRst
	End Function

	'�ӽû�ǰ ��ǰ���� ���
	Public Function getskstoaOptParameter(otherText, maxSaleQty)
		Dim strRst, strSql, optcnt, limitsu
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId							'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#�ӽû�ǰ�ڵ�
		strRst = strRst & "&colorGroupCode=C000"							'#����׷��ڵ� | ��ǰ����׷���ȸ ����(IF_API_00_006) C000 : ��ü����
		strRst = strRst & "&patternGroupCode=P999"							'#���̱׷��ڵ� | ��ǰ���̱׷���ȸ ����(IF_API_00_009) P999 : ��Ÿ
'		strRst = strRst & "&colorCode="										'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&patternCode="									'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&sizeCode="										'ũ���ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&formCode="										'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
		strRst = strRst & "&otherText=" & URLEncodeUTF8Plus(otherText)		'��ǰ��Ÿ | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1, ��ǰ�� �����å : �ѱ�,����,����, [~] [!] [*] [-][_][=][+][|][:][.][/][?], \\^\\(\\)\\[\\]\\{\\}]
		strRst = strRst & "&maxSaleQty=" & maxSaleQty						'#�ִ��Ǹż��� | ���ڸ� �Է°���		
		getskstoaOptParameter = strRst
	End Function

	'�ӽû�ǰ �̹��� ���(URL)
	Public Function getskstoaImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId							'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#�ӽû�ǰ�ڵ�
		strRst = strRst & "&imgUrlBase=" & FbasicImage 						'#�����̹��� URL
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					Select Case i
						Case "1"		imgurlparam = "&imgUrlA"
						Case "2"		imgurlparam = "&imgUrlB"
						Case "3"		imgurlparam = "&imgUrlC"
						Case "4"		imgurlparam = "&imgUrlD"
					End Select
					strRst = strRst & imgurlparam &"=http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				End If
				rsget.MoveNext
				If i >= 4 Then Exit For
			Next
		End If
		rsget.Close
		getskstoaImageParameter = strRst
	End Function

	'�ӽû�ǰ ����������� ���
	Public Function getskstoaGosiRegParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","����")

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId								'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#�ӽû�ǰ�ڵ�
		strRst = strRst & "&typeCode=" & mallinfodiv							'#��ǰ�����ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)
		strRst = strRst & "&offerCode=" & mallinfocd							'#�׸��ڵ� | ��ǰ����������� ǰ�� �׸� ����(IF_API_00_023)
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)	'#��������� | CLOB type
		getskstoaGosiRegParameter = strRst
	End Function

	'�ӽû�ǰ �����������
	Public Function getskstoaCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum, notSendinfoDiv
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId								'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#�ӽû�ǰ�ڵ�
		strRst = strRst & "&typeCode=" & getSkstoaOfferType						'#��ǰ�����ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)
' - �Ʒ� �׸��� ��� ������ȣ�� ���� ����
' 16.�Ƿ���
' 18.ȭ��ǰ
' 20.��ǰ(����깰)
' 21.������ǰ
' 22.�ǰ���ɽ�ǰ
' 27.ȣ��/��� ����
' 28.������Ű��
' 29.�װ���
' 30.�ڵ��� �뿩 ����
' 33.������ ������(����/����/���ͳݰ��� ��)
' 34.��ǰ��/����
' 35.���������
' 36.��ȭ/����
' 37.��Ÿ�뿪

		Select Case getSkstoaOfferType
			Case "16", "18", "20", "21", "22", "27", "28", "29", "30", "33", "34", "35", "36", "37"
				notSendinfoDiv = "Y"
			Case Else
				notSendinfoDiv = "N"
		End Select

		If notSendinfoDiv = "Y" Then
			strRst = strRst & "&kcYn=0"
			strRst = strRst & "&kcNo="
		Else
			strSql = ""
			strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isnull(t.certNum, '') as certNum "
			strSql = strSql & " FROM db_item.dbo.tbl_item as i "
			strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
			strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
			strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				safetyDiv	= rsget("safetyDiv")
				certNum		= rsget("certNum")
				isRegCert	= "Y"
			Else
				isRegCert	= "N"
			End If
			rsget.Close

			If (safetyDiv = "10" OR safetyDiv = "20" OR safetyDiv = "40" OR safetyDiv = "50") AND (certNum <> "") Then
				strRst = strRst & "&kcYn=1"
				strRst = strRst & "&kcNo=" & certNum
			Else
				strRst = strRst & "&kcYn=0"
				strRst = strRst & "&kcNo="
			End If
		End If
		getskstoaCertParameter = strRst
	End Function

	'�ӽû�ǰ ���ο�û
	Public Function getskstoaConfirmParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId								'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#�ӽû�ǰ�ڵ�
		getskstoaConfirmParameter = strRst
	End Function

	'��ǰ �Ǹ��ߴ� ó��
	Public Function getSkstoaSellynParameter(ichgSellYn)
		Dim strRst
		'saleNoCode
		'https://wapi.10x10.co.kr/outmall/skstoa/skstoaActProc.asp?act=commonCode&interfaceId=IF_API_00_021
		'101 : ��ü�ε�, 102 : ��ǰ���޺Ҿ���, 103 : ����ó������, 104 : ǰ������, 105 : �����ߴ�, 999 : �ŷ�����

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		If ichgSellYn = "Y" Then
			'��ü�� �� �Ǹŷ� �Ѵ�? Ȯ���غ�����..
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=00"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=" 							'#�ǸźҰ��ڵ� | �ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		ElseIf ichgSellYn = "N" Then
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=11"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=105" 						'#�ǸźҰ��ڵ� | �ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		ElseIf ichgSellYn = "X" Then
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=19"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=105"							'#�ǸźҰ��ڵ� | �ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�
			strRst = strRst & "&saleNoNote=�Ǹ�����"					'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		End If
		getSkstoaSellynParameter = strRst
	End Function

	'�ǸŻ�ǰ ��ȸ(��)
	Public Function getSkstoaItemViewParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
'		strRst = strRst & "&bDate="										'#��ȸ �������� | ����� ����  YYYYMMDDŸ��. ex) 20130118"		
'		strRst = strRst & "&eDate="										'#��ȸ ���������� | ����� ���� YYYYMMDDŸ��. ex) 20130118"		
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'��ǰ�ڵ� ������ȸ. �ڵ� ��ȸ�� ����� ���� ���� ����
		getSkstoaItemViewParameter = strRst
	End Function

	'�ǸŻ�ǰ �������� ����
	Public Function getskstoaItemEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ� | ��ǰ������ �ʼ�
		strRst = strRst & "&goodsName=" & getItemNameFormat()			'#��ǰ��
		strRst = strRst & "&keyword=" & getKeywords						'#�˻��� | [default:��ǰ��]
		strRst = strRst & "&arsName=" & getItemNameFormat()				'#ARS�� | [default:��ǰ��]
'		strRst = strRst & "&mobileGoodsName="							'#����ϻ�ǰ�� | [default:��ǰ��
'		strRst = strRst & "&weight="									'#���� | [default:0]
'		strRst = strRst & "&volumeDesc="								'#���Ǽ���
		strRst = strRst & "&shipCostCode=" & skshipCostCode				'#���δ��ۺ񱸺��ڵ� | ��ۺ���å��ȸ ����(IF_API_00_025)
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")		'#���λ�ǰ���� | 0:N, 1:Y
		strRst = strRst & "&orderMinQty=1"								'#�ֹ��ּҼ���
		strRst = strRst & "&orderMaxQty=" & getOrderMaxNum				'#�ֹ��ִ����
		strRst = strRst & "&collectYn=0"								'#���ҿ��� | 0:N, 1:Y, [default:0]
		strRst = strRst & "&returnNoYn=0"								'#��ȯ/��ǰ�Ұ����� | 0: ��ȯ/��ǰ����, 1: ��ȯ/��ǰ�Ұ�
'		strRst = strRst & "&modelNo="									'�𵨹�ȣ
'		strRst = strRst & "&modelName="									'�𵨸�
		getskstoaItemEditParameter = strRst
	End Function

	'���»� ���ݵ��
	Public Function getskstoaEditPriceParameter()
		Dim strRst
		Dim saleVat, buyPrice, buyCost, buyVat
		buyPrice	= MustBuyPrice(MustPrice())
		buyVat		= REPLACE(Formatnumber(buyPrice / 11, 0), ",", "")
		buyCost		= buyPrice - buyVat
		saleVat		= REPLACE(Formatnumber(MustPrice / 11, 0), ",", "")

		If FVatInclude = "N" Then
			buyVat	= 0
			buyCost = buyPrice
			saleVat	= 0
		End If

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&applyDate="	& Replace(DATE(), "-", "/")		'#�������� | YYYY/MM/DDŸ��. ex) 2018/09/06
		strRst = strRst & "&salePrice=" & MustPrice						'#�ǸŰ�
		strRst = strRst & "&buyPrice="& buyPrice						'���԰�
		getskstoaEditPriceParameter = strRst
	End Function

	'�ǸŻ�ǰ ����� ���
	Public Function getskstoaEditContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&descCode=998" 								'#������ڵ� | ������׸� ��ȸ ����, 200 : ��ǰ����, 201 : �������, 202 : ���/��ǰ/AS�ȳ�, 301 : TV��ǰ�� ����1, 302 : TV��ǰ�� ����2, 998 : ����ϼ���, 999 : WEB����
		strRst = strRst & "&descExt=" & getSkstoaContParamToReg()		'#��������� | CLOB type
		getskstoaEditContentParameter = strRst
	End Function

	'�ǸŻ�ǰ �̹��� ���(URL)
	Public Function getskstoaEditImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&imgUrlBase=" & FbasicImage 					'�����̹��� URL
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					Select Case i
						Case "1"		imgurlparam = "&imgUrlA"
						Case "2"		imgurlparam = "&imgUrlB"
						Case "3"		imgurlparam = "&imgUrlC"
						Case "4"		imgurlparam = "&imgUrlD"
					End Select
					strRst = strRst & imgurlparam &"=http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				End If
				rsget.MoveNext
				If i >= 4 Then Exit For
			Next
		End If
		rsget.Close
		getskstoaEditImageParameter = strRst
	End Function

	'�ǸŻ�ǰ ����������� ���
	Public Function getskstoaGosiEditParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","����")

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&typeCode=" & mallinfodiv					'#Ÿ���ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)		
		strRst = strRst & "&offerCode=" & mallinfocd					'#�׸��ڵ� | ��ǰ����������� ǰ�� �׸� ����(IF_API_00_023)		
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)				'#�׸񳻿� | CLOB type
		getskstoaGosiEditParameter = strRst
	End Function

	'�ǸŻ�ǰ �����������
	Public Function getskstoaEditCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum, notSendinfoDiv
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&typeCode=" & getSkstoaOfferType				'#��ǰ�����ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)
' - �Ʒ� �׸��� ��� ������ȣ�� ���� ����
' 16.�Ƿ���
' 18.ȭ��ǰ
' 20.��ǰ(����깰)
' 21.������ǰ
' 22.�ǰ���ɽ�ǰ
' 27.ȣ��/��� ����
' 28.������Ű��
' 29.�װ���
' 30.�ڵ��� �뿩 ����
' 33.������ ������(����/����/���ͳݰ��� ��)
' 34.��ǰ��/����
' 35.���������
' 36.��ȭ/����
' 37.��Ÿ�뿪
		Select Case getSkstoaOfferType
			Case "16", "18", "20", "21", "22", "27", "28", "29", "30", "33", "34", "35", "36", "37"
				notSendinfoDiv = "Y"
			Case Else
				notSendinfoDiv = "N"
		End Select

		If notSendinfoDiv = "Y" Then
			strRst = strRst & "&kcYn=0"
			strRst = strRst & "&kcNo="
		Else
			strSql = ""
			strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isnull(t.certNum, '') as certNum "
			strSql = strSql & " FROM db_item.dbo.tbl_item as i "
			strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
			strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
			strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				safetyDiv	= rsget("safetyDiv")
				certNum		= rsget("certNum")
				isRegCert	= "Y"
			Else
				isRegCert	= "N"
			End If
			rsget.Close

			If (safetyDiv = "10" OR safetyDiv = "20" OR safetyDiv = "40" OR safetyDiv = "50") AND (certNum <> "") Then
				strRst = strRst & "&kcYn=1"
				strRst = strRst & "&kcNo=" & certNum
			Else
				strRst = strRst & "&kcYn=0"
				strRst = strRst & "&kcNo="
			End If
		End If
		getskstoaEditCertParameter = strRst
	End Function

	'�ǸŻ�ǰ �����
	Public Function geskstoaOptionQtyParam(outmallOptCode, optsu)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#�ǸŴ�ǰ�ڵ�
		strRst = strRst & "&inplanQty=" & optsu							'#�ǸŰ��ɼ���
		geskstoaOptionQtyParam = strRst
	End Function

	'��ǰ �Ǹ��ߴ� ó��
	Public Function geskstoaOptionStatParam(outmallOptCode, isalegb)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#�ǸŴ�ǰ�ڵ�
		strRst = strRst & "&saleGb=" & isalegb							'#�Ǹű���
		If isalegb = "11" Then
			strRst = strRst & "&saleNoCode=105" 						'#�Ұ����� �ڵ� | "�ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�"
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		End If
		geskstoaOptionStatParam = strRst
	End Function

	'�ǸŻ�ǰ ��ǰ���� ���
	Public Function getskstoaOptionAddParam(otherText, maxSaleQty)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#�����ڵ� | SKB���� �ο��� �����ڵ�
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#��ü�ڵ� | SKB���� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & skstoaentpId						'#��ü�����ID | SKB���� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#��üPASSWORD | SKB���� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#�ǸŻ�ǰ�ڵ�
'		strRst = strRst & "&sizeCode="									'ũ���ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&colorCode="									'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&patternCode="								'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&formCode="									'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
		strRst = strRst & "&dtText=" & URLEncodeUTF8Plus(otherText)		'�ؽ�Ʈ�Է� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
		strRst = strRst & "&maxSaleQty=" & maxSaleQty					'�ִ��Ǹż��� | ���ڸ� �Է°���
		getskstoaOptionAddParam = strRst
	End Function
End Class

Class CSkstoa
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

	Public Sub getskstoaNotRegOneItem
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
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.skstoaStatCD,-9) as skstoaStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		strSql = strSql & "	, isnull(am.lgroup, '') as lgroup "
		strSql = strSql & "	, isnull(am.mgroup, '') as mgroup "
		strSql = strSql & "	, isnull(am.sgroup, '') as sgroup "
		strSql = strSql & "	, isnull(am.dgroup, '') as dgroup "
		strSql = strSql & "	, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_skstoa_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_category as tm on am.lgroup = tm.lgroup and am.mgroup = tm.mgroup and am.sgroup = tm.sgroup and am.dgroup = tm.dgroup "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_regItem as R on i.itemid = R.itemid"
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
		strSql = strSql & " and i.itemdiv in ('01', '16', '07') "		'01 : �Ϲ�, 16 : �ֹ�����, 07 : �������� / �ż���tvȨ������ �ֹ����� ����(06) �Ұ�!
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
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_skstoa_regItem WHERE skstoaStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.	'skstoa��ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql																				'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSkstoaItem
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
				FOneItem.FSocname_kor		= rsget("socname_kor")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FSkstoaStatCD		= rsget("skstoaStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FLgroup			= rsget("lgroup")
				FOneItem.FMgroup			= rsget("mgroup")
				FOneItem.FSgroup			= rsget("sgroup")
				FOneItem.FDgroup			= rsget("dgroup")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOutmallstandardMargin = rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub getSkstoaTmpRegedOneItem(isSendConfrm)
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.itemid, r.skstoaTmpGoodNo, i.smallImage, i.basicImage, i.mainimage, i.mainimage2, c.itemcontent "
		strSql = strSql & " ,ordercomment, isNull(r.reglevel, 0) as reglevel, i.limityn, i.limitno, i.limitsold, C.infoDiv "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_skstoa_regItem as r "
		strSql = strSql & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and r.itemid = '"& FRectItemID &"' "
		strSql = strSql & " and isNull(skstoaTmpGoodNo, '') <> '' "
		If isSendConfrm = "Y" Then
			strSql = strSql & " and sendConfirm = 'Y' "
		End If
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSkstoaItem
				FOneItem.FItemid					= rsget("itemid")
				FOneItem.FSkstoaTmpGoodNo			= rsget("skstoaTmpGoodNo")
				FOneItem.FsmallImage				= rsget("smallImage")
				FOneItem.FbasicImage				= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage					= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2				= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
                FOneItem.FbasicimageNm 				= rsget("basicimage")
				FOneItem.FReglevel 					= rsget("reglevel")
				FOneItem.FItemcontent				= db2html(rsget("itemcontent"))
				FOneItem.FOrdercomment				= db2html(rsget("ordercomment"))
				FOneItem.FLimityn					= rsget("limityn")
				FOneItem.FLimitno					= rsget("limitno")
				FOneItem.FLimitsold					= rsget("limitsold")
				FOneItem.FinfoDiv					= rsget("infoDiv")
		End If
		rsget.Close
	End Sub

	Public Sub getSkstoaEditOneItem
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
		strSql = strSql & "	, m.skstoaTmpGoodNo, m.skstoaGoodNo, m.skstoaprice, m.skstoaSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & "	, isnull(am.lgroup, '') as lgroup "
		strSql = strSql & "	, isnull(am.mgroup, '') as mgroup "
		strSql = strSql & "	, isnull(am.sgroup, '') as sgroup "
		strSql = strSql & "	, isnull(am.dgroup, '') as dgroup "
		strSql = strSql & "	, isNULL(m.skstoaStatCD,-9) as skstoaStatCD, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
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
		strSql = strSql & "		or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "		or isnull(am.lgroup, '') = '' "		'ī�װ� �̸���
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_skstoa_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_skstoa_category as tm on am.lgroup = tm.lgroup and am.mgroup = tm.mgroup and am.sgroup = tm.sgroup and am.dgroup = tm.dgroup "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.skstoaGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CSkstoaItem
				FOneItem.Fitemid				= rsget("itemid")
				FOneItem.FtenCateLarge			= rsget("cate_large")
				FOneItem.FtenCateMid			= rsget("cate_mid")
				FOneItem.FtenCateSmall			= rsget("cate_small")
				FOneItem.Fitemname				= db2html(rsget("itemname"))
				FOneItem.FitemDiv				= rsget("itemdiv")
				FOneItem.FsmallImage			= rsget("smallImage")
				FOneItem.Fmakerid				= rsget("makerid")
				FOneItem.Fregdate				= rsget("regdate")
				FOneItem.FlastUpdate			= rsget("lastUpdate")
				FOneItem.ForgPrice				= rsget("orgPrice")
				FOneItem.ForgSuplyCash			= rsget("orgSuplyCash")
				FOneItem.FSellCash				= rsget("sellcash")
				FOneItem.FBuyCash				= rsget("buycash")
				FOneItem.FsellYn				= rsget("sellYn")
				FOneItem.FsaleYn				= rsget("sailyn")
				FOneItem.FisUsing				= rsget("isusing")
				FOneItem.FLimitYn				= rsget("LimitYn")
				FOneItem.FLimitNo				= rsget("LimitNo")
				FOneItem.FLimitSold				= rsget("LimitSold")
				FOneItem.Fkeywords				= rsget("keywords")
				FOneItem.ForderComment			= db2html(rsget("ordercomment"))
				FOneItem.FoptionCnt				= rsget("optionCnt")
				FOneItem.FbasicImage			= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage				= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2			= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea			= rsget("sourcearea")
				FOneItem.Fmakername				= rsget("makername")
				FOneItem.FUsingHTML				= rsget("usingHTML")
				FOneItem.Fitemcontent			= db2html(rsget("itemcontent"))
				FOneItem.FSkstoaTmpGoodNo		= rsget("skstoaTmpGoodNo")
				FOneItem.FSkstoaGoodNo			= rsget("skstoaGoodNo")
				FOneItem.FSkstoaprice			= rsget("skstoaprice")
				FOneItem.FSkstoaSellYn			= rsget("skstoaSellYn")

                FOneItem.FoptionCnt       		= rsget("optionCnt")
                FOneItem.FregedOptCnt     		= rsget("regedOptCnt")
                FOneItem.FaccFailCNT      		= rsget("accFailCNT")
                FOneItem.FlastErrStr      		= rsget("lastErrStr")
                FOneItem.Fdeliverytype    		= rsget("deliverytype")
                FOneItem.FrequireMakeDay  		= rsget("requireMakeDay")

                FOneItem.FinfoDiv       		= rsget("infoDiv")
                FOneItem.Fsafetyyn      		= rsget("safetyyn")
                FOneItem.FsafetyDiv     		= rsget("safetyDiv")
                FOneItem.FsafetyNum     		= rsget("safetyNum")
                FOneItem.FmaySoldOut    		= rsget("maySoldOut")

                FOneItem.FDeliveryType			= rsget("deliveryType")
                FOneItem.Fregitemname			= rsget("regitemname")
                FOneItem.FregImageName			= rsget("regImageName")
                FOneItem.FbasicImageNm			= rsget("basicimage")
				FOneItem.FOrderMaxNum 			= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin = rsget("outmallstandardMargin")
				FOneItem.Fvatinclude			= rsget("vatinclude")
		End If
		rsget.Close
	End Sub
End Class

Function GetRaiseValue(value)
    If Fix(value) < value Then
    	GetRaiseValue = Fix(value) + 1
    Else
    	GetRaiseValue = Fix(value)
    End If
End Function

Function getOptionList(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Skstoa_ItemOptionMapping_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getOptionList = rsget.getRows
	end if
	rsget.Close
End Function

Function getInfoCodeMapList(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_skstoa_InfoCodeMap_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getInfoCodeMapList = rsget.getRows
	end if
	rsget.Close
End Function

Function getOptiopnMapList(iitemid, gbn)
	Dim strSql
	strSql = ""
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Skstoa_OptionMappingByEdit_Get] '"&iitemid&"', '"& gbn &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getOptiopnMapList = rsget.getRows
	end if
	rsget.Close
End Function
 
Function getOptiopnMayAddList(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Skstoa_OptionMappingByAdd_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getOptiopnMayAddList = rsget.getRows
	end if
	rsget.Close
End Function

Function getSkstoaOptCnt(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption "
	strSql = strSql & " WHERE mallid = '"& CMALLNAME &"' "
	strSql = strSql & " and itemid = '"& iitemid &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getSkstoaOptCnt = rsget("cnt")
	end if
	rsget.Close
End Function

Function getMayCertYn(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, t.certNum "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
	strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
	strSql = strSql & " WHERE i.itemid = '"& iitemid &"' "
	strSql = strSql & " and t.safetyDiv in ('10', '20', '40', '50', '70', '80') "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.Eof Then
		getMayCertYn	= "Y"
	Else
		getMayCertYn	= "N"
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

function skstoaAPIURL()
	If application("Svr_Info") = "Dev" Then
		skstoaAPIURL = "http://dev-sel.skstoa.com"
	Else
		skstoaAPIURL = "https://open-api.skstoa.com"
	End If
end function

function skstoalinkCode()
	If application("Svr_Info") = "Dev" Then
		skstoalinkCode = "TENBYTEN"
	Else
		skstoalinkCode = "TENBYTEN"
	End If
end function

function skstoaentpCode()
	If application("Svr_Info") = "Dev" Then
		skstoaentpCode = "112644"
	Else
		skstoaentpCode = "112644"
	End If
end function

function skstoaentpId()
	If application("Svr_Info") = "Dev" Then
		skstoaentpId = "E112644"
	Else
		skstoaentpId = "E112644"
	End If
end function

function skmdCode()
	If application("Svr_Info") = "Dev" Then
		skmdCode = "0263"		'������
	Else
		skmdCode = "0263"		'������
	End If
end function

function skentpManSeq()
	If application("Svr_Info") = "Dev" Then
		skentpManSeq = "002"
	Else
		skentpManSeq = "002"
	End If
end function

function skshipManSeq()
	If application("Svr_Info") = "Dev" Then
		skshipManSeq = "006"
	Else
		skshipManSeq = "006"
	End If
end function

function skreturnManSeq()
	If application("Svr_Info") = "Dev" Then
		skreturnManSeq = "007"
	Else
		skreturnManSeq = "007"
	End If
end function

function skaccountManSeq()
	If application("Svr_Info") = "Dev" Then
		skaccountManSeq = "004"
	Else
		skaccountManSeq = "004"
	End If
end function

function skshipCostCode()
	If application("Svr_Info") = "Dev" Then
		skshipCostCode = "CN001"
	Else
		skshipCostCode = "CN001"
	End If
end function

function skmdKind()
	If application("Svr_Info") = "Dev" Then
		skmdKind = "0005"			'��Ȱ��ǰ
	Else
		skmdKind = "0005"			'��Ȱ��ǰ
	End If
end function

function skstoaentpPass()
	Dim skstoaStrSql
	skstoaStrSql = ""
	skstoaStrSql = skstoaStrSql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
	skstoaStrSql = skstoaStrSql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
	skstoaStrSql = skstoaStrSql & " where mallid='skstoa' " & VbCRLF
	skstoaStrSql = skstoaStrSql & " and inikey='pass'"
	rsget.CursorLocation = adUseClient
	rsget.Open skstoaStrSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		skstoaentpPass	= rsget("iniVal")
	end if
	rsget.close
end function


%>