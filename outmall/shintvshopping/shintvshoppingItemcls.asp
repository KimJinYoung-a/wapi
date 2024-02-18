<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "shintvshopping"
CONST CUPJODLVVALID = TRUE								''��ü ���ǹ�� ��� ���ɿ���
CONST CMAXLIMITSELL = 5									'' �� ���� �̻��̾�� �Ǹ���. // �ɼ������� ��������.

Class CShintvshoppingItem
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
	Public FShintvshoppingStatCD
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
	Public FShintvshoppingGoodNo
	Public FShintvshoppingprice
	Public FShintvshoppingSellYn

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
			getOrderMaxNum = 9999
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
		strRst = FKeywords
		strRst = replace(strRst, "�α�", "")
		strRst = replace(strRst, "��ġ", "")
		strRst = replace(strRst, "�����ġ", "")
		If strRst = "" Then
			strRst = "�ٹ�����"
		End If
		getKeywords = Server.URLEncode(strRst)
	End Function

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

			If FShintvshoppingPrice = 0 Then
				If (GetTenTenMargin < outmallstandardMargin) Then
					tmpPrice = Forgprice
				Else
					tmpPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < outmallstandardMargin Then
					If (Forgprice < Round(FShintvshoppingPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FShintvshoppingPrice * 0.35, 0)/10)*10)
					ElseIf Clng(Forgprice) > Clng(Round(FShintvshoppingPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FShintvshoppingPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = Forgprice
					End If
				Else
					If (FSellCash < Round(FShintvshoppingPrice * 0.35, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FShintvshoppingPrice * 0.35, 0)/10)*10)
					ElseIf Clng(FSellCash) > Clng(Round(FShintvshoppingPrice * 1.65, 0)) Then
						tmpPrice = CStr(GetRaiseValue(Round(FShintvshoppingPrice * 1.65, 0)/10)*10)
					Else
						tmpPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'// Shintvshopping �Ǹſ��� ��ȯ
	Public Function getShintvshoppingSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getShintvshoppingSellYn = "Y"
			Else
				getShintvshoppingSellYn = "N"
			End If
		Else
			getShintvshoppingSellYn = "N"
		End If
	End Function

	'// Shintvshopping �Ǹſ��� ��ȯ
	Public Function getShintvshoppingOfferType()
		Dim buf
		Select Case FinfoDiv
			Case "35"	buf = "38"
			Case "36"	buf = "35"
			Case "47"	buf = "39"
			Case "48"	buf = "40"
			Case Else	buf = FinfoDiv
		End Select
		getShintvshoppingOfferType = buf
	End Function

	Public Function fnShipCostCode()
		Dim buf, sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 shipCostCode "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_shintvshopping_beasongCodeItem_master] m "
		sqlStr = sqlStr & " JOIN db_etcmall.[dbo].[tbl_shintvshopping_beasongCodeItem_detail] d on m.idx = d.midx "
		sqlStr = sqlStr & " WHERE m.isusing = 'Y' "
		sqlStr = sqlStr & " and GETDATE() between m.startDate and m.enddate "
		sqlStr = sqlStr & " and d.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			fnShipCostCode	= Trim(rsget("shipCostCode"))
		Else
			fnShipCostCode = shipCostCode
		End If
		rsget.Close
	End Function

	Public Function getShintvshoppingContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_shintvshopping.jpg""></p><br>")
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
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_shintvshopping.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		getShintvshoppingContParamToReg = strRst
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

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Shintvshopping_OptionValid_Get] " & FItemid
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
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_shintvshopping_makeCompanyCode] "
		strSql = strSql & " WHERE makeCompanyName like '%"& html2db(Fmakername) &"%' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getMakecoCode = rsget("makeCompanyCode")
		Else
			getMakecoCode = makecoCode
		End If
		rsget.Close
	End Function

	Function getOriginCode()
		Dim strSql
		strSql = strSql & " SELECT TOP 1 originCode "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_shintvshopping_originCode] "
		strSql = strSql & " WHERE originName like '%"& html2db(Fsourcearea) &"%' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			getOriginCode = rsget("originCode")
		Else
			getOriginCode = originCode
		End If
		rsget.Close
	End Function

	Function getBrandCode()
		getBrandCode = brandCode	'2023-06-08 ������ ����..������������ brandCode �� ����
		' Dim strSql
		' strSql = strSql & " SELECT TOP 1 brandCode "
		' strSql = strSql & " FROM db_etcmall.[dbo].[tbl_shintvshopping_brandCode] "
		' strSql = strSql & " WHERE brandName = '"& html2db(FSocname_kor) &"' "
		' rsget.CursorLocation = adUseClient
		' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		' If (Not rsget.EOF) Then
		' 	getBrandCode = rsget("brandCode")
		' Else
		' 	getBrandCode = brandCode
		' End If
		' rsget.Close
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

	Public Function getShopLeadTime()
		If FItemdiv = "06" OR FItemdiv = "16" Then
			getShopLeadTime = 15
		Else
			If CStr(FtenCateLarge) = "040" Then
				getShopLeadTime = 15
			Else
				getShopLeadTime = 7
			End If
		End If
	End Function

	'�ӽû�ǰ �������� ���_v2								
	Public Function getshintvshoppingItemRegParameter(iShipcostCode)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsName=" & getItemNameFormat()			'#��ǰ�� | . : , ; ( ! ? ) + - * / = [ ] Ư������ �Է� ����
'		strRst = strRst & "&arsName=" & getItemNameFormat()				'ARS�� | [default:��ǰ��]
'		strRst = strRst & "&slipNamePrintYn=0							'������ǰ�� ��¿��� | 0:N, 1:Y [default:0]
'		strRst = strRst & "&slipName=" & getItemNameFormat()			'������ǰ�� | ������ǰ�� ��¿��ΰ� 1�� ��� �Է°����ϸ�,  ������ǰ�� �ʼ� �Է�
'		strRst = strRst & "&mobileGoodsName=" & getItemNameFormat()		'����ϻ�ǰ�� | [default:��ǰ��]	
		strRst = strRst & "&entpManSeq=" & entpManSeq					'#��ü����� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&mdCode=" & mdCode							'#MD | MD����Ʈ ����(IF_API_00_001)
		strRst = strRst & "&taxYn=" & CHKIIF(FVatInclude="N","0","1")	'#��������(�ǸŰ�������) | 0:�鼼, 1:����
		strRst = strRst & "&codeLgroup=" & FLgroup						'#CAT | �ű� ��ǰCAT��ȸ ����(IF_API_00_002)
		strRst = strRst & "&codeMgroup=" & FMgroup						'#��з� | �ű� ��ǰ��з���ȸ ����(IF_API_00_003)
		strRst = strRst & "&codeSgroup=" & FSgroup						'#�ߺз� | �ű� ��ǰ�ߺз���ȸ ����(IF_API_00_004)
		strRst = strRst & "&codeDgroup=" & FDgroup						'#�Һз� | �ű� ��ǰ�Һз���ȸ ����(IF_API_00_005)
		strRst = strRst & "&codeTgroup=" & FTgroup						'#���з� | �ű� ��ǰ���з���ȸ ����(IF_API_00_028)
		strRst = strRst & "&shipCostCode=" & iShipcostCode				'#��ۺ���å�ڵ� | "��ۺ���å ��ȸ ����(IF_API_00_025) ���ҹ�ۿ��ΰ� 1�� ��� ��������å[A01 �Ǵ� A001]���� ����
		strRst = strRst & "&delyBoxQty=1"								'#��۹ڽ����� | [default:1]
		strRst = strRst & "&mixPackYn=1"								'�����尡�ɿ��� | 0:N, 1:Y [default:0]
		strRst = strRst & "&installYn=0"								'#��ġ��ۿ��� | 0:N, 1:Y, [default:0]
		strRst = strRst & "&codYn=0"									'���ҹ�ۿ��� | 0:N, 1:Y, [default:0] ��ġ��ۿ��ΰ� 1�� ��� ���ҹ�ۿ��� �Է� ����
		strRst = strRst & "&groupGoods="& Chkiif(IsMakeItem()="Y", "80", "")	'�׷��ǰ | 40: �ؿܱ��Ŵ���, 80: �ֹ����� ���� ������ �� ��� �ش�Ӽ� ��� �Ұ�
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")		'#���λ�ǰ���� | 0:N, 1:Y
		strRst = strRst & "&makecoCode=" & getMakecoCode				'#������ü | ������ü��ȸ ����(IF_API_00_019)
		strRst = strRst & "&originCode=" & getOriginCode				'#������ | ��������ȸ ����(IF_API_00_018)
		strRst = strRst & "&oemEntpName="								'OEM��� | �������� �ѱ��� �ƴϰ� ������ü�Ը� �߼ұ���̸� �ʼ��Է�
		strRst = strRst & "&brandCode=" & getBrandCode					'#�귣�� | �귣����ȸ ����(IF_API_00_015)
		strRst = strRst & "&buyPrice=" & Clng(MustPrice()*0.88)			'#���԰�
		strRst = strRst & "&salePrice=" & MustPrice						'#�ǸŰ�
		strRst = strRst & "&shipManSeq=" & shipManSeq					'#������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&returnManSeq=" & returnManSeq				'#ȸ������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&offerType="	& getShintvshoppingOfferType	'#��������Ÿ�� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)
'		strRst = strRst & "&weight="									'���� | [default:0]
'		strRst = strRst & "&vWidth="									'���� | [default:0] (����:cm)
'		strRst = strRst & "&vLength="									'���� | [default:0] (����:cm)
'		strRst = strRst & "&vHeight="									'���� | [default:0] (����:cm)
		strRst = strRst & "&costTaxYn=" & CHKIIF(FVatInclude="N","0","1")	'#���԰������� | 0:�鼼, 1:����
		strRst = strRst & "&taxSmallYn=0"								'�������� | 0:�Ϲ�, 1:���� (DEFAULT:0 �Ϲ�)
		strRst = strRst & "&parallelImportYn=0"							'������Կ��� | 0:N, 1:Y (DEFAULT:0)
		strRst = strRst & "&modifier="									'���ľ�
		strRst = strRst & "&doNotIslandDelyYn=0"						'����/�갣 ��ۺҰ� ���� | 0: ��۰���, 1: ��� �Ұ� [default : 0]
		strRst = strRst & "&doNotJejuDelyYn=0"							'���� ��ۺҰ� ���� | 0: ��۰���, 1: ��� �Ұ� [default : 0]
		strRst = strRst & "&unitGoodsYn="								'��ǰ�ɼǱ��� | 
		strRst = strRst & "&optionGroupCode1="							'�ɼǱ׷�1�ڵ� | 
		strRst = strRst & "&optionGroupName1="							'�ɼǱ׷�1�� | 
		strRst = strRst & "&optionGroupCode2="							'�ɼǱ׷�2�ڵ� | 
		strRst = strRst & "&optionGroupName2="							'�ɼǱ׷�2�� | 
		strRst = strRst & "&optionGroupCode3="							'�ɼǱ׷�3�ڵ� | 
		strRst = strRst & "&optionGroupName3="							'�ɼǱ׷�3�� | 
		strRst = strRst & "&optionGroupCode4="							'�ɼǱ׷�4�ڵ� | 
		strRst = strRst & "&optionGroupName4="							'�ɼǱ׷�4�� | 
		strRst = strRst & "&formCode=F999"								'�����ڵ� | 
		strRst = strRst & "&sizeCode=S999"								'ũ���ڵ� | 
		strRst = strRst & "&suGoodsCode=" & FItemid						'�������Ȼ�ǰ�ڵ� | ������ü ���� ��ǰ�ڵ�
'		strRst = strRst & "&mdManId=" & mdManId							'���MD ID | ���MD ��ȸ ����(IF_API_00_029)		// 2022-07-19 15:00 �������� ���� ��û
'		strRst = strRst & "&avgDelyLeadtime=" & getShopLeadTime()		'��ۼҿ���
		strRst = strRst & "&avgDelyLeadtime=5"							'��ۼҿ���
		getshintvshoppingItemRegParameter = strRst
'		response.end
	End Function

	'�ӽû�ǰ ����� ���
	Public Function getshintvshoppingContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�
		strRst = strRst & "&descCode=998" 								'#����� �ڵ� | ����� ��ȸ ����(IF_API_00_016) | 101 : ��ǰ����, 301: ��۾ȳ�, 302 : ��ǰ/��ȯ�ȳ�, 303 : AS�ȳ�, 997 : ����ϱ����(QS), 998 : ����ϱ����
		strRst = strRst & "&descExt=" & getShintvshoppingContParamToReg()	'#����� ����
		getshintvshoppingContentParameter = strRst
	End Function

	'�ӽû�ǰ ��ǰ���� ���
	Public Function getshintvshoppingOptParameter(otherText, maxSaleQty)
		Dim strRst, strSql, optcnt, limitsu
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�
'		strRst = strRst & "&colorGroupCode="							'#����׷��ڵ� | ��ǰ����׷���ȸ ����(IF_API_00_006)
'		strRst = strRst & "&patternGroupCode="							'#���̱׷��ڵ� | ��ǰ���̱׷���ȸ ����(IF_API_00_009)
'		strRst = strRst & "&colorCode="									'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&patternCode="								'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&sizeCode="									'ũ���ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&formCode="									'�����ڵ� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
		strRst = strRst & "&otherText=" & URLEncodeUTF8Plus(otherText)	'��ǰ��Ÿ | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
'		strRst = strRst & "&modelName="									'�𵨸�
		strRst = strRst & "&maxSaleQty=" & maxSaleQty					'#�ִ��Ǹż��� | ���ڸ� �Է°���		
		getshintvshoppingOptParameter = strRst
	End Function

	'�ӽû�ǰ �̹��� ���(URL)
	Public Function getshintvshoppingImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�
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
		getshintvshoppingImageParameter = strRst
	End Function

	'�ӽû�ǰ ����������� ���
	Public Function getshintvshoppingGosiRegParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","����")

		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�
		strRst = strRst & "&typeCode=" & mallinfodiv					'#��ǰ�����ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)		
		strRst = strRst & "&offerCode=" & mallinfocd					'#�׸��ڵ� | ��ǰ����������� ǰ�� �׸� ����(IF_API_00_023)		
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)	'�׸񳻿�
		getshintvshoppingGosiRegParameter = strRst
	End Function

	'�ӽû�ǰ �����������
	Public Function getshintvshoppingCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum
		Dim safetyCertYn, safetyCertNo, safetyConfirmYn, safetyConfirmNo, childSafetyCertYn, childSafetyCertNo, childSafetyConfirmYn, childSafetyConfirmNo
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�

		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, t.certNum "
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

		safetyCertYn			= "0"
		safetyCertNo			= ""
		safetyConfirmYn			= "0"
		safetyConfirmNo			= ""
		childSafetyCertYn		= "0"
		childSafetyCertNo		= ""
		childSafetyConfirmYn	= "0"
		childSafetyConfirmNo	= ""

		Select Case safetyDiv
			Case "10", "40"
				safetyCertYn			= "1"
				safetyCertNo			= certNum
			Case "20", "50"
				safetyConfirmYn			= "1"
				safetyConfirmNo			= certNum
			Case "70"
				childSafetyCertYn		= "1"
				childSafetyCertNo		= certNum
			Case "80"
				childSafetyConfirmYn	= "1"
				childSafetyConfirmNo	= certNum
		End Select

		strRst = strRst & "&safetyCertYn=" & safetyCertYn					'#������������ | �ش� ��ǰ�� �������� ����		
		strRst = strRst & "&safetyCertNo=" & safetyCertNo					'����������ȣ | �ش� ��ǰ�� �ο��� ����������ȣ		
		strRst = strRst & "&safetyConfirmYn=" & safetyConfirmYn				'#����Ȯ�ο��� | �ش� ��ǰ�� ����Ȯ�� ����		
		strRst = strRst & "&safetyConfirmNo=" & safetyConfirmNo				'����Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ����Ȯ�ι�ȣ		
		strRst = strRst & "&suppSuitYn=0"									'#���������ռ�Ȯ�ο��� | �ش� ��ǰ�� ���������ռ� Ȯ�ο���		
		strRst = strRst & "&suppSuitNo="									'���������ռ�Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ���������ռ�Ȯ�ι�ȣ		
		strRst = strRst & "&radioWaveCertYn=0"								'#������������ | �ش� ��ǰ�� �������� ����		
		strRst = strRst & "&radioWaveCertNo="								'����������ȣ | �ش� ��ǰ�� �ο��� ����������ȣ		
		strRst = strRst & "&childSafetyCertYn=" & childSafetyCertYn			'#��̾����������� | �ش� ��ǰ�� ��� Ư������ ���� �������� ����		
		strRst = strRst & "&childSafetyCertNo=" & childSafetyCertNo			'��̾���������ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ����������ȣ		
		strRst = strRst & "&childSafetyConfirmYn=" & childSafetyConfirmYn	'#��̾���Ȯ�ο��� | �ش� ��ǰ�� ��� Ư������ ���� ����Ȯ�� ����		
		strRst = strRst & "&childSafetyConfirmNo=" & childSafetyConfirmNo	'��̾���Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ����Ȯ�ι�ȣ		
		strRst = strRst & "&childSuppSuitYn=0"								'#��̰��������ռ�Ȯ�ο��� | �ش� ��ǰ�� ��� Ư������ ���� ���������ռ� Ȯ�ο���		
		strRst = strRst & "&childSuppSuitNo="								'��̰��������ռ�Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ���������ռ�Ȯ�ι�ȣ
		strRst = strRst & "&chemiSafetyConfirmYn=0"							'#��Ȱȭ����ǰ����Ȯ�ο���
		strRst = strRst & "&chemiSafetyConfirmNo="							'��Ȱȭ����ǰ����Ȯ�ι�ȣ
		getshintvshoppingCertParameter = strRst
	End Function

	'�ӽû�ǰ ���ο�û
	Public Function getshintvshoppingConfirmParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&scmGoodsCode=" & FShintvshoppingGoodNo		'#��ǰ�ڵ�
		getshintvshoppingConfirmParameter = strRst
	End Function

	'��ǰ �Ǹ��ߴ� ó��
	Public Function getShintvshoppingSellynParameter(ichgSellYn)
		Dim strRst
		'saleNoCode
		'https://wapi.10x10.co.kr/outmall/shintvshopping/shintvshoppingActProc.asp?act=commonCode&interfaceId=IF_API_00_021
		'101 : ��ü�ε�, 102 : ��ǰ���޺Ҿ���, 103 : ����ó������, 104 : ��� ǰ���̽� (������ ONLY), 105 : �����ߴ�, 106 : ǰ������, 201 : �ӽþ�ü��ǰ, 999 : �ŷ�����

		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
		If ichgSellYn = "Y" Then
			'��ü�� �� �Ǹŷ� �Ѵ�? Ȯ���غ�����..
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=00"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=" 							'#�Ұ����� �ڵ� | "�ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�"
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		ElseIf ichgSellYn = "N" Then
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=11"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=105" 						'#�Ұ����� �ڵ� | "�ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�"
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		ElseIf ichgSellYn = "X" Then
			strRst = strRst & "&goodsdtCode=000" 						'#��ǰ�ڵ� | �ڵ尪 000�� ��� ��ǰ ��ü ó��
			strRst = strRst & "&saleGb=19"								'#�Ǹű��� | 00:�Ǹ�����, 11:�Ͻ��ߴ�, 19:��������
			strRst = strRst & "&saleNoCode=105"							'#�Ұ����� �ڵ� | "�ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�"
			strRst = strRst & "&saleNoNote=�Ǹ�����" 					'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		End If
		getShintvshoppingSellynParameter = strRst
	End Function

	'�ǸŻ�ǰ ��ȸ(��)_v2
	Public Function getShintvshoppingItemViewParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
'		strRst = strRst & "&bDate="										'#��ȸ �������� | ����� ����  YYYYMMDDŸ��. ex) 20130118"		
'		strRst = strRst & "&eDate="										'#��ȸ ���������� | ����� ���� YYYYMMDDŸ��. ex) 20130118"		
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'��ǰ�ڵ� ������ȸ. �ڵ� ��ȸ�� ����� ���� ���� ����
		getShintvshoppingItemViewParameter = strRst
	End Function

	'�ǸŻ�ǰ �������� ����_v2
	Public Function getshintvshoppingItemEditParameter(iShipcostCode)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&shipManSeq=" & shipManSeq					'#������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&returnManSeq=" & returnManSeq				'#ȸ������� | ��ü�������ȸ ����(IF_API_00_017)
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ� | ��ǰ������ �ʼ�
		strRst = strRst & "&goodsName=" & getItemNameFormat()			'#��ǰ�� | . : , ; ( ! ? ) + - * / = [ ] Ư������ �Է� ����
		strRst = strRst & "&arsName=" & getItemNameFormat()				'#ARS�� | [default:��ǰ��]
'		strRst = strRst & "&mobileGoodsName=" & getItemNameFormat()		'#����ϻ�ǰ�� | [default:��ǰ��] => �ʼ���� �� �� �Ѱܵ� ��
'		strRst = strRst & "&slipNamePrintYn=0							'������ǰ�� ��¿��� | 0:N, 1:Y [default:0]
'		strRst = strRst & "&slipName=" & getItemNameFormat()			'������ǰ�� | ������ǰ�� ��¿��ΰ� 1�� ��� �Է°����ϸ�,  ������ǰ�� �ʼ� �Է�
'		strRst = strRst & "&weight="									'#���� | [default:0] => �ʼ���� �� �� �Ѱܵ� ��
'		strRst = strRst & "&vWidth="									'���� | [default:0] (����:cm)
'		strRst = strRst & "&vLength="									'���� | [default:0] (����:cm)
'		strRst = strRst & "&vHeight="									'���� | [default:0] (����:cm)
		strRst = strRst & "&installYn=0"								'��ġ��ۿ��� | 0:N, 1:Y, [default:0]
		strRst = strRst & "&codYn=0"									'���ҹ�ۿ��� | 0:N, 1:Y, [default:0] ��ġ��ۿ��ΰ� 1�� ��� ���ҹ�ۿ��� �Է� ����
		strRst = strRst & "&groupGoods="& Chkiif(IsMakeItem()="Y", "80", "")	'�׷��ǰ | 40: �ؿܱ��Ŵ���, 80: �ֹ����� '���� ������ �� ��� �ش�Ӽ� ��� �Ұ�
		strRst = strRst & "&shipCostCode=" & iShipcostCode				'#��ۺ���å�ڵ� | "��ۺ���å ��ȸ ����(IF_API_00_025) ���ҹ�ۿ��ΰ� 1�� ��� ��������å[A01 �Ǵ� A001]���� ����
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")		'#���λ�ǰ���� | 0:N, 1:Y
		strRst = strRst & "&orderMinQty=1"								'#�ֹ��ּҼ���
		strRst = strRst & "&orderMaxQty="&getOrderMaxNum				'#�ֹ��ִ����
		strRst = strRst & "&parallelImportYn=0"							'#������Կ��� | 0:N, 1:Y
		strRst = strRst & "&mixPackYn=1"								'�����尡�ɿ��� | 0:N, 1:Y [default:0]
		strRst = strRst & "&doNotIslandDelyYn=0"						'����/�갣 ��ۺҰ� ���� | 0: ��۰���, 1: ��� �Ұ� [default : 0]
		strRst = strRst & "&doNotJejuDelyYn=0"							'���� ��ۺҰ� ���� | 0: ��۰���, 1: ��� �Ұ� [default : 0]
		strRst = strRst & "&originCode=" & getOriginCode				'#������ | ��������ȸ ����(IF_API_00_018)
		strRst = strRst & "&oemEntpName="								'OEM��� | �������� �ѱ��� �ƴϰ� ������ü�Ը� �߼ұ���̸� �ʼ��Է�
'		strRst = strRst & "&avgDelyLeadtime=" & getShopLeadTime()		'��ۼҿ���
		strRst = strRst & "&avgDelyLeadtime=5"							'��ۼҿ���
		getshintvshoppingItemEditParameter = strRst
	End Function

	'���»� ���ݵ��
	Public Function getshintvshoppingEditPriceParameter()
		Dim strRst
		Dim saleVat, buyPrice, buyCost, buyVat
		buyPrice	= Clng(MustPrice()*0.88)
		buyVat		= REPLACE(Formatnumber(buyPrice / 11, 0), ",", "")
		buyCost		= buyPrice - buyVat
		saleVat		= REPLACE(Formatnumber(MustPrice / 11, 0), ",", "")

		If FVatInclude = "N" Then
			buyVat	= 0
			buyCost = buyPrice
			saleVat	= 0
		End If

		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&immediatelyApplyYn=1"						'������뿩�� | 0:N, 1:Y [default:0] 0 => ���������Ͻ� �������� ��ǰ���� ���, 1 => ���������Ͻÿ� ������� ������� ����. ���������Ͻ�(applyDate) �ʼ����� ( ����������� �������� ).��, ���������� ���� �������� ������� ��� �Ұ�
'		strRst = strRst & "&applyDate="
		strRst = strRst & "&buyPrice="& buyPrice						'#���԰�
		strRst = strRst & "&buyCost=" & buyCost							'#���Դܰ�(vat����)
		strRst = strRst & "&buyVat=" & buyVat							'#����vat
		strRst = strRst & "&salePrice=" & MustPrice						'#�ǸŰ�
		strRst = strRst & "&saleVat=" & saleVat							'#�Ǹ�vat
'		strRst = strRst & "&custPrice="									'�����ǸŰ� | �̻��
'		strRst = strRst & "&signGb="									'��û�ܰ� | �������� ���°� (00:�ӽ�����, 10:Ȯ�ο�û) (DEFAULT : 10)
		getshintvshoppingEditPriceParameter = strRst
	End Function

	'�ǸŻ�ǰ ����� ���
	Public Function getshintvshoppingEditContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#��ǰ�ڵ�
		strRst = strRst & "&descCode=998" 								'#����� �ڵ� | ����� ��ȸ ����(IF_API_00_016) | 101 : ��ǰ����, 301: ��۾ȳ�, 302 : ��ǰ/��ȯ�ȳ�, 303 : AS�ȳ�, 997 : ����ϱ����(QS), 998 : ����ϱ����
		strRst = strRst & "&descExt=" & getShintvshoppingContParamToReg()	'#����� ����
		getshintvshoppingEditContentParameter = strRst
	End Function

	'�ǸŻ�ǰ �̹��� ���(URL)
	Public Function getshintvshoppingEditImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
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
		getshintvshoppingEditImageParameter = strRst
	End Function

	'�ǸŻ�ǰ ����������� ���
	Public Function getshintvshoppingGosiEditParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","����")

		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&typeCode=" & mallinfodiv					'#Ÿ���ڵ� | ��ǰ����������� ��ǰ���� ��ȸ ����(IF_API_00_022)		
		strRst = strRst & "&offerCode=" & mallinfocd					'#�׸��ڵ� | ��ǰ����������� ǰ�� �׸� ����(IF_API_00_023)		
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)				'#�׸񳻿�
		getshintvshoppingGosiEditParameter = strRst
	End Function

	'�ǸŻ�ǰ �����������
	Public Function getshintvshoppingEditCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum
		Dim safetyCertYn, safetyCertNo, safetyConfirmYn, safetyConfirmNo, childSafetyCertYn, childSafetyCertNo, childSafetyConfirmYn, childSafetyConfirmNo
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#���������� ����� �ǸŻ�ǰ�� ��ǰ�ڵ�

		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, t.certNum "
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

		safetyCertYn			= "0"
		safetyCertNo			= ""
		safetyConfirmYn			= "0"
		safetyConfirmNo			= ""
		childSafetyCertYn		= "0"
		childSafetyCertNo		= ""
		childSafetyConfirmYn	= "0"
		childSafetyConfirmNo	= ""

		Select Case safetyDiv
			Case "10", "40"
				safetyCertYn			= "1"
				safetyCertNo			= certNum
			Case "20", "50"
				safetyConfirmYn			= "1"
				safetyConfirmNo			= certNum
			Case "70"
				childSafetyCertYn		= "1"
				childSafetyCertNo		= certNum
			Case "80"
				childSafetyConfirmYn	= "1"
				childSafetyConfirmNo	= certNum
		End Select

		strRst = strRst & "&safetyCertYn=" & safetyCertYn					'#������������ | �ش� ��ǰ�� �������� ����		
		strRst = strRst & "&safetyCertNo=" & safetyCertNo					'����������ȣ | �ش� ��ǰ�� �ο��� ����������ȣ		
		strRst = strRst & "&safetyConfirmYn=" & safetyConfirmYn				'#����Ȯ�ο��� | �ش� ��ǰ�� ����Ȯ�� ����		
		strRst = strRst & "&safetyConfirmNo=" & safetyConfirmNo				'����Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ����Ȯ�ι�ȣ		
		strRst = strRst & "&suppSuitYn=0"									'#���������ռ�Ȯ�ο��� | �ش� ��ǰ�� ���������ռ� Ȯ�ο���		
		strRst = strRst & "&suppSuitNo="									'���������ռ�Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ���������ռ�Ȯ�ι�ȣ		
		strRst = strRst & "&radioWaveCertYn=0"								'#������������ | �ش� ��ǰ�� �������� ����		
		strRst = strRst & "&radioWaveCertNo="								'����������ȣ | �ش� ��ǰ�� �ο��� ����������ȣ		
		strRst = strRst & "&childSafetyCertYn=" & childSafetyCertYn			'#��̾����������� | �ش� ��ǰ�� ��� Ư������ ���� �������� ����		
		strRst = strRst & "&childSafetyCertNo=" & childSafetyCertNo			'��̾���������ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ����������ȣ		
		strRst = strRst & "&childSafetyConfirmYn=" & childSafetyConfirmYn	'#��̾���Ȯ�ο��� | �ش� ��ǰ�� ��� Ư������ ���� ����Ȯ�� ����		
		strRst = strRst & "&childSafetyConfirmNo=" & childSafetyConfirmNo	'��̾���Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ����Ȯ�ι�ȣ		
		strRst = strRst & "&childSuppSuitYn=0"								'#��̰��������ռ�Ȯ�ο��� | �ش� ��ǰ�� ��� Ư������ ���� ���������ռ� Ȯ�ο���		
		strRst = strRst & "&childSuppSuitNo="								'��̰��������ռ�Ȯ�ι�ȣ | �ش� ��ǰ�� �ο��� ��� Ư������ ���� ���������ռ�Ȯ�ι�ȣ		
		strRst = strRst & "&chemiSafetyConfirmYn=0"							'#��Ȱȭ����ǰ����Ȯ�ο���
		strRst = strRst & "&chemiSafetyConfirmNo="							'��Ȱȭ����ǰ����Ȯ�ι�ȣ
		getshintvshoppingEditCertParameter = strRst
	End Function

	'�ǸŻ�ǰ �����
	Public Function geshintvshoppingOptionQtyParam(outmallOptCode, optsu)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#�ǸŴ�ǰ�ڵ�
		strRst = strRst & "&inplanQty=" & optsu							'#�ǸŰ��ɼ���
		geshintvshoppingOptionQtyParam = strRst
	End Function

	'��ǰ �Ǹ��ߴ� ó��
	Public Function geshintvshoppingOptionStatParam(outmallOptCode, isalegb)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#�ǸŴ�ǰ�ڵ�
		strRst = strRst & "&saleGb=" & isalegb							'#�Ǹű���
		If isalegb = "11" Then
			strRst = strRst & "&saleNoCode=105" 						'#�Ұ����� �ڵ� | "�ǸźҰ����� ��ȸ(API_0016) ����, ����(����/����) ó���� �ʼ�"
			strRst = strRst & "&saleNoNote=" 							'�Ұ� �ڸ�Ʈ | �������� ó���� ���� �ڸ�Ʈ ���
		End If
		geshintvshoppingOptionStatParam = strRst
	End Function

	'�ǸŻ�ǰ ��ǰ���� ���_v2
	Public Function geshintvshoppingOptionAddParam(otherText, maxSaleQty)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & linkCode						'#�����ڵ� | ����Ŀ: SLINK [ TCODE.LGROUP : xxx ]
		strRst = strRst & "&entpCode=" & entpCode						'#��ü�ڵ� | �ż���TV���ο��� �ο��� ��ü�ڵ� 6�ڸ�
		strRst = strRst & "&entpId=" & entpId							'#��ü�����ID | �ż���TV���ο��� �ο��� ��ü����� ID
		strRst = strRst & "&entpPass=" & entpPass						'#��üPASSWORD | �ż���TV���ο��� ����� ��ü����� ��й�ȣ
		strRst = strRst & "&goodsCode=" & FShintvshoppingGoodNo			'#�ǸŻ�ǰ�ڵ�
'		strRst = strRst & "&optionCode1="								'�ɼ�1�ڵ� | �ɼǱ׷�1�� �ش��ϴ� �ڵ� �Է�
'		strRst = strRst & "&optionCode2="								'�ɼ�2�ڵ� | �ɼǱ׷�2�� ���� ������ �ʼ� �Է�, �ɼǱ׷�2�� �ش��ϴ� �ڵ� �Է�
'		strRst = strRst & "&optionCode3="								'�ɼ�3�ڵ� | �ɼǱ׷�3�� ���� ������ �ʼ� �Է�, �ɼǱ׷�3�� �ش��ϴ� �ڵ� �Է�
'		strRst = strRst & "&optionCode4="								'�ɼ�4�ڵ� | �ɼǱ׷�4�� ���� ������ �ʼ� �Է�, �ɼǱ׷�4�� �ش��ϴ� �ڵ� �Է�
		strRst = strRst & "&dtText=" & URLEncodeUTF8Plus(otherText)		'�ؽ�Ʈ�Է� | �ڵ��Է� �Ǵ� �ؽ�Ʈ�Է��� �� 1
		strRst = strRst & "&maxSaleQty=" & maxSaleQty					'�ִ��Ǹż��� | ���ڸ� �Է°���
'		strRst = strRst & "&modelName="
		geshintvshoppingOptionAddParam = strRst
	End Function
End Class

Class CShintvshopping
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

	Public Sub getShintvshoppingNotRegOneItem
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
            addSql = addSql & " or optAddCNT>0"
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.shintvshoppingStatCD,-9) as shintvshoppingStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		strSql = strSql & "	, isnull(am.lgroup, '') as lgroup "
		strSql = strSql & "	, isnull(am.mgroup, '') as mgroup "
		strSql = strSql & "	, isnull(am.sgroup, '') as sgroup "
		strSql = strSql & "	, isnull(am.dgroup, '') as dgroup "
		strSql = strSql & "	, isnull(am.tgroup, '') as tgroup "
		strSql = strSql & "	, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_shintvshopping_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_category as tm on am.lgroup = tm.lgroup and am.mgroup = tm.mgroup and am.sgroup = tm.sgroup and am.dgroup = tm.dgroup and am.tgroup = tm.tgroup "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as R on i.itemid = R.itemid"
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
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_shintvshopping_regItem WHERE shintvshoppingStatCD >= 3) "	''��ϿϷ��̻��� ��Ͼȵ�.	'shintvshopping��ϻ�ǰ ����
		strSql = strSql & " and cm.mapCnt is Not Null "'	ī�װ� ��Ī ��ǰ��
		strSql = strSql & addSql																				'ī�װ� ��Ī ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CShintvshoppingItem
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
                FOneItem.FShintvshoppingStatCD		= rsget("shintvshoppingStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FLgroup			= rsget("lgroup")
				FOneItem.FMgroup			= rsget("mgroup")
				FOneItem.FSgroup			= rsget("sgroup")
				FOneItem.FDgroup			= rsget("dgroup")
				FOneItem.FTgroup			= rsget("tgroup")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOutmallstandardMargin = rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub getShintvshoppingTmpRegedOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.itemid, r.shintvshoppingGoodNo, i.smallImage, i.basicImage, i.mainimage, i.mainimage2, c.itemcontent "
		strSql = strSql & " ,ordercomment, isNull(r.reglevel, 0) as reglevel, i.limityn, i.limitno, i.limitsold "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_shintvshopping_regItem as r "
		strSql = strSql & " JOIN db_item.dbo.tbl_item as i on r.itemid = i.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " WHERE 1=1 "
		strSql = strSql & " and r.itemid = '"& FRectItemID &"' "
		strSql = strSql & " and isNull(shintvshoppingGoodNo, '') <> '' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CShintvshoppingItem
				FOneItem.FItemid					= rsget("itemid")
				FOneItem.FShintvshoppingGoodNo		= rsget("shintvshoppingGoodNo")
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
		End If
		rsget.Close
	End Sub

	Public Sub getShintvshoppingEditOneItem
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
		strSql = strSql & "	, m.shintvshoppingGoodNo, m.shintvshoppingprice, m.shintvshoppingSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & "	, isnull(am.lgroup, '') as lgroup "
		strSql = strSql & "	, isnull(am.mgroup, '') as mgroup "
		strSql = strSql & "	, isnull(am.sgroup, '') as sgroup "
		strSql = strSql & "	, isnull(am.dgroup, '') as dgroup "
		strSql = strSql & "	, isnull(am.tgroup, '') as tgroup "
		strSql = strSql & "	, isNULL(m.shintvshoppingStatCD,-9) as shintvshoppingStatCD, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
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
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_shintvshopping_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_shintvshopping_category as tm on am.lgroup = tm.lgroup and am.mgroup = tm.mgroup and am.sgroup = tm.sgroup and am.dgroup = tm.dgroup and am.tgroup = tm.tgroup  "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.shintvshoppingGoodNo is Not Null "									'#��� ��ǰ��
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CShintvshoppingItem
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
				FOneItem.FShintvshoppingGoodNo	= rsget("shintvshoppingGoodNo")
				FOneItem.FShintvshoppingprice	= rsget("shintvshoppingprice")
				FOneItem.FShintvshoppingSellYn	= rsget("shintvshoppingSellYn")

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
				FOneItem.Fvatinclude        = rsget("vatinclude")
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
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Shintvshopping_ItemOptionMapping_Get] '"&iitemid&"' "
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
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Shintvshopping_InfoCodeMap_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getInfoCodeMapList = rsget.getRows
	end if
	rsget.Close
End Function

Function getOptiopnMapList(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Shintvshopping_OptionMappingByEdit_Get] '"&iitemid&"' "
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
	strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Shintvshopping_OptionMappingByAdd_Get] '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getOptiopnMayAddList = rsget.getRows
	end if
	rsget.Close
End Function

Function getShintvshoppingOptCnt(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM db_item.dbo.tbl_OutMall_regedoption "
	strSql = strSql & " WHERE mallid = '"& CMALLNAME &"' "
	strSql = strSql & " and itemid = '"& iitemid &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.EOF) then
		getShintvshoppingOptCnt = rsget("cnt")
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
%>