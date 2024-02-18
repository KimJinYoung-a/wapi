<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "skstoa"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

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
		strRst = "텐바이텐," & FKeywords
		strRst = replace(strRst, "인기", "")
		strRst = replace(strRst, "인치", "")
		strRst = replace(strRst, "모기퇴치", "")
		strRst = RightCommaDel(strRst)
		If strRst = "" Then
			strRst = "텐바이텐"
		End If

		If stringCount(strRst) > 100 Then
			strRst = chrbyte(strRst, 100, "N")
		End If
		getKeywords = URLEncodeUTF8Plus(strRst)
	End Function

	Function stringCount(strString)
		Dim intPos, chrTemp, intLength
		'문자열 길이 초기화
		intLength = 0
		intPos = 1

		'문자열 길이만큼 돈다
		while ( intPos <= Len( strString ) )
			'문자열을 한문자씩 비교한다
			chrTemp = ASC(Mid( strString, intPos, 1))
			if chrTemp < 0 then '음수값(-)이 나오면 한글임
				intLength = intLength + 2 '한글일 경우 2바이트를 더한다
			else
				intLength = intLength + 1 '한글이 아닐경우 1바이트를 더한다
			end If
			intPos = intPos + 1
		wend
		stringCount = intLength
	End function

	'// 품절여부
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
		sqlStr = sqlStr & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : 수입
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

	'// Skstoa 판매여부 반환
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

	'// Skstoa 정보제공타입
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
			strRst = strRst & "- 주문시 유의사항 :<br>" & URLEncodeUTF8(Fordercomment) & "<br>"
		End If

		'#기본 상품설명
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

		'# 추가 상품 설명이미지 접수
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

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage & """ border=""0""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & Server.URLEncode("<img src=""" & FmainImage2 & """ border=""0""><br>")

		'#배송 주의사항
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
		'//결과 반환
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

				If i = 0 Then		'0원 옵션의 재고가 5개 이하면 품절
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
			'// 이중옵션확인
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
				'// 이중옵션 일때
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
				'// 단일옵션일 때
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
		'//결과 반환
		checkTenItemOptionValid = chkRst
	End Function
 
    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST상품] "&FItemName
		Else
			'buf = "[텐바이텐] "&FItemName
			buf = FItemName		'2022-02-07 변장혁님 요청 / 상품명앞 텐바이텐 삭제
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
		buf = replace(buf,"_","/")
        buf = replace(buf,"%","프로")
		buf = replace(buf,"&","/")
        buf = replace(buf,"&amp;","")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
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
			getMakecoCode = "9999"	'기타
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
			getOriginCode = "9999"	'기타
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
			getBrandCode = "999999"	'기타
		End If
		rsget.Close
	End Function

	Public Function IsFreeBeasong()
		IsFreeBeasong = False
		If (FdeliveryType=2) or (FdeliveryType=4) or (FdeliveryType=5) then				'2(텐무), 4,5(업무)
			IsFreeBeasong = True
		End If
'		If (FSellcash>=30000) then IsFreeBeasong=True
		If (FdeliveryType=9) Then														'업체조건
'			If (Clng(FSellcash) >= Clng(FdefaultfreeBeasongLimit)) then
'				IsFreeBeasong=True
'			End If
			IsFreeBeasong = False
		End If
    End Function

	'임시상품 기초정보 등록							
	Public Function getskstoaItemRegParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId							'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsName=" & getItemNameFormat()				'#상품명
		strRst = strRst & "&keyword=" & getKeywords							'검색어 | [default:상품명]
'		strRst = strRst & "&arsName=" & getItemNameFormat()					'ARS명 | [default:상품명]
'		strRst = strRst & "&mobileGoodsName="								'모바일상품명 | [default:상품명]
		strRst = strRst & "&entpManSeq=" & skentpManSeq						'#업체담당자 | 업체담당자조회 참조(IF_API_00_017)
		strRst = strRst & "&mdCode=" & skmdCode								'#MD | MD리스트 참조(IF_API_00_001), 소싱매체가 쇼핑몰인 MD중 택1
		strRst = strRst & "&taxYn=" & CHKIIF(FVatInclude="N","0","1")		'판매과세여부 | 0:면세, 1:과세
		strRst = strRst & "&codeLgroup=" & FLgroup							'#대분류 | 상품 대분류 조회 참조(IF_API_00_002)
		strRst = strRst & "&codeMgroup=" & FMgroup							'#중분류 | 상품 중분류 조회 참조(IF_API_00_003)
		strRst = strRst & "&codeSgroup=" & FSgroup							'#소분류 | 상품 소분류 조회 참조(IF_API_00_004)
		strRst = strRst & "&codeDgroup=" & FDgroup							'#세분류 | 상품 세분류 조회 참조(IF_API_00_005)
		strRst = strRst & "&shipCostCode=" & skshipCostCode					'#고객부담배송비구분 | 배송비정책 조회 참조(IF_API_00_025)
		strRst = strRst & "&delyBoxQty=1"									'#배송Box수량 | [default:1]
		strRst = strRst & "&installYn=0"									'#설치배송여부 | 0:N, 1:Y, [default:0]
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")	'#성인상품여부 | 0:N, 1:Y
		strRst = strRst & "&makecoCode=" & getMakecoCode					'#제조업체 | 제조업체조회 참조(IF_API_00_019)
		strRst = strRst & "&originCode=" & getOriginCode					'#원산지 | 원산지조회 참조(IF_API_00_018)
		strRst = strRst & "&brandCode="	& getBrandCode						'#브랜드 | 브랜드조회 참조(IF_API_00_015)		
		strRst = strRst & "&buyPrice=" & MustBuyPrice(MustPrice())			'#매입가 | 소수점 불가
		strRst = strRst & "&salePrice=" & MustPrice()						'#판매가
		strRst = strRst & "&custPrice=" & MustPrice()						'#시중판매가 ???
		strRst = strRst & "&shipManSeq=" & skshipManSeq						'#출고담당자 | 업체담당자조회 참조(IF_API_00_017)
		strRst = strRst & "&returnManSeq=" & skreturnManSeq					'#회수담당자 | 업체담당자조회 참조(IF_API_00_017)
		strRst = strRst & "&offerType=" & getSkstoaOfferType				'#정보제공타입 | 상품정보제공고시 상품유형 조회 참조(IF_API_00_022)
'		strRst = strRst & "&weight="										'무게 | [default:0]
'		strRst = strRst & "&volumeDesc="									'부피설명
		strRst = strRst & "&formCode=F999"									'#단품형태 | 단품형태그룹 조회 참조(IF_API_00_008), F999(기타)
		strRst = strRst & "&sizeCode=S999"									'#크기 | 단품크기그룹 조회 참조(IF_API_00_007), S999(기타)
		strRst = strRst & "&costTaxYn=" & CHKIIF(FVatInclude="N","0","1")	'매입과세여부 | "미사용 컬럼으로 전환 - 판매과세만 입력받아서 판매과세와 동일한 데이터를 자동으로 세팅함.
		strRst = strRst & "&accountManSeq=" & skaccountManSeq				'#회계담당자 | 업체담당자조회 참조(IF_API_00_017)
		strRst = strRst & "&islandsDelyOut=0"								'#도서산간배송불가 | 0: 배송가능, 1:배송불가
		strRst = strRst & "&jejuDelyOut=0"									'#제주지역배송불가 | 0: 배송가능, 1:배송불가
		strRst = strRst & "&mdKind=" & skmdKind								'#MD분류코드 | MD분류리스트 참조(IF_API_00_026)
		strRst = strRst & "&collectYn=0"									'#착불여부 | 0:N, 1:Y, [default:0]
		strRst = strRst & "&returnNoYn=0"									'#교환/반품불가여부 | 0: 교환/반품가능, 1: 교환/반품불가
		strRst = strRst & "&globalDelyYn=0"									'#해외직배송여부 | 0: N, 1: Y
'		strRst = strRst & "&globalCode="									'이전되는국가코드 | 원산지조회와 동일(IF_API_00_018) 해외직배송일경우 필수 입력
'		strRst = strRst & "&modelNo="										'모델번호 | 모델명api를 통하여 검색 한 모델의 modelId
'		strRst = strRst & "&modelName="										'모델명 | 모델명api를 통하여 검색 한 모델의 modelName
'		strRst = strRst & "&dawnYn="										'새벽배송상품여부 | 0: 일반, 1: 동원새벽배송, 2: 기타새벽배송
		strRst = strRst & "&orderCreateYn=" & Chkiif(IsMakeItem()="Y", "1", "0")	'주문제작여부 | 0:N, 1:Y, [default:0]
'		strRst = strRst & "&bookYn="										'도서여부 | 0: N, 1: Y, 1인 경우 매입과세,판매과세는 0으로 연동 필수, ISBN 값 연동 필수
'		strRst = strRst & "&isbn="											'도서ISBN | bookYn이 1인 경우 필수
		getskstoaItemRegParameter = strRst
'		response.end
	End Function

	'임시상품 기술서 등록
	Public Function getskstoaContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId							'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#임시상품코드
		strRst = strRst & "&descCode=998" 									'#기술서 코드 | 기술서항목 조회 참조, 200 : 상품구성, 201 : 배송정보, 202 : 취소/반품/AS안내, 301 : TV상품상세 설명1, 302 : TV상품상세 설명2, 998 : 모바일설명서, 999 : WEB설명서
		strRst = strRst & "&descExt=" & getSkstoaContParamToReg()			'#기술서 내용 | CLOB type
		getskstoaContentParameter = strRst
	End Function

	'임시상품 단품정보 등록
	Public Function getskstoaOptParameter(otherText, maxSaleQty)
		Dim strRst, strSql, optcnt, limitsu
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId							'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#임시상품코드
		strRst = strRst & "&colorGroupCode=C000"							'#색상그룹코드 | 단품색상그룹조회 참조(IF_API_00_006) C000 : 전체색상
		strRst = strRst & "&patternGroupCode=P999"							'#무늬그룹코드 | 단품무늬그룹조회 참조(IF_API_00_009) P999 : 기타
'		strRst = strRst & "&colorCode="										'색상코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&patternCode="									'무늬코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&sizeCode="										'크기코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&formCode="										'형태코드 | 코드입력 또는 텍스트입력중 택 1
		strRst = strRst & "&otherText=" & URLEncodeUTF8Plus(otherText)		'단품기타 | 코드입력 또는 텍스트입력중 택 1, 단품명 허용정책 : 한글,영문,숫자, [~] [!] [*] [-][_][=][+][|][:][.][/][?], \\^\\(\\)\\[\\]\\{\\}]
		strRst = strRst & "&maxSaleQty=" & maxSaleQty						'#최대판매수량 | 숫자만 입력가능		
		getskstoaOptParameter = strRst
	End Function

	'임시상품 이미지 등록(URL)
	Public Function getskstoaImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode						'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode						'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId							'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass						'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo				'#임시상품코드
		strRst = strRst & "&imgUrlBase=" & FbasicImage 						'#메인이미지 URL
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

	'임시상품 정보제공고시 등록
	Public Function getskstoaGosiRegParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","프로")

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId								'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#임시상품코드
		strRst = strRst & "&typeCode=" & mallinfodiv							'#상품유형코드 | 상품정보제공고시 상품유형 조회 참조(IF_API_00_022)
		strRst = strRst & "&offerCode=" & mallinfocd							'#항목코드 | 상품정보제공고시 품목 항목 참조(IF_API_00_023)
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)	'#기술서내역 | CLOB type
		getskstoaGosiRegParameter = strRst
	End Function

	'임시상품 인증정보등록
	Public Function getskstoaCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum, notSendinfoDiv
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId								'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#임시상품코드
		strRst = strRst & "&typeCode=" & getSkstoaOfferType						'#상품유형코드 | 상품정보제공고시 상품유형 조회 참조(IF_API_00_022)
' - 아래 항목의 경우 인증번호를 받지 않음
' 16.의료기기
' 18.화장품
' 20.식품(농수산물)
' 21.가공식품
' 22.건강기능식품
' 27.호텔/펜션 예약
' 28.여행패키지
' 29.항공권
' 30.자동차 대여 서비스
' 33.디지털 콘텐츠(음원/게임/인터넷강의 등)
' 34.상품권/쿠폰
' 35.모바일쿠폰
' 36.영화/공연
' 37.기타용역

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

	'임시상품 승인요청
	Public Function getskstoaConfirmParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode							'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode							'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId								'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass							'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&scmGoodsCode=" & FSkstoaTmpGoodNo					'#임시상품코드
		getskstoaConfirmParameter = strRst
	End Function

	'상품 판매중단 처리
	Public Function getSkstoaSellynParameter(ichgSellYn)
		Dim strRst
		'saleNoCode
		'https://wapi.10x10.co.kr/outmall/skstoa/skstoaActProc.asp?act=commonCode&interfaceId=IF_API_00_021
		'101 : 업체부도, 102 : 상품수급불안정, 103 : 사후처리미흡, 104 : 품질보완, 105 : 구매중단, 999 : 거래종료

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		If ichgSellYn = "Y" Then
			'전체를 다 판매로 한다? 확인해봐야함..
			strRst = strRst & "&goodsdtCode=000" 						'#단품코드 | 코드값 000일 경우 단품 전체 처리
			strRst = strRst & "&saleGb=00"								'#판매구분 | 00:판매진행, 11:일시중단, 19:영구중지
			strRst = strRst & "&saleNoCode=" 							'#판매불가코드 | 판매불가사유 조회(API_0016) 참조, 중지(영시/영구) 처리시 필수
			strRst = strRst & "&saleNoNote=" 							'불가 코멘트 | 영구중지 처리시 사유 코멘트 등록
		ElseIf ichgSellYn = "N" Then
			strRst = strRst & "&goodsdtCode=000" 						'#단품코드 | 코드값 000일 경우 단품 전체 처리
			strRst = strRst & "&saleGb=11"								'#판매구분 | 00:판매진행, 11:일시중단, 19:영구중지
			strRst = strRst & "&saleNoCode=105" 						'#판매불가코드 | 판매불가사유 조회(API_0016) 참조, 중지(영시/영구) 처리시 필수
			strRst = strRst & "&saleNoNote=" 							'불가 코멘트 | 영구중지 처리시 사유 코멘트 등록
		ElseIf ichgSellYn = "X" Then
			strRst = strRst & "&goodsdtCode=000" 						'#단품코드 | 코드값 000일 경우 단품 전체 처리
			strRst = strRst & "&saleGb=19"								'#판매구분 | 00:판매진행, 11:일시중단, 19:영구중지
			strRst = strRst & "&saleNoCode=105"							'#판매불가코드 | 판매불가사유 조회(API_0016) 참조, 중지(영시/영구) 처리시 필수
			strRst = strRst & "&saleNoNote=판매종료"					'불가 코멘트 | 영구중지 처리시 사유 코멘트 등록
		End If
		getSkstoaSellynParameter = strRst
	End Function

	'판매상품 조회(상세)
	Public Function getSkstoaItemViewParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
'		strRst = strRst & "&bDate="										'#조회 시작일자 | 등록일 기준  YYYYMMDD타입. ex) 20130118"		
'		strRst = strRst & "&eDate="										'#조회 마지막일자 | 등록일 기준 YYYYMMDD타입. ex) 20130118"		
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'상품코드 개별조회. 코드 조회시 등록일 기준 조건 제외
		getSkstoaItemViewParameter = strRst
	End Function

	'판매상품 기초정보 수정
	Public Function getskstoaItemEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드 | 상품수정시 필수
		strRst = strRst & "&goodsName=" & getItemNameFormat()			'#상품명
		strRst = strRst & "&keyword=" & getKeywords						'#검색어 | [default:상품명]
		strRst = strRst & "&arsName=" & getItemNameFormat()				'#ARS명 | [default:상품명]
'		strRst = strRst & "&mobileGoodsName="							'#모바일상품명 | [default:상품명
'		strRst = strRst & "&weight="									'#무게 | [default:0]
'		strRst = strRst & "&volumeDesc="								'#부피설명
		strRst = strRst & "&shipCostCode=" & skshipCostCode				'#고객부담배송비구분코드 | 배송비정책조회 참조(IF_API_00_025)
		strRst = strRst & "&adultYn=" & Chkiif(IsAdultItem()="Y", "1", "0")		'#성인상품여부 | 0:N, 1:Y
		strRst = strRst & "&orderMinQty=1"								'#주문최소수량
		strRst = strRst & "&orderMaxQty=" & getOrderMaxNum				'#주문최대수량
		strRst = strRst & "&collectYn=0"								'#착불여부 | 0:N, 1:Y, [default:0]
		strRst = strRst & "&returnNoYn=0"								'#교환/반품불가여부 | 0: 교환/반품가능, 1: 교환/반품불가
'		strRst = strRst & "&modelNo="									'모델번호
'		strRst = strRst & "&modelName="									'모델명
		getskstoaItemEditParameter = strRst
	End Function

	'협력사 가격등록
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
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&applyDate="	& Replace(DATE(), "-", "/")		'#적용일자 | YYYY/MM/DD타입. ex) 2018/09/06
		strRst = strRst & "&salePrice=" & MustPrice						'#판매가
		strRst = strRst & "&buyPrice="& buyPrice						'매입가
		getskstoaEditPriceParameter = strRst
	End Function

	'판매상품 기술서 등록
	Public Function getskstoaEditContentParameter
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&descCode=998" 								'#기술서코드 | 기술서항목 조회 참조, 200 : 상품구성, 201 : 배송정보, 202 : 취소/반품/AS안내, 301 : TV상품상세 설명1, 302 : TV상품상세 설명2, 998 : 모바일설명서, 999 : WEB설명서
		strRst = strRst & "&descExt=" & getSkstoaContParamToReg()		'#기술서내역 | CLOB type
		getskstoaEditContentParameter = strRst
	End Function

	'판매상품 이미지 등록(URL)
	Public Function getskstoaEditImageParameter
		Dim strRst, strSQL, imgurlparam
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&imgUrlBase=" & FbasicImage 					'메인이미지 URL
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

	'판매상품 정보제공고시 등록
	Public Function getskstoaGosiEditParameter(mallinfocd, mallinfodiv, infocontent)
		Dim strRst
		infocontent = replace(infocontent,"%","프로")

		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&typeCode=" & mallinfodiv					'#타입코드 | 상품정보제공고시 상품유형 조회 참조(IF_API_00_022)		
		strRst = strRst & "&offerCode=" & mallinfocd					'#항목코드 | 상품정보제공고시 품목 항목 참조(IF_API_00_023)		
		strRst = strRst & "&offerContents=" & URLEncodeUTF8Plus(infocontent)				'#항목내용 | CLOB type
		getskstoaGosiEditParameter = strRst
	End Function

	'판매상품 인증정보등록
	Public Function getskstoaEditCertParameter()
		Dim strRst, strSql, isRegCert, safetyDiv, certNum, notSendinfoDiv
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&typeCode=" & getSkstoaOfferType				'#상품유형코드 | 상품정보제공고시 상품유형 조회 참조(IF_API_00_022)
' - 아래 항목의 경우 인증번호를 받지 않음
' 16.의료기기
' 18.화장품
' 20.식품(농수산물)
' 21.가공식품
' 22.건강기능식품
' 27.호텔/펜션 예약
' 28.여행패키지
' 29.항공권
' 30.자동차 대여 서비스
' 33.디지털 콘텐츠(음원/게임/인터넷강의 등)
' 34.상품권/쿠폰
' 35.모바일쿠폰
' 36.영화/공연
' 37.기타용역
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

	'판매상품 재고변경
	Public Function geskstoaOptionQtyParam(outmallOptCode, optsu)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#판매단품코드
		strRst = strRst & "&inplanQty=" & optsu							'#판매가능수량
		geskstoaOptionQtyParam = strRst
	End Function

	'상품 판매중단 처리
	Public Function geskstoaOptionStatParam(outmallOptCode, isalegb)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
		strRst = strRst & "&goodsdtCode=" & outmallOptCode				'#판매단품코드
		strRst = strRst & "&saleGb=" & isalegb							'#판매구분
		If isalegb = "11" Then
			strRst = strRst & "&saleNoCode=105" 						'#불가사유 코드 | "판매불가사유 조회(API_0016) 참조, 중지(영시/영구) 처리시 필수"
			strRst = strRst & "&saleNoNote=" 							'불가 코멘트 | 영구중지 처리시 사유 코멘트 등록
		End If
		geskstoaOptionStatParam = strRst
	End Function

	'판매상품 단품정보 등록
	Public Function getskstoaOptionAddParam(otherText, maxSaleQty)
		Dim strRst
		strRst = ""
		strRst = strRst & "linkCode=" & skstoalinkCode					'#연결코드 | SKB에서 부여한 연결코드
		strRst = strRst & "&entpCode=" & skstoaentpCode					'#업체코드 | SKB에서 부여한 업체코드 6자리
		strRst = strRst & "&entpId=" & skstoaentpId						'#업체사용자ID | SKB에서 부여한 업체사용자 ID
		strRst = strRst & "&entpPass=" & skstoaentpPass					'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
		strRst = strRst & "&goodsCode=" & FSkstoaGoodNo					'#판매상품코드
'		strRst = strRst & "&sizeCode="									'크기코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&colorCode="									'색상코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&patternCode="								'무늬코드 | 코드입력 또는 텍스트입력중 택 1
'		strRst = strRst & "&formCode="									'형태코드 | 코드입력 또는 텍스트입력중 택 1
		strRst = strRst & "&dtText=" & URLEncodeUTF8Plus(otherText)		'텍스트입력 | 코드입력 또는 텍스트입력중 택 1
		strRst = strRst & "&maxSaleQty=" & maxSaleQty					'최대판매수량 | 숫자만 입력가능
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
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
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
		strSql = strSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'택배(일반)
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv in ('01', '16', '07') "		'01 : 일반, 16 : 주문제작, 07 : 구매제한 / 신세계tv홈쇼핑은 주문제작 문구(06) 불가!
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_skstoa_regItem WHERE skstoaStatCD >= 3) "	''등록완료이상은 등록안됨.	'skstoa등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "'	카테고리 매칭 상품만
		strSql = strSql & addSql																				'카테고리 매칭 상품만
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
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

        ''//연동 제외상품
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
		strSql = strSql & " 	or i.itemdiv not in ('01', '16', '07') "		'01 : 일반, 16 : 주문제작, 07 : 구매제한
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "		or isnull(am.lgroup, '') = '' "		'카테고리 미매핑
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
		strSql = strSql & " and m.skstoaGoodNo is Not Null "									'#등록 상품만
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

'// 상품이미지 존재여부 검사
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
		skmdCode = "0263"		'최희정
	Else
		skmdCode = "0263"		'최희정
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
		skmdKind = "0005"			'생활용품
	Else
		skmdKind = "0005"			'생활용품
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