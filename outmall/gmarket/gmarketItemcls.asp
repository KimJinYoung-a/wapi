<%
CONST CMAXMARGIN = 18
CONST CMALLNAME = "gmarket1010"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST gmarketAPIURL = "http://tpl.gmarket.co.kr"
CONST gmarketSSLAPIURL = "https://tpl.gmarket.co.kr"
CONST gmarketTicket = "0A2799EE6A1B65CC78DA96AA52C7546B2181855E48A0A31EDD4F3A77C3C61015856FE3DE5D7828B129A31AAD5914D7060556616D3AB7F2A84008A600C89F5953A0362065429900D0EB25CEBEA0E1CAF9E784FBC4F36E86608F2CF44B40113ADF"
CONST CDEFALUT_STOCK = 999
CONST CRETURNFEE = 3000
CONST MAKERNO = "100005224"	'기타 // 추후 수정 가능성 있음

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

	'// 품절여부
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
		sqlStr = sqlStr & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : 수입
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

	'// 지마켓 판매여부 반환
	Public Function getGmarketSellYn()
		'판매상태 (10:판매진행, 20:품절)
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
			FItemName = "TEST상품 "&FItemName
		End If

		If Date() >="2017-03-10" and Date() <= "2017-03-12" Then
			Select Case FItemid
				Case "1625309"		FItemName = FItemName & " 미키 커피드리퍼"
				Case "1569915"		FItemName = FItemName & " 미키 팝콘메이커"
				Case "1565223"		FItemName = FItemName & " 앨리스 접시"
				Case "1523844"		FItemName = FItemName & " 정글북 파우치"
				Case "1523843"		FItemName = FItemName & " 정글북 파우치"
				Case "1523841"		FItemName = FItemName & " 정글북 여행용 파우치"
				Case "1523840"		FItemName = FItemName & " 정글북 멀티 파우치"
				Case "1523839"		FItemName = FItemName & " 정글북 메이크업 파우치"
				Case "1523838"		FItemName = FItemName & " 정글북 안대"
				Case "1523836"		FItemName = FItemName & " 정글북 스탠드"
				Case "1523835"		FItemName = FItemName & " 크림 글래스"
				Case "1523833"		FItemName = FItemName & " 정글북 비치타월"
				Case "1520151"		FItemName = FItemName & " 정글북 러그"
				Case "1520149"		FItemName = FItemName & " 정글북 러그"
				Case "1509355"		FItemName = FItemName & " 앨리스 코스터"
				Case "1488156"		FItemName = FItemName & " 미키 인퓨져"
				Case "1488140"		FItemName = FItemName & " 푸우 인퓨져"
				Case "1473441"		FItemName = FItemName & " 크림 글래스"
				Case "1471075"		FItemName = FItemName & " 앨리스 접시"
				Case "1471073"		FItemName = FItemName & " 앨리스 슈가볼 크리머"
				Case "1422085"		FItemName = FItemName & " 앨리스 찻잔"
				Case "1407891"		FItemName = FItemName & " 앨리스 코지"
				Case "1405564"		FItemName = FItemName & " 앨리스 티스푼포크"
				Case "1405559"		FItemName = FItemName & " 앨리스 티팟"
			End Select
		End If
        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
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
		'//결과 반환
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

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getGmarketItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_gmarket.jpg></p>&#xA;")

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "&#xA;")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "&#xA;")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%"">&#xA;")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%"">&#xA;")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%"">&#xA;")

		'#배송 주의사항
		strRst = strRst & ("&#xA;<img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_gmarket.jpg>")

		strRst = strRst & ("</div>")
		getGmarketItemContParamToReg = strRst

		''2013-06-10 김진영 추가(롯데닷컴처럼 상품이미지가 길면 엑박나오는 현상)
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
		strRst = strRst & "					DefaultImage="""&FbasicImage&""""			'#상품 기본 이미지 URL | 600 × 600 이미지 권장( jpg 이미지 )
		strRst = strRst & "					LargeImage="""&FbasicImage&""""				'상품 큰 이미지 URL | 600 × 600
		strRst = strRst & "					SmallImage="""&FListImage&""""				'상품 작은 이미지 URL | 100 × 100

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
		Else
			getiszeroWonSoldOut = "N"
		End If
	End Function

	Public Function getGmarketShippingParam()
		Dim strRst, NewGroupYn, GroupCode
		NewGroupYn = True
		If Not(NewGroupYn) Then
			strRst = strRst & "				<Shipping "
			strRst = strRst & "					SetType=""New"""			'#배송비 구분 | New : 배송비 그룹번호 신규 생성, Use : 기존 배송비 그룹번호 사용
			strRst = strRst & "					BundleNo=""0"""				'묶음번호 | 배송비 그룹코드를 발주처 기준으로 묶음 배송비로 생성 할 경우 등록 AddAddressBook의 BundleNO로 RefundAddrNum과 정합 관리
			strRst = strRst & "					GroupCode="""""				'배송비 그룹코드 | SetType: Use 인 경우
			strRst = strRst & "					RefundAddrNum=""740092"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"	'반품배송지 번호 | AddAddressBook의 AddressCode
			strRst = strRst & "					<NewItemShipping "
			strRst = strRst & "						FeeCondition=""ConditionalFee"" "	'상품별 배송비 종류 | SetType이 New 인 경우 Free : 무료, ConditionalFee : 조건부무료, FixedFee : 유료, PrepayableOnDelivery : 착불선결제, PayOnDelivery : 착불
			strRst = strRst & "						FeeBasePrice=""30000"""				'상품별 배송비 조건 | 조건부무료일 경우
			strRst = strRst & "						Fee=""2500"""						'조건부무료이거나 유료일 경우
			strRst = strRst & "					/>"
			strRst = strRst & "				</Shipping>"
		Else
			'GroupCode = "389827401"
			GroupCode = "856237774"		'5만원 미만 3천원 배송비코드 2020-01-10 김진영 수정

			strRst = strRst & "				<Shipping "
			strRst = strRst & "					SetType=""Use"""				'#배송비 구분 | New : 배송비 그룹번호 신규 생성, Use : 기존 배송비 그룹번호 사용
			strRst = strRst & "					BundleNo=""0"""					'묶음번호 | 배송비 그룹코드를 발주처 기준으로 묶음 배송비로 생성 할 경우 등록 AddAddressBook의 BundleNO로 RefundAddrNum과 정합 관리
			strRst = strRst & "					GroupCode="""&GroupCode&""""	'배송비 그룹코드 | SetType: Use 인 경우
			strRst = strRst & "					RefundAddrNum=""740092"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"	'반품배송지 번호 | AddAddressBook의 AddressCode
'			strRst = strRst & "					<NewItemShipping "
'			strRst = strRst & "						FeeCondition=""Free or ConditionalFee or PayOnDelivery or PrepayableOnDelivery or FixedFee"" "
'			strRst = strRst & "						FeeBasePrice=""decimal"""
'			strRst = strRst & "						Fee=""decimal"""
'			strRst = strRst & "					/>"
			strRst = strRst & "				</Shipping>"
		End If
		getGmarketShippingParam = strRst
	End Function

	'기본정보 Gmarket 등록 soap XML
	Public Function getGmarketItemRegParameter(isReg)
		Dim strRst, tt, isMadeInKorea
		If Fsourcearea = "한국" OR Fsourcearea = "대한민국" Then
			isMadeInKorea = "Domestic"		'국내
		Else
			isMadeInKorea = "Foreign"		'수입
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
		strRst = strRst & "				OutItemNo="""&FItemid&""""							'#외부상품번호 | 제휴사 상품 번호
		strRst = strRst & "				CategoryCode="""&FDepthCode&""""					'#소분류코드
		If isReg Then
			strRst = strRst & "			GmktItemNo="""&FGmarketGoodNo&""""					'G마켓 상품번호 | 상품정보 수정시
		End If
		strRst = strRst & "				ItemName="""&getItemNameFormat&""""					'#상품명
'		strRst = strRst & "				ItemEngName=""string"""								'영문상품명
		strRst = strRst & "				ItemDescription="""""								'#상품상세정보
		strRst = strRst & "				GdHtml="""&replaceRst(getGmarketItemContParamToReg)&""""		'New 상품 상세정보 - 상품정보
'		strRst = strRst & "				GdHtml=""string"""									'New 상품 상세정보 - 상품정보
'		strRst = strRst & "				GdAddHtml=""string"""								'New 상품 상세정보 - 추가구성
'		strRst = strRst & "				GdPrmtHtml=""string"""								'New 상품 상세정보 - 광고/홍보
		strRst = strRst & "				MakerNo="""&MAKERNO&""""							'#제조사번호 | 우선 기타로 전부 설정했음
'		strRst = strRst & "				BrandNo="""&FBrandCode&""""							'브랜드번호
		strRst = strRst & "				BrandNo=""100356"""									'브랜드번호  2019-02-21 16:16 김진영 수정(텐바이텐 코드(100356)로 픽스)
'		strRst = strRst & "				ModelName=""string"""								'모델명
		strRst = strRst & "				IsAdult="""&Chkiif(IsAdultItem() = "Y", "true", "false")&""""	'#성인용품 여부 | true, false
		strRst = strRst & "				Tax="""&CHKIIF(FVatInclude="N","Free","VAT")&""""	'#부가세 면세여부 | VAT, Free
'		strRst = strRst & "				MadeDate=""date"""									'제조(출판)년월일
'		strRst = strRst & "				AppearedDate=""date"""								'출시년월
		strRst = strRst & "				ExpirationDate=""2078-12-31"""						'#유효일 | ex. 2011-01-02 | 1/1/1900 12:00:00 AM and 6/6/2079 11:59:59 PM.
'		strRst = strRst & "				FreeGift=""string"""								'사은품
		strRst = strRst & "				ItemKind=""Shipping"""								'#상품종류 | Shipping: 배송상품 / Ecoupon: 이쿠폰상품
'		strRst = strRst & "				InventoryNo=""string"""								'판매자관리코드 | 상품 정보가 변경된 시점(상품정보 , 가격정보)에 등록된 code를 주문정보에 포함하여 전달
'		strRst = strRst & "				ItemWeight=""double"""								'상품 무게
		strRst = strRst & "				IsOverseaTransGoods=""false"""						'해외배송 가능 여부 | True : 전체 국가 배송 가능, False : 전체 국가 배송 불가
'		strRst = strRst & "				IsGift=""false"""									'선물하기 상품 여부 | 선물하기 상품 여부 선택 true: 선물하기 가능 false: 선물하기 불가능 * 미입력/skip 시 default = true 로 등록
'		strRst = strRst & "				FreeDelFeeType=""int"""								'무료배송비 타입 | 1 : 지역별 차등 무료, 2 : 설치 배송비, 3 : 직접 수령 가능 상품
'		strRst = strRst & "				IsGmktDiscount=""boolean"""							'G마켓 할인 적용 여부 | True : 미입력시 할인 동의, False인 경우 G마켓 부담 할인 적용 불가
		strRst = strRst & "				>"
'		strRst = strRst & "				<ReferencePrice "
'		strRst = strRst & "					Kind=""Quotation or Department or HomeShopping"""					'참고가격 종류 | Quotation : 시중가, Department : 백화점가, HomeShopping : 홈쇼핑가
'		strRst = strRst & "					Price=""decimal"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'참고가격
		strRst = strRst & "				<Refusal "
		strRst = strRst & "					IsPriceCompare=""false"""											'가격비교 노출제외 | true, false ..2017-02-17 true->false로 수정
		strRst = strRst & "					IsNego=""true"""													'흥정하기 노출제외 | true, false-
		strRst = strRst & "					IsJaehuDiscount=""true"""											'제휴할인 제한 | true, false
		strRst = strRst & "					IsPack=""false"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'장바구니 불가 | True 인 경우 장바구니 비 노출 처리이며, 최초 false로 연동 할 경우는 이슈가 없으나, 최초true로 넣은 후 상품 수정시 false 또는 null로 보낼 경우 장바구니 노출 제한이 풀림
		strRst = strRst & getGmarketAddImageParam()
		strRst = strRst & "				<As Telephone=""1644-6035"" Address=""Seller"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'#연락처, AS센터 주소/정보 | Manufacturing_Seller : 제조사AS 센터나 판매자에게 문의, Seller : 판매자에게 문의
		strRst = strRst & getGmarketShippingParam()
'		strRst = strRst & "				<BundleOrder BuyUnitCount=""int"" MinBuyCount=""int"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'BuyUnitCount : 최소구매수량, MinBuyCount : 구매수량단위
		strRst = strRst & "				<OrderLimit OrderLimitCount="""&getOrderMaxNum&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'OrderLimitCount : 최대구매가능수량, OrderLimitPeriod : 구매수량확인
	If FDepth4Code <> "" Then
		strRst = strRst & "				<AttributeCode AttributeCode="""&FDepth4Code&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'분류속성코드
	End If
		strRst = strRst & "				<Origin Code="""&isMadeInKorea&""" Place="""&Fsourcearea&""" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'원산지 구분 | Domestic : 국내, Foreign : 국외, Etc : 모름...원산지명
'		strRst = strRst & "				<Book ISBN=""string"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"		'도서 ISBN 코드 | ISBN 등록시 주문 옵션 등록 불가
		strRst = strRst & "				<GoodsKind GoodsKind=""New"" GoodsStatus=""NotUsed"" GoodsTag=""Default"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"	'상품상태 | New : 신상품, Used : 중고상품...NotUsed : 미사용, AlmostNew : 거의 새것, Fine : 양호, Old : 약간 낡음, ForCollect : 사용 불가(수집용)
'		strRst = strRst & "				<GoodsKind GoodsKind=""Unknown or New or Stock or Used or Returned or Displayed or Refurbished"" GoodsStatus=""None or Under3Months or Under6Months or Under1Year or Over2Years or NotUsed or AlmostNew or Fine or Old or ForCollect or Sealed or Unsealed or UsedAfterUnsealed or DisplayedNotUsed or DisplayedAlmostNew or DisplayedFine or DisplayedOld or DisplayedForCollect"" GoodsTag=""Default or New or Hot or Sale or MDRecommend or InterestFree or Limited or Gift or LowestPrice or NoMargin or Donation or SpecialBargain or EyeCatch or PowerDealer or Premium2Days or Premium7Days or Premium2Weeks or Premium1Month or Premium2Month or Premium3Month or ImmediateDelivery or Patronage or PremiumPlus"" xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"" />"
		strRst = strRst & "			</AddItem>"
		strRst = strRst & "		</AddItem>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
'response.write strRst
'response.end
		getGmarketItemRegParameter = strRst
	End Function

	'옵션등록 Soap XML
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
		strRst = strRst & "					<ItemText Name=""텍스트를 입력하세요"" />"
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

		If FOptionCnt = 0 Then			'단품
			strRst = "<ItemSelectionList IsInventory=""true"" IsCombination="""&IsCombination&""" OptionImageLevel=""0""  xmlns=""http://tpl.gmarket.co.kr/tpl.xsd""></ItemSelectionList>"
		Else							'옵션있는 상품
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

			'1. strRst 셀렉션 시작
			strRst = ""
			strRst = strRst & "				<ItemSelectionList "
			strRst = strRst & "					IsInventory=""true"""					'#재고사용여부 | 옵션별 재고 사용시 필수
			strRst = strRst & "					OptionSortType=""Register"""			'#정렬순서 | Register(등록순), Price(가격순), Name(이름순)
			strRst = strRst & "					IsCombination="""&IsCombination&""""	'#조합형 사용 | 미 입력시 False로 처리, True인 경우 조합형 옵션 등록
			strRst = strRst & "					OptionImageLevel=""0"""					'New옵션 사용이 True인 경우 필수, 0 : 이미지 매칭 미 사용, 1 : 옵션명 1에 매칭, 2 : 옵션명 2에 매칭
			strRst = strRst & "					xmlns=""http://tpl.gmarket.co.kr/tpl.xsd"">"
			'1.strRst 셀렉션 시작 끝

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
						strRst = strRst & "					Name="""&arrMultiTypeNm&""""				'#정보명 | 항목수는 최대 5개,총 선택수가 최대 500개, 옵션명 별 최대 25자 New옵션 사용 여부가 True인 경우 구분자 등록 처리  ex) 색상^|^사이즈^|^사은품
					Else
						If db2Html(rsget("optionTypeName")) <> "" Then
							optNm = db2Html(rsget("optionTypeName"))
						Else
							optNm = "옵션"
						End If
						strRst = strRst & "					Name="""&optNm&""""							'#정보명 | 항목수는 최대 5개,총 선택수가 최대 500개, 옵션명 별 최대 25자 New옵션 사용 여부가 True인 경우 구분자 등록 처리  ex) 색상^|^사이즈^|^사은품
					End If
					strRst = strRst & "						Code="""&itemoption&"""" 					'옵션 판매자 코드
					strRst = strRst & "						Value="""&Replace(replace(optDc, ",", "^|^"), "선택안함", "선택안함.")&"""" 	'#정보값 | 선택 조건 별 최대 10자, New옵션 사용 여부가 True인 경우 구분자 등록 처리 등록되는 옵션 값의 개수는 옵션명의 구분과 동일 하야 함, ex) 빨강^|^90^|^밥솥
					'strRst = strRst & "					Value="""&replace(optDc, ",", "^|^")&"""" 	'#정보값 | 선택 조건 별 최대 10자, New옵션 사용 여부가 True인 경우 구분자 등록 처리 등록되는 옵션 값의 개수는 옵션명의 구분과 동일 하야 함, ex) 빨강^|^90^|^밥솥
					strRst = strRst & "						Price="""&optaddprice&"""" 					'#가격 | 항목별로 가격이 0원인 것이 1개 이상 존재, 판매가격의 -50% ~ +100% 이내. 가격은 10원 단위, ‘,’ 입력 불가
					strRst = strRst & "						Remain="""&optLimit&""""			 		'#재고수량
					If IsCombination = "true" Then
						strRst = strRst & "					OptionImageUrl=""-"""						'옵션 이미지 URL | 조합형 옵션 등록시 0 : 이미지 매칭 미사용일 경우에도 값(문자1개)를 넣어 해당 field호출
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

	'이미지 수정 Soap XML
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

	Public Function getGmarketAddPriceParameter(isReged, mustPrice, idisplayDate)		''뭔가 XML수정이 필요하다면..incGmarketFunction의 getGmarketAddPriceParameter도 같이 수정
		Dim strSQL, strRst, GetTenTenMargin, iStockQty

		'노출 기간 설정
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

		'재고 수량 설정
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
		strRst = strRst & "				GmktItemNo="""&FGmarketGoodNo&""""			'#G마켓 상품번호
		strRst = strRst & "				DisplayDate="""&idisplayDate&""""			'#주문기간 | 최대 1년
		strRst = strRst & "				SellPrice="""&mustPrice&""""				'#판매가격 | 최소 100원 이상 1,000,000,000원 미만 (100원 단위)
		strRst = strRst & "				StockQty="""&iStockQty&""""					'#재고수량
		strRst = strRst & "				InventoryNo="""&FItemid&""" />"				'판매자관리코드
		strRst = strRst & "		</AddPrice>"
		strRst = strRst & "	</soap:Body>"
		strRst = strRst & "</soap:Envelope>"
		getGmarketAddPriceParameter = strRst
	End Function

	'G9 등록 soap XML
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
		strRst = strRst & "				GmktItemNo="""&FGmarketGoodNo&""""			'#G마켓 상품번호
		strRst = strRst & "				SellManageYn=""N"""							'#판가/재고 관리 여부 | Y : G9 판매용 가격/재고 사용 리턴되는 복제 상품코드로 기존 가격 재고 API를 통해 관리 (addprice), N : 복제 대상 상품의 가격/재고를 사용 - 복제 대상 상품 가격/재고 변경시 복제 상품의 가격/재고도 동일하게 동기화 처리 - 미 입력시 False로 처리
		strRst = strRst & "				CostManageYn=""N"""							'#G9 판매자 할인 여부 | Y : G9 판매용 별도 할인 정책 사용 리턴되는 복제 상품코드로 기존 할인 API를 통해 관리 (AddPremiumItem), N : 복제 대상 상품의 할인 정책 유지 복제 대상 할인 정책 변경시 복제 상품의 할인 정책도 동일하게 동기화 처리 - 미 입력시 False로 처리
		strRst = strRst & "				ItemManageYn=""N"" />"						'#기본상품정보 관리 여부 | Y : G9 판매용 별도 상품정보 사용 리턴되는 복제 상품코드로 기존 상품정보 등록수정 API를 통해 관리 (AddItem), N : 복제 대상 상품의 상품 기본정보를 사용 - 복제 대상 상품의 상품명, 카테고리정보, 배송비 정보 등의 정보를 동일하게 동기화 처리
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

	'// 미등록 상품 목록(등록용)
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
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "					'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.itemid in ('2594140', '2594138', '2594139', '2557733' , '2558483', '2549730', '2549728') THEN 'Y' "	'2019-12-02 윤현주..등록마진이하이나 등록 요청'
		strSql = strSql & " 				WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"' and i.mwdiv <> 'M')) "	'등록제외 카테고리
		strSql = strSql & "	and isnull(R.APIadditem, '') <> 'Y' "									'기본정보 등록되있으면 등록하면 안 됨
		strSql = strSql & "	and isnull(R.GmarketGoodNo, '') = '' "
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
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

	'// 미등록 옵션(등록용)
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

	'// 수정용 이미지
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
		strSql = strSql & " and m.GmarketGoodNo is Not Null "									'#등록 상품만
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
			'선택상품이 있다면
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
		strSql = strSql & " and m.GmarketGoodNo is Not Null "									'#등록 상품만
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

	'//G9 미등록 상품 목록(등록용)
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

'지마켓 상품코드 얻기
Function getGmarketGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 Gmarketgoodno FROM db_etcmall.dbo.tbl_gmarket_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getGmarketGoodno = rsget("Gmarketgoodno")
	rsget.Close
End Function

'지마켓 어린이 인증 카테고리 인지 확인
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

'// 상품이미지 존재여부 검사
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
