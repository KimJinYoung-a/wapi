<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "cjmall"
CONST CMAXLIMITSELL = 5        '' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CCJMALLMARGIN = 12       ''마진 12%...// 왜 12? // 2013-11-05 김진영..12->15로 수정 =>12로 수정 유미희.(2013/11/21)
CONST CitemGbnKey ="K1099999" ''상품구분키 ''하나로 통일
CONST CUPJODLVVALID = True   ''업체 조건배송 등록 가능여부

CONST CVENDORID = 411378					'협력업체코드
CONST CVENDORCERTKEY = "CJ03074113780"		'인증키
CONST CUNIQBRANDCD = 24049000				'브랜드코드
CONST MD_CODE = "6648"						'MD_Code | 2015-10-14일전까지 5103

Class CCJMallItem
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
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FcjmallPrdNo
	Public Fcjmallprice
	Public FcjmallSellYn
	Public FaccFailCnt
	Public FlastErrStr
	Public FsafetyDiv
	Public FsafetyNum
	Public FitemGbnKey
	Public FcjmallStatCD
	Public FRectMode
	Public Fdeliverfixday
	Public Fdeliverytype
	Public Fsocname_kor
	Public Fcddkey

	Public FItemOption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold
	Public FmaySoldOut

	Public MustPrice
	Public FAdultType
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "9999" Then
			getOrderMaxNum = 9999
		End If
	End Function

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

	Public Function IsRegedOptionSellyn
		Dim sqlStr, sellynCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_item.dbo.tbl_Outmall_regedoption WHERE itemid="&FItemid&" and mallid = 'cjmall' and outmallSellyn = 'Y' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			sellynCnt = rsget("cnt")
		rsget.Close

		If (sellynCnt = 0) Then
			IsRegedOptionSellyn = "N"
		Else
			IsRegedOptionSellyn = "Y"
		End If
	End Function

    Function getCJmallSuplyPrice2()
'        getCJmallSuplyPrice2 = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'하단은 CJ메뉴얼에 적힌 내용
		'* 마진율 확인요함
		'1. 과세상품 : 매입원가(VAT제외) = Round(판매가/1.1 - 0.1 * (판매가/1.1)), 0)
		'2. 면세상품 : 매입원가(VAT제외) = Round(판매가 - 0.1 * 판매가, 0)
		Dim CJMargin
		CJMargin = CCJMALLMARGIN
		If (Now() > #06/13/2016 00:00:00# AND Now() < #06/22/2016 23:59:59#) Then
			If getMarginChgCategory = "Y" Then
				CJMargin = 15
			End If
		End If

		If FVatInclude = "Y" Then		'과세
			getCJmallSuplyPrice2 = Round((MustPrice) /1.1 - (CJMargin/100) * ((MustPrice)/1.1))
		Else							'면세
			getCJmallSuplyPrice2 = Round((MustPrice) - (CJMargin/100) * (MustPrice))
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

        buf = replace(FItemName,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        getItemNameFormat = buf
    end function

	Public Function getItemKeyword
		Dim p, spKey, tmpU, arrKeyword, keyCnt
		spKey = ""
		If trim(Fkeywords) = "" Then Exit Function

		arrKeyword = Split(Fkeywords, ",")
		keyCnt = Ubound(Split(Fkeywords, ","))

		If keyCnt >= 3 Then
			tmpU = 3
		Else
			tmpU = Ubound(Split(Fkeywords, ","))
		End If

		For p=0 to tmpU
			spKey = spKey&arrKeyword(p)&","
		Next

		If Right(spKey,1) = "," Then
			spKey = Left(spKey,Len(spKey)-1)
		End If
		spKey = replace(spKey,",",";")
		getItemKeyword = "텐바이텐;"&spKey
	End Function

	'화물배송 관련
	Public Function getdeliverfixday()
		If (Fdeliverfixday = "C") or (Fdeliverfixday = "X") or (Fdeliverfixday = "G") Then
			getdeliverfixday = 20
		Else
			getdeliverfixday = 10
		End If
	End Function

    '협조문 번호 No. 28566 에 따른 수수료 변경할 관리 카테고리 12->15%
    Public Function getMarginChgCategory()
		dim ret, isCate
'        ret = (FtenCateLarge="010")															'디자인문구
'		ret = ret or (FtenCateLarge="020")													'오피스/개인소품
'		ret = ret or (FtenCateLarge="025" and FtenCateMid="117")							'디지털	아이폰6/플러스 케이스
'		ret = ret or (FtenCateLarge="025" and FtenCateMid="118")							'디지털	갤럭시노트4/엣지 케이스
'		ret = ret or (FtenCateLarge="025" and FtenCateMid="120")							'디지털	갤럭시S6 케이스
'		ret = ret or (FtenCateLarge="030")													'키덜트
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="010")							'여행/취미	캐리어
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="011")							'여행/취미	트래블백
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="012")							'여행/취미	수납용품
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="013")							'여행/취미	여행소품
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="014")							'여행/취미	여행 편의용품
'		ret = ret or (FtenCateLarge="035" and FtenCateMid="021")							'여행/취미	애완용품
'		ret = ret or (FtenCateLarge="050")													'홈/데코
'		ret = ret or (FtenCateLarge="045")													'수납/생활
'		ret = ret or (FtenCateLarge="060")													'키친
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="030")							'가방/슈즈/쥬얼리	패션슈즈
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="010")							'가방/슈즈/쥬얼리	패션가방
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="140")							'가방/슈즈/쥬얼리	캐주얼가방
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="201")							'가방/슈즈/쥬얼리	백팩
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="202")							'가방/슈즈/쥬얼리	파우치
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="150")							'가방/슈즈/쥬얼리	지갑
'		ret = ret or (FtenCateLarge="070" and FtenCateMid="050")							'가방/슈즈/쥬얼리	패션소품
'		ret = ret or (FtenCateLarge="100")													'베이비

		ret = (FtenCateLarge="035" and FtenCateMid="021")									'여행/취미	애완용품
'		ret = ret or (FtenCateLarge="020")													'오피스/개인소품
'		ret = ret or (FtenCateLarge="025")													'디지털
'		ret = ret or (FtenCateLarge="030")													'키덜트
'		ret = ret or (FtenCateLarge="035")													'여행/취미
'		ret = ret or (FtenCateLarge="045")													'수납/생활
'		ret = ret or (FtenCateLarge="050")													'홈/데코
'		ret = ret or (FtenCateLarge="060")													'키친
'		ret = ret or (FtenCateLarge="070")													'가방/슈즈/쥬얼리
'		ret = ret or (FtenCateLarge="080")													'Women
'		ret = ret or (FtenCateLarge="090")													'Men
'		ret = ret or (FtenCateLarge="100")													'베이비

		If ret Then
			isCate = "Y"
		Else
			isCate = "N"
		End If
        getMarginChgCategory = isCate
    End Function

    ''주문제작 여부
    Public Function getzCostomMadeInd()
		dim ret, CMadeInd
        ret = (Fitemdiv="06" or Fitemdiv="16")
        ret = ret or (FtenCateLarge="010" and FtenCateMid="070" and FtenCateSmall="070")	'디자인문구	스탬프	주문제작
		ret = ret or (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010")	'여행/취미	드라이브	주차판
		ret = ret or (FtenCateLarge="040")													'가구
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001")	'수납/생활	보관/정리용품	수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	'수납/생활	보관/정리용품	틈새수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005")	'수납/생활	보관/정리용품	잡지꽂이
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'수납/생활	보관/정리용품	벽걸이수납함
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	'수납/생활	보관/정리용품	벽걸이수납함
		ret = ret or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019")	'수납/생활	보관/정리용품	이동식수납장
		ret = ret or (FtenCateLarge="045" and FtenCateMid="003")							'수납/생활	데스크수납
		ret = ret or (FtenCateLarge="045" and FtenCateMid="006")							'수납/생활	데코수납
		ret = ret or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	'수납/생활	키즈수납	키즈 서랍장
		ret = ret or (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050")	'홈/데코	조명	이니셜/메세지조명
		ret = ret or (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010")	'홈/데코	장식소품	이니셜장식
		ret = ret or (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120")	'홈/데코	홈갤러리	수작업 주문제작
		ret = ret or (FtenCateLarge="055" and FtenCateMid="070")							'패브릭 > 침구세트
		ret = ret or (FtenCateLarge="055" and FtenCateMid="080")							'패브릭 > 커튼
		ret = ret or (FtenCateLarge="055" and FtenCateMid="090")							'패브릭 > 쿠션/방석
		ret = ret or (FtenCateLarge="055" and FtenCateMid="100")							'패브릭 > 매트/러그
		ret = ret or (FtenCateLarge="055" and FtenCateMid="110")							'패브릭 > 패브릭소품
		ret = ret or (FtenCateLarge="055" and FtenCateMid="120")							'패브릭 > 침구단품
		ret = ret or (FtenCateLarge="060" and FtenCateMid="130")							'키친 > 작가 생활자기
		ret = ret or (FtenCateLarge="070" and FtenCateMid="160")							'가방/슈즈/쥬얼리 > 쥬얼리
		ret = ret or (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010")	'Men > 쥬얼리/잡화 > 시계/쥬얼리
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020")	'베이비 > 가구/침구/수납 > 데코스티커/벽지
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040")	'베이비 > 가구/침구/수납 > 수납함/책꽂이
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	'베이비 > 가구/침구/수납 > 의자
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060")	'베이비 > 가구/침구/수납 > 조명/액자
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066")	'베이비 > 가구/침구/수납 > 테이블/책상
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070")	'베이비 > 가구/침구/수납 > 안전용품
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	'베이비 > 가구/침구/수납 > 아기침대
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110")	'베이비 > 가구/침구/수납 > 플레이매트
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120")	'베이비 > 가구/침구/수납 > 블랑켓/아기담요
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130")	'베이비 > 가구/침구/수납 > 모빌
		ret = ret or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140")	'베이비 > 가구/침구/수납 > 쿠션/침구/커튼
		If ret Then
			CMadeInd = "Y"
		Else
			CMadeInd = "N"
		End If
        getzCostomMadeInd = CMadeInd
    End Function

    ''리드타임 얻기
    Public Function getzLeadTime()
		If (FtenCateLarge="040") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="001") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="002")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="005") or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="010")	or (FtenCateLarge="045" and FtenCateMid="002" and FtenCateSmall="019") or (FtenCateLarge="045" and FtenCateMid="003")	or (FtenCateLarge="045" and FtenCateMid="006") or (FtenCateLarge="045" and FtenCateMid="007" and FtenCateSmall="008")	or (FtenCateLarge="055" and FtenCateMid="070") or (FtenCateLarge="055" and FtenCateMid="080")	or (FtenCateLarge="055" and FtenCateMid="090") or (FtenCateLarge="055" and FtenCateMid="100")	or (FtenCateLarge="055" and FtenCateMid="110") or (FtenCateLarge="055" and FtenCateMid="120")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="040") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="050")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="066") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="100")	or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="120") or (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="140") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="020") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="060") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="070") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="110") OR (FtenCateLarge="100" and FtenCateMid="060" and FtenCateSmall="130") OR (FtenCateLarge = "050" and FtenCateMid = "120" and FtenCateSmall = "080") OR (FtenCateLarge = "050" and FtenCateMid = "045" and FtenCateSmall = "100") OR (FtenCateLarge = "070" and FtenCateMid = "070") OR (FtenCateLarge = "070" and FtenCateMid = "160") Then
			getzLeadTime = "15"
		ElseIf (FtenCateLarge = "010" and FtenCateMid = "070" and FtenCateSmall = "070") OR (FtenCateLarge="035" and FtenCateMid="016" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="010" and FtenCateSmall="050") OR (FtenCateLarge="050" and FtenCateMid="030" and FtenCateSmall="010") OR (FtenCateLarge="050" and FtenCateMid="045" and FtenCateSmall="120") OR (FtenCateLarge="060" and FtenCateMid="130") OR (FtenCateLarge="070" and FtenCateMid="160") OR (FtenCateLarge="090" and FtenCateMid="070" and FtenCateSmall="010") Then
			getzLeadTime = "03"
		Else
			getzLeadTime = "03"		'이 구문없는 데 추가..2021-03-25 진영
		End If
	End Function

    public Function IsCjFreeBeasong()
        IsCjFreeBeasong = False
    end Function

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

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold < CLIMIT_SOLDOUT_NO))
	End Function

	'// 상품등록: 옵션 파라메터 생성(상품등록용)
	Public Function getCJOptionParamToReg()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin, i, specialPrice, ownItemCnt
		strSql = ""
		strSql = strSql & " SELECT mustPrice "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		strSql = strSql & " WHERE mallgubun = '"& CMALLNAME &"' "
		strSql = strSql & " and itemid = '"& Fitemid &"' "
		strSql = strSql & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CNT "
		strSql = strSql & " FROM db_partner.dbo.tbl_partner "
		strSql = strSql & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : 수입
		strSql = strSql & " and id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ownItemCnt = rsget("CNT")
		End If
		rsget.Close

		If specialPrice <> "" Then
			MustPrice = specialPrice
		ElseIf ownItemCnt > 0 Then
			MustPrice = Forgprice
		Else
			'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 시작
			GetTenTenMargin = CLng(10000 - Fbuycash / FSellCash * 100 * 100) / 100
			If GetTenTenMargin < FOutmallstandardMargin Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
			'2013-07-24 김진영//텐텐마진이 CJMALL의 마진보다 작을 때 orgprice로 전송 끝
		End If

		optaddprice		= 0
		strSql = ""
		strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
		strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
		strSql = strSql & " WHERE i.itemid = "&Fitemid
		strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
					FItemOption = "0000"
					'optionname = DdotFormat(chrbyte(getItemNameFormat,40,""),20)
					optionname = "단일상품"
					itemSu = GetCJLmtQty
					optaddprice		= 0
				Else
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo

					if rsget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if
				End If

				If rsget("deliverfixday") = "C" OR rsget("deliverfixday") = "X" OR rsget("deliverfixday") = "G" Then
					fixday = "60"
				Else
					fixday = "20"
				End If
				strRst = strRst &"	<tns:unit>"
				''strRst = strRst &"		<tns:unitNm><![CDATA["&DDotFormat(optionname, 16)&"]]></tns:unitNm>"	'단품정보 - 단품상세(옵션명을 텍스트로 넘기면 됨)
				strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"
				strRst = strRst &"		<tns:unitRetail>"&MustPrice+optaddprice&"</tns:unitRetail>"				'단품정보 - 판매가
				strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"					'단품정보 - 매입원가
				strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'단품정보 - 공급가능수량 (상품 재고 파악이 안되는경우는 999같은 숫자를 넣습니다.)
			If getzCostomMadeInd = "Y" Then
				strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"						'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
'			ElseIf Left(FCddkey,2) = "35" OR Left(FCddkey,2) = "37" Then										'상품등록시 대분류값(35 전기전자/37 정보통신)일경우 리드타임의 값은 '02' 등록만 가능하도록 처리되어있습니다.
'				strRst = strRst &"		<tns:leadTime>02</tns:leadTime>"										'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
			Else
				strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"										'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
			End If
				strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'단품정보 - 적용사유 (10 : 적용안함, 20 : 상품포장, 30 : 상품생산, 40 : 입고검사, 50 : 출고검사, 60 : 설치상품)
				strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'단품정보 - 판매시작일자
				strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'단품정보 - 판매종료일자 (판매상태수정에서..)
			If (application("Svr_Info")="Dev") OR (Fitemid="899506") Then
				strRst = strRst &"		<tns:vpn>"&rsget("itemid")&"_Q"&FItemOption&"</tns:vpn>"				'단품정보 - 협력사상품코드(899506만 Q라는 문자삽입)
			Else
				strRst = strRst &"		<tns:vpn>"&rsget("itemid")&"_"&FItemOption&"</tns:vpn>"					'단품정보 - 협력사상품코드
			End If
				strRst = strRst &"	</tns:unit>"
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getCJOptionParamToReg = strRst
	End Function

	public function GetCJLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetCJLmtQty = 0
			Else
				GetCJLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetCJLmtQty = 999
		End If
	End Function

    Function getCJmallSuplyPrice(optaddprice)
'        getCJmallSuplyPrice = CLNG(FSellCash * (100-CCJMALLMARGIN) / 100)
		'하단은 CJ메뉴얼에 적힌 내용
		'* 마진율 확인요함
		'1. 과세상품 : 매입원가(VAT제외) = Round(판매가/1.1 - 0.1 * (판매가/1.1)), 0)
		'2. 면세상품 : 매입원가(VAT제외) = Round(판매가 - 0.1 * 판매가, 0)
		Dim CJMargin
		CJMargin = CCJMALLMARGIN
		If (Now() > #04/05/2018 00:00:00# AND Now() < #04/22/2018 23:59:59#) Then
			If getMarginChgCategory = "Y" Then
				CJMargin = 15
			End If
		End If

		If FVatInclude = "Y" Then		'과세
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) /1.1 - (CJMargin/100) * ((MustPrice+optaddprice)/1.1))
		Else							'면세
			getCJmallSuplyPrice = Round((MustPrice+optaddprice) - (CJMargin/100) * (MustPrice+optaddprice))
		End If
    End Function

	'// 상품등록: MD상품군 및 전시 카테고리 파라메터 생성(상품등록용)
	Public Function getCjCateParamToReg()
		Dim strSql, strRst, i
		strSql = ""
		strSql = strSql & " SELECT top 100 c.CateKey "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_cjmall_cate_mapping as m "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_cjMall_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " and c.isusing ='Y' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : 브랜드 / D : 일반
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = ""
			i = 0
			Do until rsget.EOF
				If i = 0 Then
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:mainInd>Y</tns:mainInd>"
					strRst = strRst &"			<tns:ctgName>" & rsget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				Else
					strRst = strRst &"		<tns:mallCtg>"
					strRst = strRst &"			<tns:ctgName>" & rsget("CateKey") & "</tns:ctgName>"
					strRst = strRst &"		</tns:mallCtg>"
				End If
				rsget.MoveNext
				i = i + 1
			Loop
		End If
		rsget.Close
		getCjCateParamToReg = strRst
	End Function

	Public Function getCjCertParamToNewReg()
		Dim strRst, strSql, certNum, certCode, certCateCd, certDate, modelName, certRegYn, certOrganName
		strSql = ""
		strSql = strSql & " SELECT TOP 1 r.certNum "
		strSql = strSql & "	,Case When r.safetyDiv in ('10', '40') THEN '400021' "
		strSql = strSql & "		  When r.safetyDiv in ('20', '50') THEN '400022' "
		strSql = strSql & " 	  When r.safetyDiv in ('30', '60') THEN '400023' "
		strSql = strSql & "		  When r.safetyDiv in ('70') THEN '400017' "
		strSql = strSql & "		  When r.safetyDiv in ('80') THEN '400018' "
		strSql = strSql & "		  When r.safetyDiv in ('90') THEN '400020' end as certCode "
		strSql = strSql & "	,Case When r.safetyDiv in ('10', '20', '30', '40', '50', '60') THEN '001' Else '002' End as certCateCd "
		strSql = strSql & " ,convert(date, f.certDate) as certDate, f.modelName, f.certOrganName  " & vbcrlf
		strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg as r " & vbcrlf
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on r.itemid = f.itemid " & vbcrlf
		strSql = strSql & " WHERE r.itemid='"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			certNum			= rsget("certNum")
			certCode		= rsget("certCode")
			certCateCd		= rsget("certCateCd")
			certDate		= rsget("certDate")
			modelName		= rsget("modelName")
			certOrganName	= rsget("certOrganName")
			certRegYn = "Y"
		Else
			certRegYn = "N"
		End If
		rsget.Close

		If certRegYn = "Y" then
			strRst = strRst &"	<tns:cert>"																'QC에러 해결하려면 아래정보가 필요한듯..(2013-06-04 김진영)
			strRst = strRst &"		<tns:certCode>"&certCode&"</tns:certCode>"							'품질인증정보 - 항목코드
			strRst = strRst &"		<tns:certSeq>1</tns:certSeq>"										'품질인증정보 - 순번
			strRst = strRst &"		<tns:certCateCd>"&certCateCd&"</tns:certCateCd>"					'품질인증정보 - 분류코드
			strRst = strRst &"		<tns:certNo>"&certNum&"</tns:certNo>"								'품질인증정보 - 인증번호 - 길이제약(50)
'			strRst = strRst &"		<tns:issueDate>2012-06-04</tns:issueDate>"							'품질인증정보 - 발급일자
			strRst = strRst &"		<tns:certDate>"&certDate&"</tns:certDate>"         					'품질인증정보 - 인증일자
'			strRst = strRst &"		<tns:avlStartDate>2012-06-04</tns:avlStartDate>"					'품질인증정보 - 유효기간(FROM)
'			strRst = strRst &"		<tns:avlEndDate>2013-06-04</tns:avlEndDate>"      					'품질인증정보 - 유효기간(TO)
			strRst = strRst &"		<tns:itemModel>"&modelName&"</tns:itemModel>"        				'품질인증정보 - 상품명 및 모델명	-길이제약(200)
			strRst = strRst &"		<tns:orgCode>"&certOrganName&"</tns:orgCode>"            			'품질인증정보 - 인증검사기관명		-길이제약(200)
'			strRst = strRst &"		<tns:certField>전기제품</tns:certField>"								'품질인증정보 - 인증분야			-길이제약(200)
'			strRst = strRst &"		<tns:originCode>원산지</tns:originCode>"     						'품질인증정보 - 원산지(제조국)
'			strRst = strRst &"		<tns:certSpec>세부</tns:certSpec>"          							'품질인증정보 - 세부사항			-길이제약(2000)
			strRst = strRst &"	</tns:cert>"
'2019-03-27 12:00 김진영..새로바뀐 전안법 data가 아니면 전안법정보 미전송으로 변경
		' Else
		' 	If FsafetyNum <> "" AND FsafetyDiv <> "" Then
		' 		Select Case FsafetyDiv
		' 			Case "10"
		' 				certCode	= "400021"
		' 				certCateCd	= "001"
		' 			Case "20"
		' 				certCode = "400021"
		' 				certCateCd	= "001"
		' 			Case "30"
		' 				certCode = "400021"
		' 				certCateCd	= "001"
		' 			Case "40"
		' 				certCode = "400021"
		' 				certCateCd	= "001"
		' 			Case "50"
		' 				certCode = "400017"
		' 				certCateCd	= "002"
		' 		End Select

		' 		If certCode <> "" AND certCateCd <> "" AND Len(FsafetyNum) > 5 Then
		' 			strRst = strRst &"	<tns:cert>"																			'QC에러 해결하려면 아래정보가 필요한듯..(2013-06-04 김진영)
		' 			strRst = strRst &"		<tns:certCode>"&certCode&"</tns:certCode>"										'품질인증정보 - 항목코드
		' 			strRst = strRst &"		<tns:certSeq>1</tns:certSeq>"													'품질인증정보 - 순번
		' 			strRst = strRst &"		<tns:certCateCd>"&certCateCd&"</tns:certCateCd>"								'품질인증정보 - 분류코드
		' 			strRst = strRst &"		<tns:certNo>"&FsafetyNum&"</tns:certNo>"										'품질인증정보 - 인증번호 - 길이제약(50)
		' 			strRst = strRst &"	</tns:cert>"
		' 		End If
		' 	End If
		End If
		getCjCertParamToNewReg = strRst
	End Function

	'상품품목정보
    public function getCjmallItemInfoCdToReg()
		Dim strSql, buf, addSql
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		Dim chkInfodiv, chkCdmKey

		strSql = ""
		strSql = strSql & " EXEC db_item.dbo.usp_API_CJMall_InfoCodeMap_Get " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
'			    infotype	= rsget("infotype")
			    infoContent = rsget("infoContent")
'				infocd		= rsget("infocd")
				mallinfodiv = rsget("mallinfodiv")

				buf = buf &"	<tns:goodsReport>"
				buf = buf &"		<tns:pedfId>"&mallinfoCd&"</tns:pedfId>"
				buf = buf &"		<tns:html><![CDATA["&infoContent&"]]></tns:html>"
				buf = buf &"	</tns:goodsReport>"
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getCjmallItemInfoCdToReg = buf
	End Function

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function getCJItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
		'2014-01-17 10:00 김진영 탑 이미지 추가
		'strRst = strRst & ("<p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>")
		'2021-05-28 18:00 김진영 / 전문관링크 제거
		strRst = strRst & ("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></p><br>")

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br>")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br>")
			Case Else
				strRst = strRst & (nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg"">")

		strRst = strRst & ("</div>")
		getCJItemContParamToReg = strRst
		''2013-06-10 김진영 추가(롯데닷컴처럼 상품이미지가 길면 엑박나오는 현상)
		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','cjmall') and linkgbn = 'contents' and itemid = '"&Fitemid&"' " & VBCRLF  '' mallid='cjmall' => mallid in ('','cjmall')
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strRst = rsget("textVal")
			strRst = "<div align=""center""><p><a href=""http://10x10.cjmall.com/ctg/specialshop_brand/main.jsp?ctg_id=292240"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_cjmall.jpg""></a></p><br>" & strRst & "<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_common.jpg""></div>"
			getCJItemContParamToReg = strRst
		End If
		rsget.Close
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성(상품등록용)
	Public Function getCJAddImageParamToReg()
		Dim strRst, strSQL, i
		strRst = ""
		If application("Svr_Info")="Dev" Then
			FbasicImage = "http://61.252.133.2/images/B000151064.jpg"
		End If

		strRst = strRst &"	<tns:image>"
		strRst = strRst &"		<tns:imageMain>"&FbasicImage&"</tns:imageMain>"
		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget

		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType")="0" Then
					strRst = strRst &"		<tns:imageSub"&i&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") &"</tns:imageSub"&i&">"
				End If
				rsget.MoveNext
				If i >= 5 Then Exit For
			Next
		End If
		rsget.Close
		strRst = strRst &"	</tns:image>"
		getCJAddImageParamToReg = strRst
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
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
				strSql = strSql & " 	and (optlimityn='N' or (optlimityn='Y' and optlimitno-optlimitsold>"&CMAXLIMITSELL&")) "
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

	'// 상품상태수정시 옵션이 추가된 경우
	Public Function getCJOptionParamToEdit()
		Dim strSql, strRst, itemSu, itemoption, validSellno, optionname, fixday, optaddprice
		Dim GetTenTenMargin, i, specialPrice, tmpPrice, vBigPrice, vSmallPrice, ownItemCnt
		strSql = ""
		strSql = strSql & " SELECT mustPrice "
		strSql = strSql & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		strSql = strSql & " WHERE mallgubun = '"& CMALLNAME &"' "
		strSql = strSql & " and itemid = '"& Fitemid &"' "
		strSql = strSql & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			specialPrice = rsget("mustPrice")
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as CNT "
		strSql = strSql & " FROM db_partner.dbo.tbl_partner "
		strSql = strSql & " WHERE purchaseType in ('3','5','6') "		'3 : PB, 5 : ODM, 6 : 수입
		strSql = strSql & " and id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
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
			If GetTenTenMargin < FOutmallstandardMargin Then
				MustPrice = Forgprice
			Else
				If (FSellCash < Round(Fcjmallprice * 0.45, 0)) Then
					MustPrice = CStr(GetRaiseValue(Round(Fcjmallprice * 0.45, 0)/10)*10)
				Else
					MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
				End If
			End If
		End If

		Dim zeroCnt
		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as zeroCnt "
		strSql = strSql & " FROM [db_item].[dbo].tbl_OutMall_regedoption "
		strSql = strSql & " WHERE itemid = " & Fitemid
		strSql = strSql & " and mallid = 'cjmall' "
		strSql = strSql & " and itemoption = '0000' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			zeroCnt = rsget("zeroCnt")
		rsget.Close

		optaddprice = 0
		If zeroCnt > 0 Then
			strSql = ""
			strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, isnull(R.outmallOptCode, '') as outmallOptCode, i.deliverfixday, isnull(o.optaddprice,'') as optaddprice " & VBCRLF
			strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid " & VBCRLF
			strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and R.itemoption = isNull(o.itemoption, '0000') " & VBCRLF
			strSql = strSql & " WHERE i.itemid = "&Fitemid
			strSql = strSql & " and R.mallid = 'cjmall' "
			strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		Else
			strSql = ""
			strSql = strSql & " SELECT top 900 i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(40),o.optionname) as optionname, o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, isnull(R.outmallOptCode, '') as outmallOptCode, i.deliverfixday, isnull(o.optaddprice,'') as optaddprice " & VBCRLF
			strSql = strSql & " ,DATALENGTH(o.optionname) as optnmLen" & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			strSql = strSql & " JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF ''LEFT Join => Join
			strSql = strSql & " LEFT JOIN [db_item].[dbo].tbl_OutMall_regedoption as R on i.itemid = R.itemid and R.itemoption = o.itemoption and R.mallid='cjmall' " & VBCRLF
			strSql = strSql & " WHERE i.itemid = "&Fitemid
			strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
		End If
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget("outmallOptCode") = "" Then
					FItemOption 	= rsget("itemoption")
					optionname 		= rsget("optionname")
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					optaddprice		= rsget("optaddprice")
					itemSu = getOptionLimitNo
					If rsget("deliverfixday") = "C" OR rsget("deliverfixday") = "X" OR rsget("deliverfixday") = "G" Then
						fixday = "60"
					Else
						fixday = "20"
					End If

                    if rsget("optnmLen")>40 then
					    optionname=DdotFormat(optionname,20)
					end if

					If itemSu <> 0 Then
						strRst = strRst &"	<tns:unit>"
						strRst = strRst &"		<tns:unitNm><![CDATA["&optionname&"]]></tns:unitNm>"					'단품정보 - 단품상세(옵션명을 텍스트로 넘기면 됨)
						strRst = strRst &"		<tns:unitRetail>"&FSellCash+optaddprice&"</tns:unitRetail>"				'단품정보 - 판매가
						strRst = strRst &"		<tns:unitCost>"&getCJmallSuplyPrice(optaddprice)&"</tns:unitCost>"		'단품정보 - 매입원가
						strRst = strRst &"		<tns:availableQty>"&itemSu&"</tns:availableQty>"						'단품정보 - 공급가능수량 (상품 재고 파악이 안되는경우는 999같은 숫자를 넣습니다.)
						If FtenCateLarge = "040" Then
							strRst = strRst &"		<tns:leadTime>02</tns:leadTime>"									'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
						Else
							If getzCostomMadeInd = "Y" Then
								strRst = strRst &"		<tns:leadTime>"&getzLeadTime()&"</tns:leadTime>"					'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
							Else
								strRst = strRst &"		<tns:leadTime>03</tns:leadTime>"									'단품정보 - 리드타임 (* 출고리드타임 예외등록 해야 함, 리드타임 기준 협의요함 00 : 오늘배송, 01 : 당일출고, 02 : 익일출고, 03 : 2일후출고, 04 : 4일, 05 : 5일, 06 : 6일.....)
							End If
						End If
						strRst = strRst &"		<tns:unitApplyRsn>"&fixday&"</tns:unitApplyRsn>"						'단품정보 - 적용사유 (10 : 적용안함, 20 : 상품포장, 30 : 상품생산, 40 : 입고검사, 50 : 출고검사, 60 : 설치상품)
						strRst = strRst &"		<tns:startSaleDt>"&FormatDate(now(), "0000-00-00")&"</tns:startSaleDt>"	'단품정보 - 판매시작일자
						strRst = strRst &"		<tns:endSaleDt>9999-12-30</tns:endSaleDt>"								'단품정보 - 판매종료일자 (판매상태수정에서..)
						strRst = strRst &"		<tns:vpn>"&rsget("itemid")&"_"&FItemOption&"</tns:vpn>"				'단품정보 - 협력사상품코드
						strRst = strRst &"	</tns:unit>"
					End If
				End If
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getCJOptionParamToEdit = strRst
	End Function

	'// 품절여부
	Public Function IsSoldOut()
		ISsoldOut = (FSellyn <> "Y") or ((FLimitYn = "Y") and (FLimitNo - FLimitSold < 1))
	End Function

	'// CJMALL 판매여부 반환
	Public Function getCjmallSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getCjmallSellYn = "Y"
			Else
				getCjmallSellYn = "N"
			End If
		Else
			getCjmallSellYn = "N"
		End If
	End Function

	Public Function getMdCode()
		Dim strRst
		If Fitemid = "899506" Then
			strRst = strRst &"	<tns:mdCode>5066</tns:mdCode>"
		Else
			If FtenCateLarge = "035" and FtenCateMid = "021" Then		'여행/취미 > 애완용품이라면..
				strRst = strRst &"	<tns:mdCode>5178</tns:mdCode>"
			Else
				strRst = strRst &"	<tns:mdCode>"&MD_CODE&"</tns:mdCode>"										'!!!MD코드	(있는 샘플도 있고, 누락된 샘플도 있음) 현아씨 문의 (텐바이텐 으로 가능)
			End If
		End If
		getMdCode = strRst
	End Function

	'상품 등록 XML
	Public Function getCjmallItemRegXML
		Dim strRst
		Dim ioriginCode, ioriginname
		Dim makercompCode, makercompName
		Dim certInfoParam
		certInfoParam = getCjCertParamToNewReg()

		ioriginCode 	= getOriginName2Code(Fsourcearea, ioriginname) 		'원산지코드
		makercompCode	= getmakerName2Code(Fsocname_kor, makercompName)	'제조사코드
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_01' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_01.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"									'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"					'!!!인증키
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"												'!!!상품분류체계 - 가등록채널구분(30:인터넷, 40:카탈로그)
		strRst = strRst &"	<tns:tGrpCd>"&FCddKey&"</tns:tGrpCd>"										'!!!상품분류체계 - 상품분류
		strRst = strRst &"	<tns:uniqBrandCd>"&CUNIQBRANDCD&"</tns:uniqBrandCd>"						'!!!상품분류체계 - 브랜드(텐바이텐:24049000)
		strRst = strRst &"	<tns:giftInd>Y</tns:giftInd>"											    '!!!상품분류체계 - 상품구분 (Y=일반판매상품, N=사은품)
		strRst = strRst &"	<tns:uniqMkrNatCd>"&ioriginCode&"</tns:uniqMkrNatCd>"						'!!!상품분류체계 - 제조국
		strRst = strRst &"	<tns:uniqMkrCompCd>"&makercompCode&"</tns:uniqMkrCompCd>"					'!!!상품분류체계 - 제조사
'		strRst = strRst &"	<tns:ingredient></tns:ingredient>"											'상품분류체계 - 주원료명	(샘플 페이지에는 누락)
'		strRst = strRst &"	<tns:zingredientOrigin></tns:zingredientOrigin>"							'상품분류체계 - 원료원산지	(샘플 페이지에는 누락) // 상품분류(대분류)가 식품일때만 원산지 필수(라던데..;;)
		strRst = strRst & getMdCode()
		If FoptionCnt = 0 Then
			strRst = strRst &"	<tns:itemDesc><![CDATA["&DdotFormat(chrbyte(getItemNameFormat,40,""),20)&"]]></tns:itemDesc>"			'!!!기본정보 - 상품명(120자 제약) (샘플에 CDATA없던거 추가)
		Else
			strRst = strRst &"	<tns:itemDesc><![CDATA["&DDotFormat(getItemNameFormat, 100)&"]]></tns:itemDesc>"			'!!!기본정보 - 상품명(120자 제약) (샘플에 CDATA없던거 추가)
		End If
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"	'!!!기본정보 - 운송장명(40자 제약)
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"		'!!!기본정보 - SMS상품명(20자 제약)
		strRst = strRst &"	<tns:vatCode>"&CHKIIF(FVatInclude="N","E","S")&"</tns:vatCode>"			 	'!!!기본정보 - 과세형태 (S:과세, E:면세, N:비과세, Z:영세)
		strRst = strRst &"	<tns:zDeliveryType>20</tns:zDeliveryType>"									'!!!기본정보 - 배송구분 (10:센터배송, 20:협력사배송, 30:직택배, 35:직택배Ⅱ, 40:직송, 99:배송없음)
		strRst = strRst &"	<tns:zShippingMethod>"&getdeliverfixday&"</tns:zShippingMethod>"			'!!!기본정보 - 배송유형 (10:택배배송, 20:설치상품, 30:배달서비스, 40:우편/등기배송) ''화물배송 확인
		strRst = strRst &"	<tns:courier>15</tns:courier>"												'!!!기본정보 - 택배사 (메인택배사 하나 지정 후 고정값 등록)(11:현대택배, 12:대한통운, 15:한진택배, 22:CJGLS, 29:CJHTH, 87:동부익스프레스) CJ택배 코드로 등록
		strRst = strRst &"	<tns:deliveryHomeCost>3000</tns:deliveryHomeCost>"							'기본정보 - 배송비 (배송구분이 협력사배송, 직송일 경우 필수 입력)
		strRst = strRst &"	<tns:zreturnNotReqInd>10</tns:zreturnNotReqInd>"							'기본정보 - 회수구분 (배송구분에 따라 필수/옵션)
'		strRst = strRst &"	<tns:zJointPackingQty></tns:zJointPackingQty>"								'기본정보 - 합포장단위 (배송구분에 따라 필수/옵션) (샘플페이지에는 누락)
		strRst = strRst &"	<tns:zCostomMadeInd>"&getzCostomMadeInd()&"</tns:zCostomMadeInd>"			'!!!기본정보 - 주문제작여부 (Y=주문제작, N=주문제작안함)) ''' 주문제작상품, 주문후제작상품 =>Y
		strRst = strRst &"	<tns:stockMgntLevel>2</tns:stockMgntLevel>"									'기본정보 - 재고관리레벨 (1=판매코드,2=단품코드)
'		strRst = strRst &"	<tns:leadtime></tns:leadtime>"												'기본정보 - 리드타임 (1. 프라자는 NULL셋팅 2.재고관리레벨이 "판매코드"일때 필수) (샘플페이지에는 누락)
'		strRst = strRst &"	<tns:leadtimeChgRsn></tns:leadtimeChgRsn>"									'기본정보 - 적용사유 (1. 프라자는 NULL셋팅 2.재고관리레벨이 "판매코드"일때 필수) (샘플페이지에는 누락)
		strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!기본정보 - 유료배송여부 (Y=유료배송,N=무료배송)        '' 확인.
		strRst = strRst &"	<tns:delayShipRewardIind>N</tns:delayShipRewardIind>"						'기본정보 - 지연보상여부 (Y=지연보상,N=지연보상안함)
'		strRst = strRst &"	<tns:packingMethod></tns:packingMethod>"									'기본정보 - 입고형태 (센터배송인 경우만 입력)
'		strRst = strRst &"	<tns:zOrderMaxQty>"&getOrderMaxNum&"</tns:zOrderMaxQty>"					'기본정보 - 1회최대주문수량 (고객당 1회 최대 주문가능 수량. 미입력시 제한없음
'		strRst = strRst &"	<tns:zDayOrderMaxQty></tns:zDayOrderMaxQty>"								'기본정보 - 1일최대주문수량 (고객당 일일 최대 주문가능 수량. 미입력시 제한없음)
		strRst = strRst &"	<tns:reserveDayInd>Y</tns:reserveDayInd>"									'기본정보 - 예약배송방식 (* 디폴트: YN-주문즉시 출하지시 Y-최초공급가능일 출하지시_Default)
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","10002")&"</tns:zContactSeqNo>"		'기본정보 - 협력사담당자
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zSupShipSeqNo>"		'기본정보 - 출하지
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zReturnSeqNo>"			'기본정보 - 회수지
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsSupShipSeqNo>"	'기본정보 - AS출하지
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsReturnSeqNo>"		'기본정보 - AS회수지
		If certInfoParam <> "" Then
			strRst = strRst & "<tns:certItemRequireYn>Y</tns:certItemRequireYn>"
		Else
			strRst = strRst & "<tns:certItemRequireYn>N</tns:certItemRequireYn>"
		End If
		strRst = strRst & "<tns:delivCostCd>00264854</tns:delivCostCd>"		'배송비 코드..	'이하 3줄 2016-04-14일 이후 추가 예정이라함, 00091063 -> 00264854로 수정 (2016-09-21 김진영)
		strRst = strRst & "<tns:delivCostType>01</tns:delivCostType>"	'배송비 속성코드 | 01 : 일반배송, 02 : 배송없음, 03 : 바로사용, 04 : 착불
		strRst = strRst & "<tns:fastDelivYn>0</tns:fastDelivYn>"	'빠른배송여부 | 0 : 빠른배송 불가(Default), 1 : 빠른 배송가능
		strRst = strRst & "<tns:harmGrd>"&Chkiif(IsAdultItem() = "Y", "19", "")&"</tns:harmGrd>"	'유해등급 | 유해등급:19, 없는경우 생략
		strRst = strRst & getCJOptionParamToReg															'단품정보
		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"텐바이텐 " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall상품정보 - CJmall상품명 , 텐바이텐 브랜드명 추가
'		strRst = strRst &"		<tns:keyword><![CDATA["&"텐바이텐;"&replace(Fkeywords,",",";")&"]]></tns:keyword>"						'!!!CJmall상품정보 - 검색키워드
		strRst = strRst &"		<tns:keyword><![CDATA["&getItemKeyword&"]]></tns:keyword>"												'!!!CJmall상품정보 - 검색키워드
		strRst = strRst & getCjCateParamToReg															'!!!메인카테고리여부(Y=카테고리,N=카테고리아님) // CJmall카테고리(세)
		strRst = strRst &"	</tns:mallitem>"
		strRst = strRst & certInfoParam															'품질인증정보
		strRst = strRst & getCjmallItemInfoCdToReg()													'상품기술서
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
														'daebeak	대백상품추가정보 빠져있음
		strRst = strRst & getCJAddImageParamToReg		'!!!이미지정보
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getCjmallItemRegXML = strRst
	End Function

	'정보 수정 XML
	Public Function getcjmallItemModXML()
		Dim strRst
		Dim ioriginCode, ioriginname
		Dim certInfoParam
		certInfoParam = getCjCertParamToNewReg()
		ioriginCode = getOriginName2Code(Fsourcearea, ioriginname)
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_02"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_02.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"												'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"								'!!!인증키
		strRst = strRst &"<tns:good>"
		strRst = strRst &"	<tns:sItem>"&FcjmallPrdNo&"</tns:sItem>"												'!!!판매상품코드(홈쇼핑)
	If Fitemid = "899506" Then
		strRst = strRst &"	<tns:loc>110</tns:loc>"																	'!!!상품분류체계 - 등록채널구분(공동구매)
	Else
		strRst = strRst &"	<tns:loc>30</tns:loc>"																	'!!!상품분류체계 - 등록채널구분(store포맷)
	End If
		strRst = strRst &"	<tns:zLocalBolDesc><![CDATA["&DDotFormat(getItemNameFormat, 10)&"]]></tns:zLocalBolDesc>"		'!!!기본정보 - 운송장명
		strRst = strRst &"	<tns:zlocalCcDesc><![CDATA["&DDotFormat(getItemNameFormat, 5)&"]]></tns:zlocalCcDesc>"			'!!!기본정보 - SMS상품명
		strRst = strRst &"	<tns:zContactSeqNo>"&chkiif(application("Svr_Info")="Dev","10002","10002")&"</tns:zContactSeqNo>"		'!!!기본정보 - 협력사담당자
		strRst = strRst &"	<tns:zSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zSupShipSeqNo>"		'!!!기본정보 - 출하지
		strRst = strRst &"	<tns:zReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zReturnSeqNo>"			'!!!기본정보 - 회수지
		strRst = strRst &"	<tns:zAsSupShipSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsSupShipSeqNo>"	'!!!기본정보 - AS출하지
		strRst = strRst &"	<tns:zAsReturnSeqNo>"&chkiif(application("Svr_Info")="Dev","23125","23125")&"</tns:zAsReturnSeqNo>"		'!!!기본정보 - AS회수지
        strRst = strRst &"	<tns:lowpriceInd>"&CHKIIF(IsCjFreeBeasong=False,"Y","N")&"</tns:lowpriceInd>"	'!!!기본정보 - 유료배송여부 (Y=유료배송,N=무료배송)        '' 확인.
		If certInfoParam <> "" Then
			strRst = strRst & "<tns:certItemRequireYn>Y</tns:certItemRequireYn>"
		Else
			strRst = strRst & "<tns:certItemRequireYn>N</tns:certItemRequireYn>"
		End If
		strRst = strRst & "<tns:delivCostCd>00264854</tns:delivCostCd>"		'배송비 코드..	'이하 3줄 2016-04-14일 이후 추가 예정이라함, 00091063 -> 00264854로 수정 (2016-09-21 김진영)
		strRst = strRst & "<tns:delivCostType>01</tns:delivCostType>"	'배송비 속성코드 | 01 : 일반배송, 02 : 배송없음, 03 : 바로사용, 04 : 착불
		strRst = strRst & "<tns:fastDelivYn>0</tns:fastDelivYn>"	'빠른배송여부 | 0 : 빠른배송 불가(Default), 1 : 빠른 배송가능
		strRst = strRst & "<tns:harmGrd>"&Chkiif(IsAdultItem() = "Y", "19", "")&"</tns:harmGrd>"	'유해등급 | 유해등급:19, 없는경우 생략
		strRst = strRst & getCJOptionParamToEdit                                                                      '' 확인해 볼것 ''864806
		strRst = strRst &"	<tns:mallitem>"
		strRst = strRst &"		<tns:mallItemDesc><![CDATA["&"텐바이텐 " & Fsocname_kor & " "&DDotFormat(getItemNameFormat, 186)&"]]></tns:mallItemDesc>"	'!!!CJmall상품정보 - CJmall상품명
		strRst = strRst &"	</tns:mallitem>"
		strRst = strRst & certInfoParam															'품질인증정보
		strRst = strRst & getCjmallItemInfoCdToReg()													'상품기술서
		strRst = strRst &"	<tns:goodsReport>"
		strRst = strRst &"		<tns:pedfId>91059</tns:pedfId>"
		strRst = strRst &"		<tns:html>"
		strRst = strRst &"			<![CDATA["&getCJItemContParamToReg&"]]>"
		strRst = strRst &"		</tns:html>"
		strRst = strRst &"	</tns:goodsReport>"
		strRst = strRst & getCJAddImageParamToReg		'!!!이미지정보
		strRst = strRst &"</tns:good>"
		strRst = strRst &"</tns:ifRequest>"
		getcjmallItemModXML = strRst
	End Function

	Public Function getCJMallPriceParameter
		Dim strRst, sqlStr, arrrows, chkOption, i, optAddPRcExists, GetTenTenMargin, specialPrice, tmpPrice, vBigPrice, vSmallPrice, ownItemCnt
		optAddPRcExists = false

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
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < FOutmallstandardMargin Then
				MustPrice = Forgprice
			Else
				If (FSellCash < Round(Fcjmallprice * 0.45, 0)) Then
					MustPrice = CStr(GetRaiseValue(Round(Fcjmallprice * 0.45, 0)/10)*10)
				Else
					MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
				End If
			End If
		End If

		sqlStr = ""
		' sqlStr = sqlStr & " SELECT distinct o.itemid, o.optAddPrice,  ro.outmallOptCode, o.itemoption"
		' sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option o "
		' sqlStr = sqlStr & " JOIN [db_item].[dbo].tbl_OutMall_regedoption ro on o.itemid=ro.itemid and ro.mallid ='"&CMALLNAME&"' and ro.itemoption = o.itemoption "
		' sqlStr = sqlStr & " WHERE o.itemid = '"&Fitemid&"' "
		' sqlStr = sqlStr & " GROUP BY o.itemid, o.optAddPrice, ro.outmallOptCode, o.itemoption"
		' sqlStr = sqlStr & " ORDER BY o.optAddPrice, o.itemoption"
		sqlStr = sqlStr & " SELECT distinct ro.itemid, isnull(o.optAddPrice, 0) as optAddPrice,  ro.outmallOptCode, ro.itemoption"
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option o "
		sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_OutMall_regedoption ro on o.itemid=ro.itemid  and ro.itemoption = o.itemoption "
		sqlStr = sqlStr & " WHERE o.itemid = '"&Fitemid&"' and ro.mallid ='"&CMALLNAME&"' "
		sqlStr = sqlStr & " GROUP BY ro.itemid, isnull(o.optAddPrice, 0),  ro.outmallOptCode, ro.itemoption"
		sqlStr = sqlStr & " ORDER BY isnull(o.optAddPrice, 0), ro.itemoption"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			arrrows = rsget.getRows
			chkOption = True
		Else
			chkOption = False
		End If
		rsget.close

		if (chkOption) then
			For i = 0 To UBound(ArrRows,2)
				optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
			Next
		end if

		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8"" ?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_04.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"						'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"		'!!!인증키

		'2015-12-31 14:27 김진영 이하 IF문으로 교체 하단(893Line If문 주석)
		If chkOption = True Then
		strRst = strRst &"<tns:itemPrices>"
		strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
		strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
		strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
		strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
		strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
		strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
		strRst = strRst &"</tns:itemPrices>"
			For i = 0 To UBound(ArrRows,2)
				strRst = strRst &"<tns:itemPrices>"
				strRst = strRst &"	<tns:typeCD>02</tns:typeCD>"						'01이면 판매코드 / 02면 단품코드
				strRst = strRst &"	<tns:itemCD_ZIP>"&arrRows(2,i)&"</tns:itemCD_ZIP>"
				strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
				strRst = strRst &"	<tns:newUnitRetail>"&MustPrice+arrRows(1,i)&"</tns:newUnitRetail>"
				strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice(arrRows(1,i))&"</tns:newUnitCost>"
				strRst = strRst &"</tns:itemPrices>"
			Next
		Else
			If (Not optAddPRcExists) OR (chkOption = False) Then
				strRst = strRst &"<tns:itemPrices>"
				strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
				strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
				strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
				strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
				strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
				strRst = strRst &"</tns:itemPrices>"
			End If
		End If

'		If (Not optAddPRcExists) OR (chkOption = False) Then
'			strRst = strRst &"<tns:itemPrices>"
'			strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
'			strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
'			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
'			strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
'			strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
'			strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
'			strRst = strRst &"</tns:itemPrices>"
'		Else
'			If chkOption = True Then
'			strRst = strRst &"<tns:itemPrices>"
'			strRst = strRst &"	<tns:typeCD>01</tns:typeCD>"								'01이면 판매코드 / 02면 단품코드
'			strRst = strRst &"	<tns:itemCD_ZIP>"&FcjmallPrdNo&"</tns:itemCD_ZIP>"
'			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
'			strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
'			strRst = strRst &"	<tns:newUnitRetail>"&MustPrice&"</tns:newUnitRetail>"
'			strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice2&"</tns:newUnitCost>"
'			strRst = strRst &"</tns:itemPrices>"
'				For i = 0 To UBound(ArrRows,2)
'					strRst = strRst &"<tns:itemPrices>"
'					strRst = strRst &"	<tns:typeCD>02</tns:typeCD>"						'01이면 판매코드 / 02면 단품코드
'					strRst = strRst &"	<tns:itemCD_ZIP>"&arrRows(2,i)&"</tns:itemCD_ZIP>"
'					strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
'					strRst = strRst &"	<tns:effectiveDate>"&FormatDate(now(), "0000-00-00")&"</tns:effectiveDate>"
'					strRst = strRst &"	<tns:newUnitRetail>"&MustPrice+arrRows(1,i)&"</tns:newUnitRetail>"
'					strRst = strRst &"	<tns:newUnitCost>"&getCJmallSuplyPrice(arrRows(1,i))&"</tns:newUnitCost>"
'					strRst = strRst &"</tns:itemPrices>"
'					optAddPRcExists = optAddPRcExists or (arrRows(1,i)>0)
'				Next
'			End If
'		End If
		strRst = strRst &"</tns:ifRequest>"
		getCJMallPriceParameter = strRst
	End Function

	'단품 수량 수정 XML
	Public Function getCJMallQTYParameter
		Dim sqlStr, oneOpt, j
		Dim arrRows, i, strRst, validSellno
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_05.xsd"">"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키

		sqlStr = ""
		sqlStr = sqlStr & " select isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName "
		sqlStr = sqlStr & " from [db_item].[dbo].tbl_OutMall_regedoption as r "
		sqlStr = sqlStr & " left join [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
		sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			oneOpt = rsget.getRows
		End If
		rsget.close

		If (UBound(oneOpt ,2) = "0") and (oneOpt(2,0) = "단일상품") Then
			strRst = strRst &"<tns:ltSupplyPlans>"
			strRst = strRst &"	<tns:unitCd>"&oneOpt(1,0)&"</tns:unitCd>"
			strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
			If GetCJLmtQty = 0 Then
				strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
			Else
				strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
			End If
			strRst = strRst &"	<tns:availSupQty>"&chkiif(GetCJLmtQty=0,"1",GetCJLmtQty)&"</tns:availSupQty>"
			strRst = strRst &"</tns:ltSupplyPlans>"
		Else
			sqlStr = ""
			sqlStr = sqlStr & " SELECT o.itemoption, o.optionTypeName, o.optionname, isnull(R.outmallOptCode, '') as outmallOptCode, (o.optlimitno-o.optlimitsold) as optLimit, o.optlimityn, o.isUsing, o.optsellyn " & VBCRLF
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option o " & VBCRLF
			sqlStr = sqlStr & " left join [db_item].[dbo].tbl_OutMall_regedoption R on o.itemid=R.itemid and o.itemoption=R.itemoption and R.mallid='"&CMALLNAME&"' " & VBCRLF
			sqlStr = sqlStr & " where R.outmallOptCode <> '' and o.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If isArray(arrRows) Then
				For i = 0 To UBound(ArrRows,2)
					validSellno = 999				'최대 999로 강제지정
					If (FSellyn <> "Y") or ((arrRows(5,i) = "Y") and (arrRows(4,i) < 1)) or (arrRows(6,i) <> "Y") or (arrRows(7,i) <> "Y") Then
						validSellno = 0
					End If

					If (arrRows(5,i) = "Y") Then
						validSellno = arrRows(4,i)
					End If

					If (validSellno < CMAXLIMITSELL) Then validSellno = 0
					If (arrRows(5,i) = "Y") and (validSellno > 0) Then
						validSellno = validSellno - CMAXLIMITSELL
					End If
					If (validSellno < 1) then validSellno = 0
					If IsSoldOut Then validSellno = 0

					strRst = strRst &"<tns:ltSupplyPlans>"
					strRst = strRst &"	<tns:unitCd>"&arrRows(3,i)&"</tns:unitCd>"
					strRst = strRst &"	<tns:chnCls>30</tns:chnCls>"
					strRst = strRst &"	<tns:strDt>"&FormatDate(now(), "0000-00-00")&"</tns:strDt>"
					If validSellno = 0 Then
						strRst = strRst &"	<tns:endDt>"&FormatDate(now(), "0000-00-00")&"</tns:endDt>"
					Else
						strRst = strRst &"	<tns:endDt>9999-12-30</tns:endDt>"
					End If
					strRst = strRst &"	<tns:availSupQty>"&chkiif(validSellno=0,"1",validSellno)&"</tns:availSupQty>"
					strRst = strRst &"</tns:ltSupplyPlans>"
				Next
			End If
		End If
		strRst = strRst &"</tns:ifRequest>"
		getCJMallQTYParameter = strRst
	End Function

	Public Function getcjmallOptSellModParameter
		Dim sqlStr, arrRows, i
		Dim itemoption, optiontypename, optionname, optLimit, optlimityn, isUsing, optsellyn, preged, optNameDiff, forceExpired, oopt, ooptCd, YtoN, NtoY, DelOpt
		Dim validSellno, strRst
		strRst = ""
		strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst &"<tns:ifRequest xmlns:tns='http://www.example.org/ifpa' tns:ifId='IF_03_03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:schemaLocation='http://www.example.org/ifpa ../IF_03_03.xsd'>"
		strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"					'!!!협력업체코드
		strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"	'!!!인증키

		sqlStr = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_cjmall 'cjmall'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open sqlStr, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		If FMaySoldOut = "Y" Then
			strRst = strRst &"<tns:itemStates>"
			strRst = strRst &"<tns:typeCd>01</tns:typeCd>"						'01=판매코드,02=단품코드
			strRst = strRst &"<tns:itemCd_zip>"&Fcjmallprdno&"</tns:itemCd_zip>"
			strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A-진행, I-일시중단
			strRst = strRst &"</tns:itemStates>"
		ElseIf FMaySoldOut = "N" Then
			strRst = strRst &"<tns:itemStates>"
			strRst = strRst &"<tns:typeCd>01</tns:typeCd>"						'01=판매코드,02=단품코드
			strRst = strRst &"<tns:itemCd_zip>"&Fcjmallprdno&"</tns:itemCd_zip>"
			strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
			strRst = strRst &"<tns:packInd>A</tns:packInd>"						'A-진행, I-일시중단
			strRst = strRst &"</tns:itemStates>"
		End If

		For i = 0 To UBound(ArrRows,2)
			itemoption		= ArrRows(1,i)
			optiontypename	= ArrRows(2,i)
			optionname		= Replace(Replace(db2Html(ArrRows(3,i)),":",""),",","")
			optLimit		= ArrRows(4,i)
			optlimityn		= ArrRows(5,i)
			isUsing			= ArrRows(6,i)
			optsellyn		= ArrRows(7,i)
			preged			= (ArrRows(11,i)=1)
			optNameDiff		= (ArrRows(12,i)=1)
			forceExpired	= (ArrRows(13,i)=1)
			oopt			= ArrRows(14,i)
			ooptCd			= ArrRows(15,i)
			YtoN			= (ArrRows(16,i)=1)
			NtoY			= (ArrRows(17,i)=1)
			DelOpt			= (ArrRows(18,i)=1)
			If FMaySoldOut = "Y" Then
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
			ElseIf (forceExpired) or (optNameDiff) or (DelOpt) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5) Then			'한정이고 수량이 5개 이하인 경우 // (isUsing="N") or (optsellyn="N") or 추가 2013/05/31..''2013-12-04 13:30 김진영..optLimit < 5를 optLimit <= 5로 수정
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>I</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
		    ElseIf (preged) and (ooptCd <> "") Then
				strRst = strRst &"<tns:itemStates>"
				strRst = strRst &"<tns:typeCd>02</tns:typeCd>"						'01=판매코드,02=단품코드
				strRst = strRst &"<tns:itemCd_zip>"&ooptCd&"</tns:itemCd_zip>"
				strRst = strRst &"<tns:chnCls>30</tns:chnCls>"
				strRst = strRst &"<tns:packInd>A</tns:packInd>"						'A:진행, I:일시중단
				strRst = strRst &"</tns:itemStates>"
			End If
		Next
		strRst = strRst &"</tns:ifRequest>"
		getcjmallOptSellModParameter = strRst
	End Function
End Class

Class CCJMall
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

	'// 미등록 상품 목록(등록용)
	Public Sub getCJMallNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			''' 옵션 추가금액 있는경우 등록 불가. //옵션 전체 품절인 경우 등록 불가.
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
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyDiv, c.safetyNum "
		strSql = strSql & "	, '"&CitemGbnKey&"' as itemGbnKey "
		strSql = strSql & "	, isNULL(R.cjmallStatCD,-9) as cjmallStatCD "
		strSql = strSql & "	, UC.socname_kor, isnull(PD.itemtypeCd, '') as cddkey, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_cjmall_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_cjmall_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_cjmall_MngDiv_mapping as PD on PD.tencatelarge = i.cate_large and PD.tencatemid = i.cate_mid and PD.tencatesmall = i.cate_small "
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
		strSql = strSql & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & " and rtrim(ltrim(isNull(i.deliverfixday, ''))) = '' "		'택배(일반)
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100)>=" & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &") ) "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_item.dbo.tbl_cjmall_regItem WHERE cjmallStatCD >= 3) "	''등록완료이상은 등록안됨.										'롯데등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCJMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FitemGbnKey        = rsget("itemGbnKey")
				FOneItem.FcjmallStatCD		= rsget("cjmallStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.Fcddkey			= rsget("cddkey")
				FOneItem.FsafetyDiv  	  	= rsget("safetyDiv")
				FOneItem.FsafetyNum    		= rsget("safetyNum")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub getCJMallNotEditOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyDiv, c.safetyNum "
		strSql = strSql & "	, m.cjmallPrdNo, m.cjmallprice, m.cjmallSellYn, m.accFailCnt, m.lastErrStr, UC.socname_kor, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_cjmall_regItem as m on i.itemid = m.itemid "
		' strSql = strSql & " LEFT JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt FROM db_etcmall.dbo.tbl_cjmall_cate_mapping GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as cm "
		' strSql = strSql & " 	on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		' strSql = strSql & " LEFT JOIN  (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as PmapCnt FROM db_etcmall.dbo.tbl_cjmall_Prddiv_mapping  GROUP BY tenCateLarge, tenCateMid, tenCateSmall ) as Pm "
		' strSql = strSql & " 	on Pm.tenCateLarge = i.cate_large and Pm.tenCateMid = i.cate_mid and Pm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1 " & addSql
		strSql = strSql & " and m.cjmallPrdNo is Not Null and m.cjmallStatCD = 7 "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCJMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FcjmallPrdNo		= rsget("cjmallPrdNo")
				FOneItem.Fcjmallprice		= rsget("cjmallprice")
				FOneItem.FcjmallSellYn		= rsget("cjmallSellYn")
                FOneItem.Fvatinclude        = rsget("vatinclude")
                FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FaccFailCnt    	= rsget("accFailCnt")
				FOneItem.FlastErrStr    	= rsget("lastErrStr")
				FOneItem.FsafetyDiv  	  	= rsget("safetyDiv")
				FOneItem.FsafetyNum    		= rsget("safetyNum")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub getCJMallNotRegEditOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT top " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyDiv, c.safetyNum "
		strSql = strSql & "	, m.cjmallPrdNo, m.cjmallprice, m.cjmallSellYn, m.accFailCnt, m.lastErrStr, UC.socname_kor "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
'		strSql = strSql & "		or ((i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.deliveryType in ('7','6') "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<"&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_cjmall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " WHERE 1 = 1 " & addSql
		strSql = strSql & " and m.cjmallPrdNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CCJMallItem
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
				FOneItem.Fsourcearea		= db2html(rsget("sourcearea"))
				FOneItem.Fmakername			= db2html(rsget("makername"))
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FcjmallPrdNo		= rsget("cjmallPrdNo")
				FOneItem.Fcjmallprice		= rsget("cjmallprice")
				FOneItem.FcjmallSellYn		= rsget("cjmallSellYn")
                FOneItem.Fvatinclude        = rsget("vatinclude")
                FOneItem.Fsocname_kor		= rsget("socname_kor")
				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FaccFailCnt    	= rsget("accFailCnt")
				FOneItem.FlastErrStr    	= rsget("lastErrStr")
				FOneItem.FsafetyDiv  	  	= rsget("safetyDiv")
				FOneItem.FsafetyNum    		= rsget("safetyNum")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

End Class

Function getCjmallPrdNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 cjmallPrdNo FROM db_item.dbo.tbl_cjmall_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getCjmallPrdNo = rsget("cjmallPrdNo")
	rsget.Close
End Function

Function getCjMallfirstItemoption(byval iitemid)
    dim ret
    dim sqlStr, iOption

    if iitemid="" then Exit function

    sqlStr = " select top 1 itemoption from db_item.dbo.tbl_OutMall_regedoption"
    sqlStr = sqlStr & " where mallid='"&CMALLNAME&"'"
    sqlStr = sqlStr & " and itemid="&iitemid
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ret = rsget("itemoption")
	End If
	rsget.close

	if (ret="") then
		sqlStr = "select top 1 itemoption from db_item.dbo.tbl_item_option where itemid = '"&iitemid&"' and isusing = 'Y' and optsellyn = 'Y' order by itemoption asc"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ret = rsget("itemoption")
		Else
			ret = "0000"
		End If
		rsget.close
	end if
	getCjMallfirstItemoption = ret
End Function

Function getOriginName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 areacode, areaName" & VBCRLF
	sqlStr = sqlStr & " FROM db_temp.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
	sqlStr = sqlStr & " WHERE areaName='"&iname&"'"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		retVal = rsget("areacode")
		ioriginName = rsget("areaName")
	end if
	rsget.Close

	If (retVal = "") Then
		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 areacode, areaName FROM db_temp.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
		sqlStr = sqlStr & " WHERE CharIndex('"&iname&"',areaName) > 0" & VBCRLF
		sqlStr = sqlStr & " or CharIndex(areaName,'"&iname&"') > 0" & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.Eof) then
			retVal = rsget("areacode")
			ioriginName = rsget("areaName")
		End If
		rsget.Close
	End If

	If (retVal = "") Then
		retVal="000"
		ioriginName = "없음"
	End If

	getOriginName2Code=retVal
	Exit Function
End Function

Function getmakerName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 code, makerName" & VBCRLF
	sqlStr = sqlStr & " FROM db_temp.dbo.tbl_cjmall_makerName" & VBCRLF
	sqlStr = sqlStr & " WHERE makerName='"&html2db(iname)&"'"
'rw sqlStr
'response.end
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		retVal = rsget("code")
		ioriginName = rsget("makerName")
	end if
	rsget.Close

	If (retVal = "") Then
		retVal="15210"
		ioriginName = "텐바이텐"
	End If

	getmakerName2Code = retVal
	Exit Function
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
%>
