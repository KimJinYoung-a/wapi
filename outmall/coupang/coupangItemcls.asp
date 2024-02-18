<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "coupang"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.

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
		sqlStr = sqlStr & " and outmallOptName <> '단일상품' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			regedCoupangOptCnt = rsget("cnt")
		rsget.Close

		If tenOptCnt > 0 AND regedCoupangOptCnt = 0 Then			'단품 -> 옵션
			IsAllOptionChange = "Y"
		ElseIf tenOptCnt = 0 AND regedCoupangOptCnt > 0 Then		'옵션 -> 단품
			IsAllOptionChange = "Y"
		Else
			IsAllOptionChange = "N"
		End If
	End Function

	Public Function getCoupangInfoDiv(infoDivName)
		Select Case infoDivName
			Case "의류"								getCoupangInfoDiv =  "01"
			Case "구두/신발"							getCoupangInfoDiv =  "02"
			Case "가방"								getCoupangInfoDiv =  "03"
			Case "패션잡화(모자/벨트/액세서리)"			getCoupangInfoDiv =  "04"
			Case "침구류/커튼"						getCoupangInfoDiv =  "05"
			Case "가구"								getCoupangInfoDiv =  "06"
			Case "영상가전(TV류)"						getCoupangInfoDiv =  "07"
			Case "가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지 등)"		getCoupangInfoDiv =  "08"
			Case "계절가전(에어컨/온풍기 등)"			getCoupangInfoDiv =  "09"
			Case "사무용기기(컴퓨터/노트북/프린터 등)"	getCoupangInfoDiv =  "10"
			Case "광학기기(디지털카메라/캠코더 등)"		getCoupangInfoDiv =  "11"
			Case "휴대폰"							getCoupangInfoDiv =  "13"
			Case "내비게이션"							getCoupangInfoDiv =  "14"
			Case "자동차용품(자동차부품/기타 자동차용품)"		getCoupangInfoDiv =  "15"
			Case "의료기기"							getCoupangInfoDiv =  "16"
			Case "주방용품"							getCoupangInfoDiv =  "17"
			Case "화장품"							getCoupangInfoDiv =  "18"
			Case "귀금속/보석/시계류"					getCoupangInfoDiv =  "19"
			Case "식품(농축수산물)"					getCoupangInfoDiv =  "20"
			Case "가공식품"							getCoupangInfoDiv =  "21"
			Case "건강기능식품"						getCoupangInfoDiv =  "22"
			Case "영유아용품"							getCoupangInfoDiv =  "23"
			Case "악기"								getCoupangInfoDiv =  "24"
			Case "스포츠용품"							getCoupangInfoDiv =  "25"
			Case "서적"								getCoupangInfoDiv =  "26"
			Case "물품대여 서비스(정수기, 비데, 공기청정기 등)"						getCoupangInfoDiv =  "31"
			Case "디지털 콘텐츠(음원, 게임, 인터넷강의 등)"							getCoupangInfoDiv =  "33"
			Case "기타 재화"							getCoupangInfoDiv =  "35"
		End Select

		If Instr(infoDivName, "소형전자(MP") > 0 Then
			getCoupangInfoDiv =  "12"
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

	Public Function getDeliverytypeName
		If (Fdeliverytype = "9") Then
			getDeliverytypeName = "<font color='blue'>[조건 "&FormatNumber(FdefaultfreeBeasongLimit,0)&"]</font>"
		ElseIf (Fdeliverytype = "7") then
			getDeliverytypeName = "<font color='red'>[업체착불]</font>"
		ElseIf (Fdeliverytype = "2") then
			getDeliverytypeName = "<font color='blue'>[업체]</font>"
		Else
			getDeliverytypeName = ""
		End If
	End Function

	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Function getLimitHtmlStr()
	    If IsNULL(FLimityn) Then Exit Function
	    If (FLimityn = "Y") Then
	        getLimitHtmlStr = "<br><font color=blue>한정:"&getLimitEa&"</font>"
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
			CASE -9 : getCoupangStatName = "미등록"
			CASE -1 : getCoupangStatName = "등록실패"
			CASE 0 : getCoupangStatName = "<font color=blue>등록예정</font>"
			CASE 1 : getCoupangStatName = "전송시도"
			CASE 2 : getCoupangStatName = "반려"
			CASE 3 : getCoupangStatName = "승인전"
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

	'// Coupang 판매여부 반환
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

			'2019-06-03 17:37 김진영 주석처리..품목이 안 맞아도 등록할 수 있게 sp 변경
			' addSql = addSql & " and c.infodiv in ( "
			' addSql = addSql & " SELECT  "
			' addSql = addSql & " 	CASE WHEN noticeCategoryName = '의류' THEN '01' "
			' addSql = addSql & "  	WHEN noticeCategoryName = '구두/신발' THEN '02' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '가방' THEN '03' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '패션잡화(모자/벨트/액세서리)' THEN '04' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '침구류/커튼' THEN '05' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '가구' THEN '06' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '영상가전(TV류)' THEN '07' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '가정용 전기제품(냉장고/세탁기/식기세척기/전자레인지 등)' THEN '08' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '계절가전(에어컨/온풍기 등)' THEN '09' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '사무용기기(컴퓨터/노트북/프린터 등)' THEN '10' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '광학기기(디지털카메라/캠코더 등)' THEN '11' "
			' addSql = addSql & " 	WHEN left(noticeCategoryName, 4) = '소형전자' THEN '12' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '휴대폰' THEN '13' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '내비게이션' THEN '14' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '자동차용품(자동차부품/기타 자동차용품)' THEN '15' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '의료기기' THEN '16' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '주방용품' THEN '17' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '화장품' THEN '18' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '귀금속/보석/시계류' THEN '19' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '가공식품' THEN '20' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '식품(농축수산물)' THEN '21' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '건강기능식품' THEN '22' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '영유아용품' THEN '23' "
			' addSql = addSql & "		WHEN noticeCategoryName = '악기' THEN '24' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '스포츠용품' THEN '25' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '서적' THEN '26' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '물품대여 서비스(정수기, 비데, 공기청정기 등)' THEN '31' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '디지털 콘텐츠(음원, 게임, 인터넷강의 등)' THEN '33' "
			' addSql = addSql & " 	WHEN noticeCategoryName = '기타 재화' THEN 35 END "
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
		strSql = strSql & " 	CASE WHEN CHARINDEX('라이터', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('가스', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('부탄가스', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('본드', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('니스', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('시너', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('레이저', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 		 WHEN CHARINDEX('전자담배', i.itemname) > 0 THEN 0 "
		strSql = strSql & " 	ELSE 1 END "
		strSql = strSql & " ) "
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
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		'strSql = strSql & "	and isnull(R.CoupangStatCD,0) < 3  "
		'strSql = strSql & " and R.itemid is NULL"  '' 2018/11/23
		strSql = strSql & " and isNull(R.CoupangGoodNo, '') = ''"  '' 2019/04/12 위 R.itemid is null 주석
		strSql = strSql & " and cm.mapCnt is Not Null "
		strSql = strSql & " and i.itemdiv not in ('06') "	''주문제작문구 상품 제외
		strSql = strSql & " and bm.outboundShippingPlaceCode is Not Null "
		strSql = strSql & "		"	& addSql											'카테고리 매칭 상품만
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
			'선택상품이 있다면
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
		strSql = strSql & " 	or i.itemdiv in ('06') "		''주문제작문구 상품 품절처리
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
		strSql = strSql & " and m.coupangGoodNo is Not Null "									'#등록 상품만
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

'// 상품이미지 존재여부 검사
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