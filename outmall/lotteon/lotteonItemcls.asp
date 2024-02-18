<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "lotteon"
CONST CUPJODLVVALID = TRUE			''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5				'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST APIATTRURL = "https://onpick-api.lotteon.com"
CONST CDEFALUT_STOCK = 99999

Class CLotteonItem
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
	Public Fsourcearea
	Public Fmakername
	Public FBrandName
	Public FBrandNameKor
	Public FItemsize
	Public FItemsource
	Public FUsingHTML
	Public FSafetyNum
	Public Fitemcontent
	Public FLotteonStatCD
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
	Public FLotteonGoodNo
	Public FLotteonprice
	Public FLotteonSellYn
	Public FStd_cat_id
	Public FDisp_cat_id

	Public FAdultType
	Public FLastStatCheckDate
	Public FOrderMaxNum

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999999" Then
			getOrderMaxNum = 999999
		End If
	End Function

	Public Function getRegedOptionCnt
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as Cnt  "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_OutMall_regedoption "
		sqlStr = sqlStr & " WHERE mallid= 'lotteon' "
		sqlStr = sqlStr & " and itemoption <> '0000' "
		sqlStr = sqlStr & " and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			getRegedOptionCnt = rsget("Cnt")
		rsget.Close
	End Function

	'// 품절여부
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
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, ownItemCnt
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

			If FLotteonPrice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					' If (FSellCash < Round(FSsgprice * 0.25, 0)) Then
					' 	MustPrice = CStr(GetRaiseValue(Round(FSsgprice * 0.25, 0)/10)*10)
					' Else
						MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					' End If
				End If
			End If
		End If
	End Function

	'// Lotteon 판매여부 반환
	Public Function getLotteonSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getLotteonSellYn = "Y"
			Else
				getLotteonSellYn = "N"
			End If
		Else
			getLotteonSellYn = "N"
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

	Public Function getShopLeadTime()
		Dim CateLargeMid, leadTime
		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "040010", "040011", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "040121", "055070", "055080", "055090", "055100", "055110", "055120", "055222"
				leadTime = 15
			Case "045001",  "045002", "045003", "045004", "045005", "045006", "045007", "045008", "045009", "045010", "045011", "045012"
			 	leadTime = 10
			Case "070010", "070020", "070030", "070040", "070050", "070070", "070110", "070120", "070140", "070150", "070160", "070200", "070201", "070202", "070203", "080007", "080010", "080020", "080030", "080031", "080040", "080050", "080051", "080060", "080070", "080071", "080080", "080090", "090005", "090010", "090011", "090020", "090030", "090040", "090050", "090060", "090061", "090070", "090071", "090080"
				leadTime = 7
			Case "050010", "050020", "050030", "050040", "050045", "050050", "050070", "050110", "050120", "050666", "050777", "060010", "060020", "060040", "060050", "060060", "060070", "060080", "060090", "060120", "060130", "060140", "060150", "060160", "100010", "100020", "100030", "100040", "100060", "100070", "100080", "100090", "100100", "100110", "100120", "100130", "100140", "100150", "100201", "100300"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select
		getShopLeadTime = leadTime
	End Function

    public function getItemNameFormat()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST상품] "&FItemName
		Else
			buf = "[텐바이텐] "&FItemName
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        buf = LeftB(buf, 130)
        getItemNameFormat = buf
    end function

	Public function getOriginCode()
		If Fsourcearea = "한국" OR Fsourcearea = "대한민국" Then
			getOriginCode = "KR"
		Else
			getOriginCode = "ETC"
		End If
	End Function


	Public function getBrandCode()
		Select Case Fmakerid
			Case "disney10x10"
				getBrandCode = "P778"
			Case "sanrio10x10"
				getBrandCode = "P47543"
			Case "universal10x10"
				getBrandCode = "P11805"
			Case "peanuts10x10"
				getBrandCode = "P5270"
			Case "sanx10x10"
				getBrandCode = "P15324"
			Case "cncglobalkr"
				getBrandCode = "P2399"
			Case Else
				getBrandCode = ""
		End Select
	End Function

    public function getItemNameFormat2()
        dim buf
		If application("Svr_Info") = "Dev" Then
			buf = "[TEST상품] "&FItemName
		Else
			buf = "[텐바이텐] "&FItemName
		End If
        buf = replace(buf,"'","")
        buf = replace(buf,"~","-")
        buf = replace(buf,"<","[")
        buf = replace(buf,">","]")
        buf = replace(buf,"%","프로")
        buf = replace(buf,"&","＆")
        buf = replace(buf,"[무료배송]","")
        buf = replace(buf,"[무료 배송]","")
        getItemNameFormat2 = buf
    end function

	Public Function getLotteonKeywordsParameter(obj)
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
			retKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			retKeyword = keyword1
		End If

		If retKeyword = "" Then
			Set obj("spdLst")(null)("scKwdLst") = jsArray()
				obj("spdLst")(null)("scKwdLst") = null
		Else
			commaSplit = Split(retKeyword,",")
			Set obj("spdLst")(null)("scKwdLst") = jsArray()								'검색키워드목록 | 5개 이하만 등록 가능
				For i = 0 To Ubound(commaSplit)
					obj("spdLst")(null)("scKwdLst")(i) = commaSplit(i)
				Next
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

	'상품정보제공고시
	Public Function getLotteonInfoCdParameter(obj)
		Dim strSql, buf, i
		Dim mallinfoCd, infoContent, mallinfodiv
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00002') THEN '상세페이지 참고' "
		strSql = strSql & "     WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' "
		strSql = strSql & "     WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
		strSql = strSql & " 	WHEN LEN(isNull(F.infocontent, '')) < 2 THEN '상세페이지 참고' "
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN ( "
		strSql = strSql & "  SELECT TOP 1 itemid, certNum FROM db_item.dbo.tbl_safetycert_tenReg where itemid = '"&FItemID&"' "
		strSql = strSql & " ) as tr on I.itemid = tr.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = '"& CMALLNAME &"' and IC.itemid='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			If CStr(rsget("mallinfodiv")) = "35"  Then
				mallinfodiv = "38"
			ElseIf CStr(rsget("mallinfodiv")) = "47"  Then
				mallinfodiv = "39"
			ElseIf CStr(rsget("mallinfodiv")) = "48"  Then
				mallinfodiv = "40"
			Else
				mallinfodiv = CStr(rsget("mallinfodiv"))
			End If

			Set obj("spdLst")(null)("pdItmsInfo") = jsObject()
				obj("spdLst")(null)("pdItmsInfo")("pdItmsCd") = mallinfodiv
				Set obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst") = jsArray()						'#상품품목항목목록
			i = 0
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
				End If

				Set obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i) = jsObject()
					obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i)("pdArtlCd") = mallinfoCd		'#상품항목코드
					obj("spdLst")(null)("pdItmsInfo")("pdItmsArtlLst")(i)("pdArtlCnts") = infoContent	'#상품항목내용 | 해당 고시정보항목의 항목값을 입력한다.

				i = i + 1
				rsget.MoveNext
			Loop
		End If
		rsget.Close
    End Function

	'안전인증목록 파라메터 생성
	Public Function getLotteonCertInfoParameter(obj)
		Dim strSql
		Dim safetyDiv, safetyId, certNum, certOrganName, isRegCert
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN 'ELC_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN 'ELC_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN 'ELC_SUPS' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN 'LIFE_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN 'LIFE_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN 'LIFE_SUPS' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN 'CHL_ATHN' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN 'CHL_CFM' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN 'CHL_SUPS' end as safetyId "
		strSql = strSql & " , t.certNum, f.certOrganName, f.makerName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			certNum			= rsget("certNum")
			certOrganName	= rsget("certOrganName")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			If safetyDiv = "30" OR safetyDiv = "60" OR safetyDiv = "90" Then
				certNum = ""
			End If

			Set obj("spdLst")(null)("sftyAthnLst") = jsArray()
				Set obj("spdLst")(null)("sftyAthnLst")(0) = jsObject()
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnTypCd") = safetyId		'#안전인증유형코드 [공통코드 : SFTY_ATHN_TYP_CD] | CHL_SUPS : [어린이제품]공급자적합성확인, CHL_ATHN : [어린이제품]안전인증, CHL_CFM : [어린이제품]안전확인, CMCN_TNTT : [방송통신기자재]잠정인증, CMCN_REG : [방송통신기자재]적합등록, CMCN_ATHN : [방송통신기자재]적합인증, LIFE_SUPS : [생활용품]공급자적합성확인, LIFE_ATHN : [생활용품]안전인증, LIFE_CFM : [생활용품]안전확인, ELC_SUPS : [전기용품]공급자적합성확인, ELC_ATHN : [전기용품]안전인증, ELC_CFM : [전기용품]안전확인, LIFE_STD : [생활용품]안전기준준수, CHEM_LIFE : [화학제품] 생활화학제품 안전기준적합확인신고번호 / 승인번호, CHEM_BIOC : [화학제품] 살생물제품 승인번호, ETC : 기타
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnOrgnNm") = certOrganName	'안전인증기관명
					obj("spdLst")(null)("sftyAthnLst")(0)("sftyAthnNo") = certNum			'안전인증번호
		Else
			Set obj("spdLst")(null)("sftyAthnLst") = jsArray()
				obj("spdLst")(null)("sftyAthnLst") = null
		End If
	End Function

	'표준카테고리속성목록
	Public Function getLotteonStdCateAttrParameter(obj)
		' Set obj("spdLst")(null)("scatAttrLst") = jsArray()
		' 	Set obj("spdLst")(null)("scatAttrLst")(null) = jsObject()
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optCd") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optValCd") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("optVal") = ""
		' 		obj("spdLst")(null)("scatAttrLst")(null)("dtlsVal") = ""
		Set obj("spdLst")(null)("scatAttrLst") = jsArray()
			obj("spdLst")(null)("scatAttrLst") = null
	End Function

	'상품설명 파라메터 생성
	Public Function getLotteonContParamToReg(obj)
		Dim strRst, strSQL, retContents, retOrderComment
		strRst = ("<div align=""center"">")
		strRst = strRst & ("<p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteon.jpg></p><br />")
		strRst = strRst & ("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & ("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & ("<tr>")
		strRst = strRst & ("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""상품명"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & ("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '돋움', sans-serif; margin:0; padding:4px 0 0 0;"">")
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
			strRst = strRst & "- 사이즈 : " & Fitemsize & "<br />"
		End if

		If Fitemsource <> "" Then
			strRst = strRst & "- 재료 : " &  Fitemsource & "<br />"
		End If

		'#기본 상품설명
		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & (Fitemcontent & "<br />")
			Case "H"
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
			Case Else
				strRst = strRst & (nl2br(Fitemcontent) & "<br />")
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
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br />")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ border=""0"" style=""width:100%""><br />")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ border=""0"" style=""width:100%""><br />")

		'#배송 주의사항
		strRst = strRst & ("<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_lotteon.jpg>")
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
			strRst = "<div align=""center""><p><img src=http://fiximage.10x10.co.kr/web2008/etc/top_notice_lotteon.jpg></p><br />" & strRst & "<br /><img src=http://fiximage.10x10.co.kr/web2008/etc/cs_info_lotteon.jpg></div>"
			retContents = strRst
		End If
		rsget.Close

		Set obj("spdLst")(null)("epnLst") = jsArray()
			Set obj("spdLst")(null)("epnLst")(0) = jsObject()
				obj("spdLst")(null)("epnLst")(0)("pdEpnTypCd") = "DSCRP"		'#상품설명유형코드 [공통코드 : PD_EPN_TYP_CD] | DSCRP : 상품기술서, AS_CNTS : A/S내용설명, PRCTN : 주의사항설명
				obj("spdLst")(null)("epnLst")(0)("cnts") = retContents			'#내용 | html입력시 사용한다.
		' If ForderComment <> "" Then
		' 	retOrderComment = "<div align=""center""><br />" & Fordercomment & "<br /></div>"
		' 	Set obj("spdLst")(null)("epnLst")(1) = jsObject()
		' 		obj("spdLst")(null)("epnLst")(1)("pdEpnTypCd") = "PRCTN"		'#상품설명유형코드 [공통코드 : PD_EPN_TYP_CD] | DSCRP : 상품기술서, AS_CNTS : A/S내용설명, PRCTN : 주의사항설명
		' 		obj("spdLst")(null)("epnLst")(1)("cnts") = retOrderComment		'#내용 | html입력시 사용한다.
		' End If
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

	Public Function getLotteonAddImageParam(obj)
		Dim addImages
		Dim strSql, i
		strSql = ""
		strSql = strSql & " SELECT TOP 30 gubun, ImgType, addimage_400, addimage_600, addimage_1000 "
		strSql = strSql & " FROM db_item.[dbo].tbl_item_addimage "
		strSql = strSql & " WHERE itemid=" & Fitemid
		strSql = strSql & " and isnull(addimage_400, '') <> '' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				addImages = addImages & "http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & "|"
				rsget.MoveNext
				If i>=9 Then Exit For
			Next
		End If
		rsget.Close

		If Right(addImages,1) = "|" Then
			addImages = Left(addImages,Len(addImages)-1)
		End If
		getLotteonAddImageParam = addImages
	End Function

	Public Function getLotteonOptionParameter(obj)
		Dim addImages, addImgSplit, i, j
		Dim limitsu, strSql
		Dim vlimitno, vlimitsold, vitemoption, voptionname, voptlimitno, voptlimitsold, voptsellyn, voptlimityn, voptaddprice
		Dim vMustprice
		vMustprice = mustPrice()
		addImages = getLotteonAddImageParam(obj)
		addImgSplit = Split(addImages, "|")

		If FOptionCnt = 0 Then			'단품
			obj("spdLst")(null)("sitmYn") = "N"													'#판매자단품여부 [Y, N] | Y이면 단품속성목록을 설정해야 한다. N이면 단품속성목록을 설정 안한다. 옵션이 없는 단품 한가지로 설정된다.
			Set obj("spdLst")(null)("itmLst") = jsArray()										'단품목록
				Set obj("spdLst")(null)("itmLst")(null) = jsObject()
					obj("spdLst")(null)("itmLst")(null)("eitmNo") = "0000"						'업체단품번호
					obj("spdLst")(null)("itmLst")(null)("sortSeq") = "1"						'#정렬순번
					obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"							'#전시여부 [Y, N]
					Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()			'결제수단예외목록 [공통코드 : PY_MNS_CD]
						obj("spdLst")(null)("itmLst")(null)("itmOptLst") = null
					Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()			'단품이미지목록 | 단품당 하나 이상의 이미지를 등록하여야 한다. 단품당 최대 10개의 이미지를 등록할 수 있다.
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
						If IsArray(addImgSplit) = True Then
							For i = 1 to Ubound(addImgSplit) + 1
								Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i) = jsObject()
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("origImgFileNm") = addImgSplit(i-1) '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("rprtImgYn") = "N"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
							Next
						End If
	'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'컬러칩이미지목록
	'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
	'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
	'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'상품단위기준정보
	'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#상품용량 | 기준단위와 기준용량은 표준카테고리 매핑 정보를 따른다. ex) 표준카테고리에 기준단위가 ml, 기준용량이 100으로 매핑되어 있는 경우 100ml당 가격이 표시된다.
					obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice						'#판매가
					obj("spdLst")(null)("itmLst")(null)("stkQty") = getLimitEa()					'#재고수량 | 재고관리여부가 Y인 경우에는 필수값
		Else							'옵션
			obj("spdLst")(null)("sitmYn") = "Y"													'#판매자단품여부 [Y, N] | Y이면 단품속성목록을 설정해야 한다. N이면 단품속성목록을 설정 안한다. 옵션이 없는 단품 한가지로 설정된다.
			Set obj("spdLst")(null)("itmLst") = jsArray()											'단품목록

			Dim vattr_id, vattr_nm
			Dim vattr_val_id, vattr_val_nm
			strSql = ""
			strSql = strSql & " SELECT TOP 1 a.attr_id, a.attr_nm, a.attr_disp_nm "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_Attribute as a "
			strSql = strSql & " WHERE attr_id in ( "
			strSql = strSql & " 	SELECT attr_id FROM db_etcmall.dbo.tbl_lotteon_StdCategory_Attr WHERE std_cat_id = '"& FStd_cat_id &"'  "
			strSql = strSql & " ) and attr_pi_type= 'I' and attr_disp_nm = '색상'  "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				vattr_id = rsget("attr_id")
				vattr_nm = rsget("attr_nm")
			End If
			rsget.Close

			If vattr_id = "" Then
				rw "맞는 속성 없음"
				Exit Function
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 attr_val_id, attr_val_nm FROM db_etcmall.dbo.tbl_lotteon_Attribute_Values "
				strSql = strSql & " WHERE attr_id = '"& vattr_id &"' and attr_val_nm = '멀티' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					vattr_val_id = rsget("attr_val_id")
					vattr_val_nm = rsget("attr_val_nm")
				End If
				rsget.Close
			End If

			If vattr_val_id = "" Then
				rw "맞는 속성값 없음"
				Exit Function
			End If

			strSql = ""
			strSql = strSql & " SELECT i.itemid, i.limityn, i.limitno ,i.limitsold, o.itemoption, optionname" & VBCRLF
			strSql = strSql & " , o.optlimitno, o.optlimitsold, o.optsellyn, o.optlimityn, o.optaddprice " & VBCRLF
			strSql = strSql & " FROM db_item.dbo.tbl_item as i " & VBCRLF
			strSql = strSql & " LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' " & VBCRLF
			strSql = strSql & " WHERE i.itemid = "&Fitemid
			strSql = strSql & " ORDER BY o.optaddprice ASC, o.itemoption ASC "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				For i = 1 to rsget.RecordCount
					If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
						vitemoption = "0000"
						voptionname = "단일상품"
						limitsu = getLimitEa()
						voptaddprice		= 0
					Else
						vitemoption 		= rsget("itemoption")
						voptionname 		= rsget("optionname")
						voptlimitno 		= rsget("optlimitno")
						voptlimitsold 		= rsget("optlimitsold")
						voptaddprice		= rsget("optaddprice")
						If FLimityn = "Y" Then
							If voptlimitno - voptlimitsold - 5 < 1 Then
								limitsu = 0
							Else
								limitsu = voptlimitno - voptlimitsold - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If
					End If
					Set obj("spdLst")(null)("itmLst")(null) = jsObject()
						obj("spdLst")(null)("itmLst")(null)("eitmNo") = vitemoption						'업체단품번호
						obj("spdLst")(null)("itmLst")(null)("sortSeq") = i								'#정렬순번
						obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"								'#전시여부 [Y, N]
						Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()				'o단품속성목록 | 판매자단품여부가 Y인 경우 필수값
							Set obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optCd") = vattr_id	'#옵션코드 [속성모듈 제공 항목] | 단품의 옵션에 해당하는 옵션코드를 입력하여야 한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optNm") = vattr_nm	'#옵션명 [속성모듈 제공 항목] | 해당 단품의 옵션명을 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optValCd") = vattr_val_id	'o옵션값코드 [속성모듈 제공 항목] | 입력하고자 하는 옵션값의 옵션값코드가 존재하지 않는 경우에는 옵션값만 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optVal") = vattr_val_nm	'o옵션값 [속성모듈 제공 항목] | 해당 단품의 옵션값을 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("dtlsVal") = voptionname	'세부값 | 세부값을 입력하는 경우 1. 범위값에 대한 고정값 입력시, 2. 옵션값에 대한 추가 표현
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'단품이미지목록 | 단품당 하나 이상의 이미지를 등록하여야 한다. 단품당 최대 10개의 이미지를 등록할 수 있다.
							Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
							If IsArray(addImgSplit) = True Then
								For j = 1 to Ubound(addImgSplit) + 1
									Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j) = jsObject()
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("origImgFileNm") = addImgSplit(j-1) '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("rprtImgYn") = "N"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
								Next
							End If
		'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'컬러칩이미지목록
		'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
		'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
		'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'상품단위기준정보
		'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#상품용량 | 기준단위와 기준용량은 표준카테고리 매핑 정보를 따른다. ex) 표준카테고리에 기준단위가 ml, 기준용량이 100으로 매핑되어 있는 경우 100ml당 가격이 표시된다.
						obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice + voptaddprice		'#판매가
						obj("spdLst")(null)("itmLst")(null)("stkQty") = limitsu							'#재고수량 | 재고관리여부가 Y인 경우에는 필수값
					rsget.MoveNext
				Next
			End If
			rsget.Close
		End If
	End Function

	Public Function getLotteonOptionEditParameter(obj)
		Dim addImages, addImgSplit, i, j
		Dim limitsu, strSql, arrRows
		Dim vlimitno, vlimitsold, vitemoption, voptionname, voptlimitno, voptlimitsold, voptsellyn, voptlimityn, voptaddprice
		Dim vMustprice, sitmNo
		vMustprice = mustPrice()
		addImages = getLotteonAddImageParam(obj)
		addImgSplit = Split(addImages, "|")

		strSql = "exec db_etcmall.dbo.usp_Ten_OutMall_optEditParamList_lotteon '"&CMallName&"'," & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
		    arrRows = rsget.getRows
		End If
		rsget.close

		If UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z" Then
			sitmNo = arrRows(15, 0)
		End If

		If FOptionCnt = 0 AND (UBound(arrRows,2) = 0 AND arrRows(0,0) = "Z") Then			'단품
			Set obj("spdLst")(null)("itmLst") = jsArray()										'단품목록
				Set obj("spdLst")(null)("itmLst")(null) = jsObject()
					obj("spdLst")(null)("itmLst")(null)("eitmNo") = "0000"						'업체단품번호
					obj("spdLst")(null)("itmLst")(null)("sitmNo") = ""&sitmNo&""				'판매자단품번호
					obj("spdLst")(null)("itmLst")(null)("sortSeq") = "1"						'정렬순번
					obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"							'전시여부 [Y, N]
					Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()			'결제수단예외목록 [공통코드 : PY_MNS_CD]
						obj("spdLst")(null)("itmLst")(null)("itmOptLst") = null
					Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'단품이미지목록 | 단품당 하나 이상의 이미지를 등록하여야 한다. 단품당 최대 10개의 이미지를 등록할 수 있다.
						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
							obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
						If IsArray(addImgSplit) = True Then
							For i = 1 to Ubound(addImgSplit) + 1
								Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i) = jsObject()
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("origImgFileNm") = addImgSplit(i-1) '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
									obj("spdLst")(null)("itmLst")(null)("itmImgLst")(i)("rprtImgYn") = "N"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
							Next
						End If
	'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'컬러칩이미지목록
	'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
	'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
	'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'상품단위기준정보
	'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#상품용량 | 기준단위와 기준용량은 표준카테고리 매핑 정보를 따른다. ex) 표준카테고리에 기준단위가 ml, 기준용량이 100으로 매핑되어 있는 경우 100ml당 가격이 표시된다.
					obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice						'#판매가
					obj("spdLst")(null)("itmLst")(null)("stkQty") = getLimitEa()					'#재고수량 | 재고관리여부가 Y인 경우에는 필수값
		Else							'옵션
			Set obj("spdLst")(null)("itmLst") = jsArray()											'단품목록

			Dim vattr_id, vattr_nm
			Dim vattr_val_id, vattr_val_nm
			strSql = ""
			strSql = strSql & " SELECT TOP 1 a.attr_id, a.attr_nm, a.attr_disp_nm "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_lotteon_Attribute as a "
			strSql = strSql & " WHERE attr_id in ( "
			strSql = strSql & " 	SELECT attr_id FROM db_etcmall.dbo.tbl_lotteon_StdCategory_Attr WHERE std_cat_id = '"& FStd_cat_id &"'  "
			strSql = strSql & " ) and attr_pi_type= 'I' and attr_disp_nm = '색상'  "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				vattr_id = rsget("attr_id")
				vattr_nm = rsget("attr_nm")
			Else
				rw "맞는 속성 없음"
				Exit Function
			End If
			rsget.Close

			If vattr_id <> "" Then
				strSql = ""
				strSql = strSql & " SELECT attr_val_id, attr_val_nm FROM db_etcmall.dbo.tbl_lotteon_Attribute_Values "
				strSql = strSql & " WHERE attr_id = '"& vattr_id &"' and attr_val_nm = '멀티' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					vattr_val_id = rsget("attr_val_id")
					vattr_val_nm = rsget("attr_val_nm")
				Else
					rw "맞는 속성값 없음"
					rsget.Close
					Exit Function
				End If
				rsget.Close
			End If

			If IsArray(arrRows) Then
				For i = 0 To UBound(arrRows, 2)
					vitemoption 		= arrRows(1, i)
					voptionname 		= arrRows(3, i)
					voptaddprice		= arrRows(16, i)
					sitmNo 				= arrRows(15, i)
					If FLimityn = "Y" Then
						If arrRows(4, i) - 5 < 1 Then
							limitsu = 0
						Else
							limitsu = arrRows(4, i) - 5
						End If
					Else
						limitsu = CDEFALUT_STOCK
					End If

					Set obj("spdLst")(null)("itmLst")(null) = jsObject()
						obj("spdLst")(null)("itmLst")(null)("eitmNo") = vitemoption						'업체단품번호
						obj("spdLst")(null)("itmLst")(null)("sitmNo") = ""&sitmNo&""					'판매자단품번호
						obj("spdLst")(null)("itmLst")(null)("sortSeq") = i								'#정렬순번
						obj("spdLst")(null)("itmLst")(null)("dpYn") = "Y"								'#전시여부 [Y, N]

					If (ArrRows(11,i)=0) and ArrRows(12,i) = "1" AND ArrRows(15,i) = "" Then		'옵션명이 다르고 옵션코드값이 없을 때 ==> 단품추가 의미// preged 0
						Set obj("spdLst")(null)("itmLst")(null)("itmOptLst") = jsArray()				'o단품속성목록 | 판매자단품여부가 Y인 경우 필수값
							Set obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optCd") = vattr_id	'#옵션코드 [속성모듈 제공 항목] | 단품의 옵션에 해당하는 옵션코드를 입력하여야 한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optNm") = vattr_nm	'#옵션명 [속성모듈 제공 항목] | 해당 단품의 옵션명을 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optValCd") = vattr_val_id	'o옵션값코드 [속성모듈 제공 항목] | 입력하고자 하는 옵션값의 옵션값코드가 존재하지 않는 경우에는 옵션값만 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("optVal") = vattr_val_nm	'o옵션값 [속성모듈 제공 항목] | 해당 단품의 옵션값을 입력한다.
								obj("spdLst")(null)("itmLst")(null)("itmOptLst")(null)("dtlsVal") = voptionname	'세부값 | 세부값을 입력하는 경우 1. 범위값에 대한 고정값 입력시, 2. 옵션값에 대한 추가 표현
					End If

						Set obj("spdLst")(null)("itmLst")(null)("itmImgLst") = jsArray()		'단품이미지목록 | 단품당 하나 이상의 이미지를 등록하여야 한다. 단품당 최대 10개의 이미지를 등록할 수 있다.
							Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0) = jsObject()
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("origImgFileNm") = FbasicImage '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
								obj("spdLst")(null)("itmLst")(null)("itmImgLst")(0)("rprtImgYn") = "Y"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
							If IsArray(addImgSplit) = True Then
								For j = 1 to Ubound(addImgSplit) + 1
									Set obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j) = jsObject()
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypCd") = "IMG"	'#노출유형코드 [공통코드 : EPSR_TYP_CD] | IMG : 이미지
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("epsrTypDtlCd") = "IMG_SQRE"	'#노출유형상세코드 [공통코드 : EPSR_TYP_DTL_CD] | IMG_SQRE : 노출유형:이미지 > 정사각형, IMG_LNTH : 노출유형:이미지 > 세로형
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("origImgFileNm") = addImgSplit(j-1) '#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
										obj("spdLst")(null)("itmLst")(null)("itmImgLst")(j)("rprtImgYn") = "N"	'#대표이미지여부 [Y, N] | 대표이미지는 하나만 설정 가능
								Next
							End If
		'				Set obj("spdLst")(null)("itmLst")(null)("clrchipLst") = jsArray()				'컬러칩이미지목록
		'					Set obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null) = jsObject()
		'						obj("spdLst")(null)("itmLst")(null)("clrchipLst")(null)("origImgFileNm") = ""	'#원본이미지파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.jpg
		'				Set obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo") = jsObject()				'상품단위기준정보
		'					obj("spdLst")(null)("itmLst")(null)("pdUtStdInfo")("pdCapa") = ""			'#상품용량 | 기준단위와 기준용량은 표준카테고리 매핑 정보를 따른다. ex) 표준카테고리에 기준단위가 ml, 기준용량이 100으로 매핑되어 있는 경우 100ml당 가격이 표시된다.
						obj("spdLst")(null)("itmLst")(null)("slPrc") = vMustprice + voptaddprice		'#판매가
						obj("spdLst")(null)("itmLst")(null)("stkQty") = limitsu							'#재고수량 | 재고관리여부가 Y인 경우에는 필수값

				Next
			End If
		End If
	End Function

	'상품등록 Json
	Public Function getLotteonItemRegParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj, tenBeasongDay
		tenBeasongDay = getShopLeadTime()

		If FItemdiv = "06" OR FItemdiv = "16" Then
			dvPdTypCd = "OD_MFG"
			sndBgtNday = "15"
		Else
			If tenBeasongDay > 3 Then
				dvPdTypCd = "OD_MFG"
				sndBgtNday = tenBeasongDay
			Else
				dvPdTypCd = "GNRL"
				sndBgtNday = "3"
			End If
		End If

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()														'등록상품목록
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"										'#거래처그룹코드 | SR : 일반셀러
					obj("spdLst")(null)("trNo") = afflTrCd										'#거래처번호
					obj("spdLst")(null)("lrtrNo") = ""											'하위거래처번호
					obj("spdLst")(null)("scatNo") = FStd_cat_id									'#표준카테고리번호
					Set obj("spdLst")(null)("dcatLst") = jsArray()								'#전시카테고리목록 | 속성모듈의 API를 통하여 표준카테고리에 매핑된 전시카테고리를 정보를 받는다. 매핑된 전시카테고리 중에서 하나 이상 선택하여 입력한다.
						Set obj("spdLst")(null)("dcatLst")(null) = jsObject()
							obj("spdLst")(null)("dcatLst")(null)("mallCd") = "LTON"				'#몰구분코드 | LTON : 롯데ON
							obj("spdLst")(null)("dcatLst")(null)("lfDcatNo") = FDisp_cat_id		'#leaf전시카테고리번호
'							obj("spdLst")(null)("dcatLst")(null)("dcatNo") = ""					'--예시는 있는데 설명은 없음..;
					obj("spdLst")(null)("epdNo") = ""&FItemid&""								'업체상품번호
					obj("spdLst")(null)("slTypCd") = "GNRL"										'#판매유형코드 | 사은품은 사은품등록 API를 사용한다. GNRL : 일반판매상품, CNSL : 상담판매상품
					obj("spdLst")(null)("pdTypCd") = "GNRL_GNRL"								'#상품유형코드 | 사은품은 사은품등록 API를 사용한다. GNRL_GNRL : 일반판매_일반상품, GNRL_ECPN : 일반판매e쿠폰상품, GNRL_GFTV : 일반판매_상품권, GNRL_ZRWON : 일반판매_0원상품, CNSL_CNSL : 상담판매_상담상품
					obj("spdLst")(null)("gftvShpCd") = null										'o상품유형구분코드가 GNRL_GFTV(상품권)인 경우에는 필수 입력, 모바일상품권의 경우에는 e쿠폰 항목을 입력하여야 한다. | PPR : 지류, MBL : 모바일
					obj("spdLst")(null)("spdNm") = getItemNameFormat()							'#판매자상품명 | 입력된 판매자상품명은 상품명 정제를 거쳐 전시상품명으로 노출된다.
					obj("spdLst")(null)("brdNo") = getBrandCode()								'브랜드번호 [속성모듈 제공 항목] | 속성모듈 API를 통하여 수신된 브랜드번호를 입력한다.
					obj("spdLst")(null)("mfcrNm") = CStr(FMakerName)							'제조사명 | TXT 값으로 입력한다.
					obj("spdLst")(null)("oplcCd") = getOriginCode()								 '#원산지코드 | 기타인 경우에는 "상품상세 참조"코드(ETC) 입력
					obj("spdLst")(null)("mdlNo") = ""											'모델번호
					obj("spdLst")(null)("barCd") = ""											'바코드
					obj("spdLst")(null)("tdfDvsCd") = CHKIIF(FVatInclude="N","02","01")			'#과세유형코드 [공통코드 : TDF_DVS_CD] | 01: 과세, 02 : 면세, 03 : 영세, 04 : 해당없음
					obj("spdLst")(null)("slStrtDttm") = FormatDate(now(), "00000000000000")		'#판매시작일시 [YYYYMMDDHH24MISS ex) 20190801100000]
					obj("spdLst")(null)("slEndDttm") = "99991231235959"							'#판매종료일시 [YYYYMMDDHH24MISS ex) 20190801100000]
					Call getLotteonInfoCdParameter(obj)											'#상품품목고시정보
					Call getLotteonCertInfoParameter(obj)										'안전인증목록
					Call getLotteonStdCateAttrParameter(obj)									'표준카테고리속성목록
					If FItemdiv = "06" Then
					Set obj("spdLst")(null)("itypOptLst") = jsArray()							'입력형옵션목록 | 최대 5개의 입력형옵션을 설정할 수 있다.
						Set obj("spdLst")(null)("itypOptLst")(null) = jsObject()
							obj("spdLst")(null)("itypOptLst")(null)("itypOptDvsCd") = "TXT"		'#입력형옵션구분코드 [공통코드 : ITYP_OPT_DVS_CD] | NO : 숫자, TXT : 텍스트, DATE : 달력형, TIME : 시간선택형
							obj("spdLst")(null)("itypOptLst")(null)("itypOptNm") = "텍스트를 입력하세요"	'#입력형옵션명
					End If
					Set obj("spdLst")(null)("purPsbQtyInfo") = jsObject()						'구매가능수량정보
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurYn") = "N"				'#단품별최소구매여부 [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurQty") = null			'o단품별최소구매수량 | 단품별최소구매여부가 Y인 경우 필수입력한다.
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMaxPurPsbQtyYn") = "Y"		'#단품별최대구매가능수량여부 [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("maxPurQty") = getOrderMaxNum		'o단품별최대구매수량 | 단품별최대구매가능수량여부가 Y인 경우 필수입력한다.
					obj("spdLst")(null)("ageLmtCd") = Chkiif(IsAdultItem()="Y", "19", "0")		'#연령제한코드 0 : 전연령 구매가능, 15 : 15세이상 구매가능, 19 : 19세이상 구매가능
					obj("spdLst")(null)("prstPsbYn") = "N"										'선물가능여부 [Y, N] | 디폴트:N
'					obj("spdLst")(null)("prstPckPsbYn") = ""									'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("prstMsgPsbYn") = ""									'--예시는 있는데 설명은 없음..;
					obj("spdLst")(null)("prcCmprEpsrYn") = "Y"									'가격비교노출여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("bookCultCstDdctYn") = "N"								'도서문화비 공제여부 [Y, N] | 디폴트:N 거래처와 표준카테고리가 모두 도서문화비 공제대상에 해당하는 경우에만 공제여부가 Y이다.
					obj("spdLst")(null)("isbnCd") = ""											'oISBN | 도서문화비 공제여부가 Y이고 카테고리가 도서관련 카테고리일 경우 ISBN NO를 입력한다.
'					obj("spdLst")(null)("impCoNm") = ""											'수입사명 | TXT 입력
'					obj("spdLst")(null)("impDvsCd") = "NONE"									'수입구분코드 [공통코드 : IMP_DVS_CD] | 수입사명이 있는 경우 입력한다. DRC_IMP : 직수입, PRL_IMP : 병행수입, NONE : 해당없음
					obj("spdLst")(null)("cshbltyPdYn") = "N"									'환금성상품여부 [Y, N] | 표준카테고리 속성을 상속 받는다. 환금성 상품으로 설정되는 경우 주문에서 결제수단에 따라 구매가 제한된다. 디폴트:N
'					obj("spdLst")(null)("dnDvPdYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("toysPdYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("intgSlPdNo") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("nmlPdYn") = ""											'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("prmmPdYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("otltPdYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("prmmInstPdYn") = ""									'--예시는 있는데 설명은 없음..;
					obj("spdLst")(null)("brkHmapPkcpPsbYn") = "N"								'폐가전수거여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("ctrtTypCd") = "A"										'계약유형코드[공통코드 : CTRT_TYP_CD] | A : 중개, B : 위탁
'					Set obj("spdLst")(null)("pdSzInfo") = jsObject()							'배송사이즈정보 | 정수만 입력 가능하다.
'						obj("spdLst")(null)("pdSzInfo")("pdWdthSz") = ""						'상품가로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdLnthSz") = ""						'상품세로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdHghtSz") = ""						'상품높이사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckWdthSz") = ""						'포장가로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckLnthSz") = ""						'포장세로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckHghtSz") = ""						'포장높이사이즈 (cm)
					obj("spdLst")(null)("pdStatCd") = "NEW"										'#상품상태코드 [공통코드 : PD_STAT_CD] | 상품상태코드가 새상품(NEW)이 아닌 경우에는 파일유형코드와 파일구분코드를 USD로 하여 상품상태이미지를 반드시 등록하여야 한다.
					obj("spdLst")(null)("dpYn") = "Y"											'전시여부 [Y, N] | 디폴트:Y
'					obj("spdLst")(null)("ltonDpYn") = ""										'--예시는 있는데 설명은 없음..;
					Call getLotteonKeywordsParameter(obj)										'검색키워드목록 | 5개 이하만 등록 가능
'					Set obj("spdLst")(null)("pdFileLst") = jsArray()							'o상품콘텐츠파일목록 | 상품상태코드가 새상품(NEW)이 아닌 경우에는 파일유형코드와 파일구분코드를 USD로 하여 상품상태이미지를 반드시 등록하여야 한다.
'						Set obj("spdLst")(null)("pdFileLst")(null) = jsObject()
'							obj("spdLst")(null)("pdFileLst")(null)("fileTypCd") = ""			'#파일유형코드 [공통코드 : FILE_TYP_CD] | USD : 상품상태, TAG_LBL : Tag/케어라벨, PD : 상품
'							obj("spdLst")(null)("pdFileLst")(null)("fileDvsCd") = ""			'#파일구분코드 [공통코드 : FILE_DVS_CD] | USD : 상품상태, TAG_LBL : Tag/케어라벨, 3D : 상품3D이미지, WDTH : 상품가로형, VDO_FILE : 상품동영상_FILE, VDO_URL : 상품동영상_URL
'							obj("spdLst")(null)("pdFileLst")(null)("origFileNm") = ""			'#원본파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.mp4
					Call getLotteonContParamToReg(obj)											'상품설명목록
					Set obj("spdLst")(null)("pyMnsExcpLst") = jsArray()							'결제수단예외목록 [공통코드 : PY_MNS_CD]
						obj("spdLst")(null)("pyMnsExcpLst") = null
					obj("spdLst")(null)("cnclPsbYn") = "Y"										'취소가능여부 [Y, N] | 취소 불가인 상품인 경우에는 'N'으로 설정 디폴트:Y
					obj("spdLst")(null)("immdCnclPsbYn") = "N"									'즉시취소가능여부 [Y, N] | 특정 시점(출고 등)까지는 문의 없이 바로 취소 가능한 경우 "Y"로 설정 디플트: Y
					obj("spdLst")(null)("dmstOvsDvDvsCd") = "DMST"								'국내해외배송구분코드 [공통코드 : DMST_OVS_DV_DVS_CD] | 디폴트:국내배송, DMST : 국내배송, OVS : 해외발송, RVRS_DPUR : 역직구
					obj("spdLst")(null)("pstkYn") = "N"											'선재고여부 [Y, N] 디폴트:N
					obj("spdLst")(null)("dvProcTypCd") = "LO_ENTP"								'#배송처리유형코드 [공통코드 : DV_PROC_TYP_CD] | LO_CNTR : e커머스 센터배송, LO_ENTP : e커머스 업체배송
					obj("spdLst")(null)("dvPdTypCd") = dvPdTypCd								'#배송상품유형코드 [공통코드 : DV_PD_TYP_CD] | TDY_SND : 오늘발송(0일), GNRL : 일반상품(3일), OD_MFG : 주문제작상품(15일), FREE_INST : 무료설치상품(3일), CHRG_INST : 유료설치상품(3일), PRMM_INST : 프리미엄설치상품(365일), ECPN : e쿠폰(0일), GFTV : 상품권(3일), OVS : 해외배송(15일)
					obj("spdLst")(null)("sndBgtNday") = sndBgtNday								'발송예정일수 | 배송상품유형코드에 따라 최대 발송예정일수를 입력한다.
					Set obj("spdLst")(null)("sndBgtDdInfo") = jsObject()						'발송예정일정보
						obj("spdLst")(null)("sndBgtDdInfo")("nldySndCloseTm") = "1500"			'#평일 발송마감시간 [HH24MI ex) 1000]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndPsbYn") = "Y"				'#토요일 발송가능여부 [Y, N]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndCloseTm") = "1300"			'o토요일 발송마감시간 [HH24MI ex) 1000] | 토요일 발송 가능여부 Y인 경우 필수
					obj("spdLst")(null)("dvRgsprGrpCd") = "GN101"								'#배송권역그룹코드 | 배송모듈을 통하여 관리되는 코드를 입력한다. | GN000(전국), GN004(제주), GN006(도서산간), GN101(전국(일부지역제외), GN102(전국(제주도 및 도서지역 제외), GN103(서울 및 수도권), GN104(전국 + 해외), GN105(서울)
					obj("spdLst")(null)("dvMnsCd") = "DPCL"										'#배송수단코드 [공통코드 : DV_MNS_CD] 단건만 입력가능 | DGNN_DV : 전담배송(직접배송), DPCL : 택배, NONE_DV : 무배송, REG_MAIL : 등기, ZIP : 우편
					obj("spdLst")(null)("owhpNo") = DVPCd(1)									'#출고지번호 | 거래처 API "(일반 Seller용) 판매자 출고지/반품지 등록"을 통하여 등록된 출고지번호를 입력한다.
					obj("spdLst")(null)("hdcCd") = "0002"										'#택배사코드 [공통코드 : DV_CO_CD] | 0002 : 대한통운
					obj("spdLst")(null)("dvCstPolNo") = DVPCd(0)								'#배송비정책번호 | 거래처의 API를 통해 선등록된 배송비정책번호를 입력한다.
					obj("spdLst")(null)("adtnDvCstPolNo") = DVPCd(3)							'추가배송비정책번호 | 거래처의 API를 통해 선등록된 추가배송비정책번호를 입력한다.
					obj("spdLst")(null)("cmbnDvPsbYn") = "Y"									'합배송가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("dvCstStdQty") = "0"									'배송비기준수량 | 디폴트:0
					obj("spdLst")(null)("qckDvUseYn") = "N"										'퀵배송사용여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("crdayDvPsbYn") = "N"									'당일배송가능여부 [Y, N] | 디폴트:N
'					Set obj("spdLst")(null)("crdayDvInfo") = jsObject()							'o당일배송정보 | 당일배송가능여부가 Y인 경우 필수값
'						obj("spdLst")(null)("crdayDvInfo")("odCloseTm") = ""					'#주문마감시간 [HH24MI ex) 1000] | 당일배송가능여부가 Y인 경우 필수값
					obj("spdLst")(null)("spicUseYn") = "N"										'스마트픽사용여부 [Y, N] | 디폴트:N
					Set obj("spdLst")(null)("spicInfo") = jsObject()							'스마트픽정보 | 스마트픽사용여부 Y인 경우 필수
						obj("spdLst")(null)("spicInfo") = null
'					obj("spdLst")(null)("spicEusePdYn") = ""									'--예시는 있는데 설명은 없음..;
					obj("spdLst")(null)("hpDdDvPsbYn") = "N"									'희망일배송가능여부 [Y, N] 디폴트:N
'					obj("spdLst")(null)("hpDdDvPsbPrd") = ""									'희망일배송가능기간 | 희망일배송가능여부 Y인 경우 필수
					obj("spdLst")(null)("saveTypCd") = "NONE"									'저장유형코드 [공통코드 : SAVE_TYP_CD] | 디폴트:해당없음 RFRG : 냉장, FRZN : 냉동, FRSH : 신선, NONE : 해당없음
'					obj("spdLst")(null)("shopCnvMsgPsbYn") = ""									'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("rgnLmtPdYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("fprdDvPsbYn") = ""										'--예시는 있는데 설명은 없음..;
'					obj("spdLst")(null)("spcfSqncPdYn") = ""									'--예시는 있는데 설명은 없음..;
					obj("spdLst")(null)("rtngPsbYn") = "Y"										'반품가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("xchgPsbYn") = "Y"										'교환가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("echgPsbYn") = "N"										'맞교환가능여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("cmbnRtngPsbYn") = "Y"									'합반품가능여부 [Y, N] | 합배송가능여부가 Y인 경우 Y, N 선택 가능. N인 경우 N만 선택 가능
					obj("spdLst")(null)("rtngHdcCd") = ""										'반품택배사코드 | 0002 : 대한통운
					obj("spdLst")(null)("rtngRtrvPsbYn") = "Y"									'반품회수가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("rtrpNo") = DVPCd(2)									'#회수지번호 | 거래처 API "(일반 Seller용) 판매자 출고지/반품지 등록"을 통하여 등록된 회수지번호를 입력한다.
'					Set obj("spdLst")(null)("ecpnInfo") = jsObject()							'(생략)e쿠폰정보 | 해당 상품이 e쿠폰인 경우에만 입력한다.
'					Set obj("spdLst")(null)("rntlPdInfo") = jsObject()							'(생략)렌탈상품정보 | 상품유형이 렌탈일 경우 필수값
'					Set obj("spdLst")(null)("opngPdInfo") = jsObject()							'(생략)개통형상품정보 | 상품유형구분코드가 일반판매_0원상품(GNRL_ZRWON)에 해당하는 개통형상품인 경우 필수입력한다.
					obj("spdLst")(null)("stkMgtYn") = "Y"										'#재고관리여부 [Y, N] | 'N'인 경우 재고가 999,999,999로 들어간다. 웹재고를 관리하지 않는다.
					Call getLotteonOptionParameter(obj)											'단품목록
'					Set obj("spdLst")(null)("slrRcPdLst") = jsArray()							'셀러추천상품목록 | 최대 10개까지 등록 가능하다.
'						Set obj("spdLst")(null)("slrRcPdLst")(null) = jsObject()
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSpdNo") = ""			'#셀러추천판매자상품번호
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSitmNo") = ""			'#셀러추천판매자단품번호
'							obj("spdLst")(null)("slrRcPdLst")(null)("epsrPrirRnkg") = ""		'#노출우선순위
		getLotteonItemRegParameter = obj.jsString
'   response.write getLotteonItemRegParameter
'   response.end
	End Function

	'상품수정 Json
	Public Function getLotteonItemEditParameter
		Dim strRst, dvPdTypCd, sndBgtNday
		Dim obj, tenBeasongDay
		tenBeasongDay = getShopLeadTime()

		If FItemdiv = "06" OR FItemdiv = "16" Then
			dvPdTypCd = "OD_MFG"
			sndBgtNday = "15"
		Else
			If tenBeasongDay > 3 Then
				dvPdTypCd = "OD_MFG"
				sndBgtNday = tenBeasongDay
			Else
				dvPdTypCd = "GNRL"
				sndBgtNday = "3"
			End If
		End If

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()														'등록상품목록
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"										'#거래처그룹코드 | SR : 일반셀러
					obj("spdLst")(null)("trNo") = afflTrCd										'#거래처번호
					obj("spdLst")(null)("lrtrNo") = ""											'하위거래처번호
					obj("spdLst")(null)("scatNo") = FStd_cat_id									'#표준카테고리번호
					Set obj("spdLst")(null)("dcatLst") = jsArray()								'#전시카테고리목록 | 속성모듈의 API를 통하여 표준카테고리에 매핑된 전시카테고리를 정보를 받는다. 매핑된 전시카테고리 중에서 하나 이상 선택하여 입력한다.
						Set obj("spdLst")(null)("dcatLst")(null) = jsObject()
							obj("spdLst")(null)("dcatLst")(null)("mallCd") = "LTON"				'#몰구분코드 | LTON : 롯데ON
							obj("spdLst")(null)("dcatLst")(null)("lfDcatNo") = FDisp_cat_id		'#leaf전시카테고리번호
					obj("spdLst")(null)("spdNo") = ""&FLotteonGoodNo&""							'#판매자상품번호
					obj("spdLst")(null)("spdNm") = getItemNameFormat()							'판매자상품명 | 입력된 판매자상품명은 상품명 정제를 거쳐 전시상품명으로 노출된다.
					obj("spdLst")(null)("brdNo") = getBrandCode()								'브랜드번호 [속성모듈 제공 항목] | 속성모듈 API를 통하여 수신된 브랜드번호를 입력한다.
					obj("spdLst")(null)("mfcrNm") = CStr(FMakerName)							'제조사명 | TXT 값으로 입력한다.
					obj("spdLst")(null)("oplcCd") = getOriginCode()								 '원산지코드 | 기타인 경우에는 "상품상세 참조"코드(ETC) 입력
					obj("spdLst")(null)("mdlNo") = ""											'모델번호
					obj("spdLst")(null)("barCd") = ""											'바코드
					obj("spdLst")(null)("tdfDvsCd") = CHKIIF(FVatInclude="N","02","01")			'#과세유형코드 [공통코드 : TDF_DVS_CD] | 01: 과세, 02 : 면세, 03 : 영세, 04 : 해당없음
					obj("spdLst")(null)("slStrtDttm") = FormatDate(now(), "00000000000000")		'#판매시작일시 [YYYYMMDDHH24MISS ex) 20190801100000]
					obj("spdLst")(null)("slEndDttm") = "99991231235959"							'#판매종료일시 [YYYYMMDDHH24MISS ex) 20190801100000]
					Call getLotteonInfoCdParameter(obj)											'#상품품목고시정보
					Call getLotteonCertInfoParameter(obj)										'안전인증목록
					If FItemdiv = "06" Then
					Set obj("spdLst")(null)("itypOptLst") = jsArray()							'입력형옵션목록 | 최대 5개의 입력형옵션을 설정할 수 있다.
						Set obj("spdLst")(null)("itypOptLst")(null) = jsObject()
							obj("spdLst")(null)("itypOptLst")(null)("itypOptDvsCd") = "TXT"		'#입력형옵션구분코드 [공통코드 : ITYP_OPT_DVS_CD] | NO : 숫자, TXT : 텍스트, DATE : 달력형, TIME : 시간선택형
							obj("spdLst")(null)("itypOptLst")(null)("itypOptNm") = "텍스트를 입력하세요"	'#입력형옵션명
					End If
					Set obj("spdLst")(null)("purPsbQtyInfo") = jsObject()						'구매가능수량정보
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurYn") = "N"				'#단품별최소구매여부 [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMinPurQty") = null			'o단품별최소구매수량 | 단품별최소구매여부가 Y인 경우 필수입력한다.
						obj("spdLst")(null)("purPsbQtyInfo")("itmByMaxPurPsbQtyYn") = "Y"		'#단품별최대구매가능수량여부 [Y, N]
						obj("spdLst")(null)("purPsbQtyInfo")("maxPurQty") = getOrderMaxNum		'o단품별최대구매수량 | 단품별최대구매가능수량여부가 Y인 경우 필수입력한다.
					obj("spdLst")(null)("prstPsbYn") = "N"										'선물가능여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("prcCmprEpsrYn") = "Y"									'가격비교노출여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("bookCultCstDdctYn") = "N"								'도서문화비 공제여부 [Y, N] | 디폴트:N 거래처와 표준카테고리가 모두 도서문화비 공제대상에 해당하는 경우에만 공제여부가 Y이다.
					obj("spdLst")(null)("isbnCd") = ""											'oISBN | 도서문화비 공제여부가 Y이고 카테고리가 도서관련 카테고리일 경우 ISBN NO를 입력한다.
'					obj("spdLst")(null)("impCoNm") = ""											'수입사명 | TXT 입력
'					obj("spdLst")(null)("impDvsCd") = "NONE"									'수입구분코드 [공통코드 : IMP_DVS_CD] | 수입사명이 있는 경우 입력한다. DRC_IMP : 직수입, PRL_IMP : 병행수입, NONE : 해당없음
					obj("spdLst")(null)("cshbltyPdYn") = "N"									'환금성상품여부 [Y, N] | 표준카테고리 속성을 상속 받는다. 환금성 상품으로 설정되는 경우 주문에서 결제수단에 따라 구매가 제한된다. 디폴트:N
					obj("spdLst")(null)("brkHmapPkcpPsbYn") = "N"								'폐가전수거여부 [Y, N] | 디폴트:N
'					Set obj("spdLst")(null)("pdSzInfo") = jsObject()							'배송사이즈정보 | 정수만 입력 가능하다.
'						obj("spdLst")(null)("pdSzInfo")("pdWdthSz") = ""						'상품가로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdLnthSz") = ""						'상품세로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pdHghtSz") = ""						'상품높이사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckWdthSz") = ""						'포장가로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckLnthSz") = ""						'포장세로사이즈 (cm)
'						obj("spdLst")(null)("pdSzInfo")("pckHghtSz") = ""						'포장높이사이즈 (cm)
					obj("spdLst")(null)("dpYn") = "Y"											'전시여부 [Y, N] | 디폴트:Y
					Call getLotteonKeywordsParameter(obj)										'검색키워드목록 | 5개 이하만 등록 가능
'					Set obj("spdLst")(null)("pdFileLst") = jsArray()							'o상품콘텐츠파일목록 | 상품상태코드가 새상품(NEW)이 아닌 경우에는 파일유형코드와 파일구분코드를 USD로 하여 상품상태이미지를 반드시 등록하여야 한다.
'						Set obj("spdLst")(null)("pdFileLst")(null) = jsObject()
'							obj("spdLst")(null)("pdFileLst")(null)("fileTypCd") = ""			'#파일유형코드 [공통코드 : FILE_TYP_CD] | USD : 상품상태, TAG_LBL : Tag/케어라벨, PD : 상품
'							obj("spdLst")(null)("pdFileLst")(null)("fileDvsCd") = ""			'#파일구분코드 [공통코드 : FILE_DVS_CD] | USD : 상품상태, TAG_LBL : Tag/케어라벨, 3D : 상품3D이미지, WDTH : 상품가로형, VDO_FILE : 상품동영상_FILE, VDO_URL : 상품동영상_URL
'							obj("spdLst")(null)("pdFileLst")(null)("origFileNm") = ""			'#원본파일명(경로명) | 파일명을 포함한 다운로드가 가능한 경로를 입력한다. ex) http://abc.com/12/34/56/78_90.mp4
					Call getLotteonContParamToReg(obj)											'상품설명목록
					Set obj("spdLst")(null)("pyMnsExcpLst") = jsArray()							'결제수단예외목록 [공통코드 : PY_MNS_CD]
						obj("spdLst")(null)("pyMnsExcpLst") = null
					obj("spdLst")(null)("cnclPsbYn") = "Y"										'취소가능여부 [Y, N] | 취소 불가인 상품인 경우에는 'N'으로 설정 디폴트:Y
					obj("spdLst")(null)("immdCnclPsbYn") = "N"									'즉시취소가능여부 [Y, N] | 특정 시점(출고 등)까지는 문의 없이 바로 취소 가능한 경우 "Y"로 설정 디플트: Y
					obj("spdLst")(null)("dvPdTypCd") = dvPdTypCd								'#배송상품유형코드 [공통코드 : DV_PD_TYP_CD] | TDY_SND : 오늘발송(0일), GNRL : 일반상품(3일), OD_MFG : 주문제작상품(15일), FREE_INST : 무료설치상품(3일), CHRG_INST : 유료설치상품(3일), PRMM_INST : 프리미엄설치상품(365일), ECPN : e쿠폰(0일), GFTV : 상품권(3일), OVS : 해외배송(15일)
					obj("spdLst")(null)("sndBgtNday") = sndBgtNday								'발송예정일수 | 배송상품유형코드에 따라 최대 발송예정일수를 입력한다.
					Set obj("spdLst")(null)("sndBgtDdInfo") = jsObject()						'발송예정일정보
						obj("spdLst")(null)("sndBgtDdInfo")("nldySndCloseTm") = "1500"			'#평일 발송마감시간 [HH24MI ex) 1000]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndPsbYn") = "Y"				'#토요일 발송가능여부 [Y, N]
						obj("spdLst")(null)("sndBgtDdInfo")("satSndCloseTm") = "1300"			'o토요일 발송마감시간 [HH24MI ex) 1000] | 토요일 발송 가능여부 Y인 경우 필수
					obj("spdLst")(null)("dvRgsprGrpCd") = "GN101"								'배송권역그룹코드 | 배송모듈을 통하여 관리되는 코드를 입력한다. | GN000(전국), GN004(제주), GN006(도서산간), GN101(전국(일부지역제외), GN102(전국(제주도 및 도서지역 제외), GN103(서울 및 수도권), GN104(전국 + 해외), GN105(서울)
					obj("spdLst")(null)("dvMnsCd") = "DPCL"										'#배송수단코드 [공통코드 : DV_MNS_CD] 단건만 입력가능 | DGNN_DV : 전담배송(직접배송), DPCL : 택배, NONE_DV : 무배송, REG_MAIL : 등기, ZIP : 우편
					obj("spdLst")(null)("owhpNo") = DVPCd(1)									'#출고지번호 | 거래처 API "(일반 Seller용) 판매자 출고지/반품지 등록"을 통하여 등록된 출고지번호를 입력한다.
					obj("spdLst")(null)("hdcCd") = "0002"										'#택배사코드 [공통코드 : DV_CO_CD] | 0002 : 대한통운
					obj("spdLst")(null)("dvCstPolNo") = DVPCd(0)								'#배송비정책번호 | 거래처의 API를 통해 선등록된 배송비정책번호를 입력한다.
					obj("spdLst")(null)("adtnDvCstPolNo") = DVPCd(3)							'추가배송비정책번호 | 거래처의 API를 통해 선등록된 추가배송비정책번호를 입력한다.
					obj("spdLst")(null)("cmbnDvPsbYn") = "Y"									'합배송가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("dvCstStdQty") = "0"									'배송비기준수량 | 디폴트:0
					obj("spdLst")(null)("qckDvUseYn") = "N"										'퀵배송사용여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("crdayDvPsbYn") = "N"									'당일배송가능여부 [Y, N] | 디폴트:N
'					Set obj("spdLst")(null)("crdayDvInfo") = jsObject()							'o당일배송정보 | 당일배송가능여부가 Y인 경우 필수값
'						obj("spdLst")(null)("crdayDvInfo")("odCloseTm") = ""					'#주문마감시간 [HH24MI ex) 1000] | 당일배송가능여부가 Y인 경우 필수값
					obj("spdLst")(null)("spicUseYn") = "N"										'스마트픽사용여부 [Y, N] | 디폴트:N
					Set obj("spdLst")(null)("spicInfo") = jsObject()							'스마트픽정보 | 스마트픽사용여부 Y인 경우 필수
						obj("spdLst")(null)("spicInfo") = null
					obj("spdLst")(null)("hpDdDvPsbYn") = "N"									'희망일배송가능여부 [Y, N] 디폴트:N
'					obj("spdLst")(null)("hpDdDvPsbPrd") = ""									'희망일배송가능기간 | 희망일배송가능여부 Y인 경우 필수
					obj("spdLst")(null)("saveTypCd") = "NONE"									'저장유형코드 [공통코드 : SAVE_TYP_CD] | 디폴트:해당없음 RFRG : 냉장, FRZN : 냉동, FRSH : 신선, NONE : 해당없음
					obj("spdLst")(null)("rtngPsbYn") = "Y"										'반품가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("xchgPsbYn") = "Y"										'교환가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("echgPsbYn") = "N"										'맞교환가능여부 [Y, N] | 디폴트:N
					obj("spdLst")(null)("cmbnRtngPsbYn") = "Y"									'합반품가능여부 [Y, N] | 합배송가능여부가 Y인 경우 Y, N 선택 가능. N인 경우 N만 선택 가능
					obj("spdLst")(null)("rtngHdcCd") = ""										'반품택배사코드 | 0002 : 대한통운
					obj("spdLst")(null)("rtngRtrvPsbYn") = "Y"									'반품회수가능여부 [Y, N] | 디폴트:Y
					obj("spdLst")(null)("rtrpNo") = DVPCd(2)									'회수지번호 | 거래처 API "(일반 Seller용) 판매자 출고지/반품지 등록"을 통하여 등록된 회수지번호를 입력한다.
'					Set obj("spdLst")(null)("ecpnInfo") = jsObject()							'(생략)e쿠폰정보 | 해당 상품이 e쿠폰인 경우에만 입력한다.
'					Set obj("spdLst")(null)("rntlPdInfo") = jsObject()							'(생략)렌탈상품정보 | 상품유형이 렌탈일 경우 필수값
'					Set obj("spdLst")(null)("opngPdInfo") = jsObject()							'(생략)개통형상품정보 | 상품유형구분코드가 일반판매_0원상품(GNRL_ZRWON)에 해당하는 개통형상품인 경우 필수입력한다.
					obj("spdLst")(null)("stkMgtYn") = "Y"										'#재고관리여부 [Y, N] | 'N'인 경우 재고가 999,999,999로 들어간다. 웹재고를 관리하지 않는다.
					Call getLotteonOptionEditParameter(obj)										'단품목록
'					Set obj("spdLst")(null)("slrRcPdLst") = jsArray()							'셀러추천상품목록 | 최대 10개까지 등록 가능하다.
'						Set obj("spdLst")(null)("slrRcPdLst")(null) = jsObject()
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSpdNo") = ""			'#셀러추천판매자상품번호
'							obj("spdLst")(null)("slrRcPdLst")(null)("slrRcSitmNo") = ""			'#셀러추천판매자단품번호
'							obj("spdLst")(null)("slrRcPdLst")(null)("epsrPrirRnkg") = ""		'#노출우선순위
		getLotteonItemEditParameter = obj.jsString
'    response.write obj.jsString
'    response.end
	End Function

	'상품 상세조회 Json
	Public Function getLotteonItemViewParameter
		Dim strRst
		Dim obj
		Set obj = jsObject()
			obj("trGrpCd") = "SR"
			obj("trNo") = afflTrCd
			obj("spdNo") = FLotteonGoodNo
		getLotteonItemViewParameter = obj.jsString
	End Function

	'상품 재고수정 Json
	Public Function getLotteonQuantityParameter
		Dim strRst
		Dim obj, sqlStr, arrRows, limitsu

		Set obj = jsObject()
			Set obj("itmStkLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, o.optlimitno, o.optlimitsold "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " WHERE r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "단일상품") Then
				Set obj("itmStkLst")(null) = jsObject()
					obj("itmStkLst")(null)("trGrpCd") = "SR"
					obj("itmStkLst")(null)("trNo") = afflTrCd
					obj("itmStkLst")(null)("spdNo") = FLotteonGoodNo
					obj("itmStkLst")(null)("sitmNo") = arrRows(1,0)
					obj("itmStkLst")(null)("stkQty") = getLimitEa()
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						limitsu = ""
						If FLimityn = "Y" Then
							If arrRows(3, i) - arrRows(4, i) - 5 < 1 Then
								limitsu = 0
							Else
								limitsu = arrRows(3, i) - arrRows(4, i) - 5
							End If
						Else
							limitsu = CDEFALUT_STOCK
						End If

						Set obj("itmStkLst")(i) = jsObject()
							obj("itmStkLst")(i)("trGrpCd") = "SR"
							obj("itmStkLst")(i)("trNo") = afflTrCd
							obj("itmStkLst")(i)("spdNo") = FLotteonGoodNo
							obj("itmStkLst")(i)("sitmNo") = arrRows(1, i)
							obj("itmStkLst")(i)("stkQty") = limitsu
					Next
				End If
			End If
		getLotteonQuantityParameter = obj.jsString
	End Function

	'상품 가격수정 Json
	Public Function getLotteonPriceParameter
		Dim strRst
		Dim obj, sqlStr, arrRows
		Dim vMustprice
		vMustprice = mustPrice()

		Set obj = jsObject()
			Set obj("itmPrcLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, isnull(o.optAddPrice, 0) optAddPrice "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "단일상품") Then
				Set obj("itmPrcLst")(null) = jsObject()
					obj("itmPrcLst")(null)("trGrpCd") = "SR"
					obj("itmPrcLst")(null)("trNo") = afflTrCd
					obj("itmPrcLst")(null)("spdNo") = FLotteonGoodNo
					obj("itmPrcLst")(null)("sitmNo") = arrRows(1,0)
					obj("itmPrcLst")(null)("slPrc") = vMustprice
					obj("itmPrcLst")(null)("hstStrtDttm") = FormatDate(now(), "00000000000000")		'#가격시작일시
					obj("itmPrcLst")(null)("hstEndDttm") = "99991231235959"							'#가격종료일시
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						Set obj("itmPrcLst")(i) = jsObject()
							obj("itmPrcLst")(i)("trGrpCd") = "SR"
							obj("itmPrcLst")(i)("trNo") = afflTrCd
							obj("itmPrcLst")(i)("spdNo") = FLotteonGoodNo
							obj("itmPrcLst")(i)("sitmNo") = arrRows(1, i)
							obj("itmPrcLst")(i)("slPrc") = vMustprice + arrRows(3, i)						'#판매가
							obj("itmPrcLst")(i)("hstStrtDttm") = FormatDate(now(), "00000000000000")		'#가격시작일시
							obj("itmPrcLst")(i)("hstEndDttm") = "99991231235959"							'#가격종료일시
					Next
				End If
			End If
		getLotteonPriceParameter = obj.jsString
	End Function

	'상품 판매상태 변경 Json
	Public Function getLotteonSellynParameter(ichgSellYn)
		Dim strRst
		Dim obj, slStatCd
		Select Case ichgSellYn
			Case "Y"	slStatCd = "SALE"		'판매중
			Case "N"	slStatCd = "SOUT"		'품절
			Case "X"	slStatCd = "END"		'판매종료
		End Select

		Set obj = jsObject()
			Set obj("spdLst")= jsArray()
				Set obj("spdLst")(null) = jsObject()
					obj("spdLst")(null)("trGrpCd") = "SR"
					obj("spdLst")(null)("trNo") = afflTrCd
					obj("spdLst")(null)("spdNo") = FLotteonGoodNo
					obj("spdLst")(null)("slStatCd") = slStatCd
		getLotteonSellynParameter = obj.jsString
	End Function

	'상품 판매상태 변경 Json
	Public Function getLotteonOptStatusParameter()
		Dim strRst
		Dim obj, sqlStr, arrRows, optsellyn

		If rsget.state = "1" Then
			rsget.close
		End If

		Set obj = jsObject()
			Set obj("sitmLst")= jsArray()
			sqlStr = ""
			sqlStr = sqlStr & " SELECT isnull(o.itemoption, '') as itemoption, r.outmallOptCode, r.outmallOptName, isnull(o.optAddPrice, 0) optAddPrice "
			sqlStr = sqlStr & " , isnull(o.optionname, '') as optionname, isnull(o.isUsing, '') as isUsing, isnull(o.optsellyn, '') as optsellyn "
			sqlStr = sqlStr & " , (o.optlimitno - o.optlimitsold - 5) as optLimit "
			sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_OutMall_regedoption as r "
			sqlStr = sqlStr & " LEFT JOIN [db_item].[dbo].tbl_item_option as o on r.itemid = o.itemid and r.itemoption = o.itemoption "
			sqlStr = sqlStr & " where r.mallid = '"&CMALLNAME&"' and r.itemid="&Fitemid
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				arrRows = rsget.getRows
			End If
			rsget.close

			If (UBound(arrRows ,2) = "0") and (arrRows(2, 0) = "단일상품") Then
				optsellyn = Chkiif(FMaySoldOut="Y", "SOUT", "SALE")
				Set obj("sitmLst")(null) = jsObject()
					obj("sitmLst")(null)("trGrpCd") = "SR"
					obj("sitmLst")(null)("trNo") = afflTrCd
					obj("sitmLst")(null)("spdNo") = FLotteonGoodNo
					obj("sitmLst")(null)("sitmNo") = arrRows(1,0)
					obj("sitmLst")(null)("slStatCd") = optsellyn
			Else
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						optsellyn = ""
						'itempoption이 없다면(옵션삭제) / itemopti
						If (arrRows(0, i) = "") OR (arrRows(5, i) <> "Y") OR (arrRows(6, i) <> "Y") THEN
							optsellyn = "SOUT"
						ElseIf FLimityn = "Y" AND (arrRows(7, i) < 1) Then
							optsellyn = "SOUT"
						Else
							optsellyn = "SALE"
						End If

						Set obj("sitmLst")(i) = jsObject()
							obj("sitmLst")(i)("trGrpCd") = "SR"
							obj("sitmLst")(i)("trNo") = afflTrCd
							obj("sitmLst")(i)("spdNo") = FLotteonGoodNo
							obj("sitmLst")(i)("sitmNo") = arrRows(1, i)
							obj("sitmLst")(i)("slStatCd") = optsellyn
					Next
				End If
			End If
		getLotteonOptStatusParameter = obj.jsString
	End Function

End Class

Class CLotteon
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
	Public Sub getLotteonNotRegOneItem
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
		strSql = strSql & " , (SELECT db_etcmall.dbo.getOutmallKeywords ('"& CMALLNAME &"', i.itemid) ) as keywords "
		strSql = strSql & "	, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & "	, isNULL(R.lotteonStatCD,-9) as lotteonStatCD, IsNull(R.lotteonPrice, 0) as lotteonPrice "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & " , am.std_cat_id, am.disp_cat_id "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_lotteon_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
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
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "				'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold>"&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_lotteon_regItem WHERE lotteonStatCD >= 3) "	''등록완료이상은 등록안됨.	'lotteon등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "'	카테고리 매칭 상품만
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteonItem
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
				FOneItem.FLotteonStatCD		= rsget("lotteonStatCD")
				FOneItem.FLotteonPrice		= rsget("lotteonPrice")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FStd_cat_id		= rsget("std_cat_id")
				FOneItem.FDisp_cat_id		= rsget("disp_cat_id")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getLotteonNotEditOneItem
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
		strSql = strSql & "	, m.lotteonGoodNo, m.lotteonprice, m.lotteonSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, isnull(m.lastStatCheckDate, '1900-01-01 00:00:00.000') as lastStatCheckDate "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.std_cat_id, am.disp_cat_id, isNULL(m.lotteonStatCD,-9) as lotteonStatCD "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_lotteon_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lotteon_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.lotteonGoodno is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLotteonItem
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FLotteonGoodNo		= rsget("lotteonGoodNo")
				FOneItem.FLotteonprice		= rsget("lotteonprice")
				FOneItem.FLotteonSellYn		= rsget("lotteonSellYn")

				FOneItem.FMakerName			= db2html(rsget("makername"))
				FOneItem.FBrandName			= db2html(rsget("brandname"))
				FOneItem.FBrandNameKor		= db2html(rsget("socname_kor"))
				If (IsNULL(FOneItem.FMakerName) or (FOneItem.FMakerName="")) Then
					FOneItem.FMakerName		= FOneItem.FBrandName
				End If

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
                FOneItem.FStd_cat_id		= rsget("std_cat_id")
				FOneItem.FDisp_cat_id		= rsget("disp_cat_id")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FLotteonStatCD		= rsget("lotteonStatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FLastStatCheckDate = rsget("lastStatCheckDate")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

End Class

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

function APIURL()
	If application("Svr_Info") = "Dev" Then
		APIURL = "https://dev-openapi.lotteon.com"
	Else
		APIURL = "https://openapi.lotteon.com"
	End If
end function

function APIkey()
	If application("Svr_Info") = "Dev" Then
		APIkey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
	Else
		APIkey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
	End If
end function

function afflTrCd()
	If application("Svr_Info") = "Dev" Then
		afflTrCd = "LO10001101"
	Else
		afflTrCd = "LD304013"
	End If
end function

function DVPCd(v)
	'v : 0(배송비 정책), 1(출고지), 2(회수지), 3(추가배송비)
	If v = "0" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1000529"
		Else
			DVPCd = "DLD706463"
		End If
	ElseIf v = "1" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1300153"
		Else
			DVPCd = "BPLD304013"
		End If
	ElseIf v = "2" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "1300153"
		Else
			DVPCd = "PLD333127"
		End If
	ElseIf v = "3" Then
		If application("Svr_Info") = "Dev" Then
			DVPCd = "2009166"
		Else
			DVPCd = "2009166"
		End If
	End If
end function
%>