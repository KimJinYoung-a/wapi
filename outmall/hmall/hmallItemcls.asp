<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "hmall1010"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST HMALLMARGIN = 11
CONST CDEFALUT_STOCK = 9999

Class CHmallItem
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
	Public FHmallRegdate
	Public FHmallLastUpdate
	Public FHmallGoodNo
	Public FHmallPrice
	Public FoctyCnryGbcd
	Public FoctyCnryNm
	Public FitemLCsfCd
	Public FitemMCsfCd
	Public FitemSCsfCd
	Public FitemCsfGbcd
	Public Fitemsize	
	Public Fitemsource
	Public FHmallSellYn
	Public FMrgnRate
	Public FregUserid
	Public FHmallStatCd
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
	Public FordMakeYn
	Public ForderComment
	Public FAdultType
	Public FbasicImage
	Public FbasicimageNm
	Public FmainImage
	Public FmainImage2
	Public Ficon2Image
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
		limitYCnt = 0
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
        buf = LeftB(buf, 140)
        getItemNameFormat = buf
    end function

	Public Function getSafetyParam()
		Dim strSql, isCertYn, safeCertTypeGbcd, safetyDiv, gbnFlag
		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, t.certNum "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " WHERE i.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			isCertYn	= "Y"
			safetyDiv	= rsget("safetyDiv")
		Else
			isCertYn	= "N"
			safeCertTypeGbcd = "N"
		End If
		rsget.Close

		If isCertYn = "Y" Then
			Select Case safetyDiv
				Case "10", "40", "70"		safeCertTypeGbcd = "01"
				Case "20", "50", "80"		safeCertTypeGbcd = "02"
				Case "30", "60", "90"		safeCertTypeGbcd = "03"
			End Select

			Select Case safetyDiv
				Case "10", "20", "30"		gbnFlag = "elec"
				Case "40", "50", "60"		gbnFlag = "life"
				Case "70", "80", "90"		gbnFlag = "child"
			End Select
		End If
		getSafetyParam = isCertYn&"|_|"&safeCertTypeGbcd&"|_|"&gbnFlag
	End Function

	Public Function IsAllOptionChange
		Dim sqlStr, tenOptCnt, regedHmallOptCnt, addPriceCnt
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
		sqlStr = sqlStr & " db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and optaddprice > 0 "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			addPriceCnt = rsget("cnt")
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " select count(*) as cnt from "
		sqlStr = sqlStr & " db_etcmall.[dbo].[tbl_hmall_regedOption]  "
		sqlStr = sqlStr & " where itemid = '"&FItemid&"' "
		sqlStr = sqlStr & " and outmallOptName <> '단일옵션' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			regedHmallOptCnt = rsget("cnt")
		rsget.Close

		If tenOptCnt > 0 AND regedHmallOptCnt = 0 Then			'단품 -> 옵션
			IsAllOptionChange = "Y"
		ElseIf tenOptCnt = 0 AND regedHmallOptCnt > 0 Then		'옵션 -> 단품
			IsAllOptionChange = "Y"
		ElseIf addPriceCnt > 0 Then								'옵션추가금액 없다가 생긴 경우
			IsAllOptionChange = "Y"
		Else
			IsAllOptionChange = "N"
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

	'간이 과세 브랜드 여부 체크
	Public Function fnCheckMakerid()
		Dim strSql, cntMakerId

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_partner.dbo.tbl_partner_group G "
		strSql = strSql & " JOIN db_partner.dbo.tbl_partner as P on G.groupid = p.groupid "
		strSql = strSql & " WHERE G.jungsan_gubun = '간이과세' "
		strSql = strSql & " and p.id = '"& FMakerId &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.EOF Then
			cntMakerId = rsget("cnt")
		End If
		rsget.Close

		If cntMakerId > 0 Then
			fnCheckMakerid = True
		Else
			fnCheckMakerid = False
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

	'// hmall 판매여부 반환
	Public Function gethmallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				gethmallSellYn = "Y"
			Else
				gethmallSellYn = "N"
			End If
		Else
			gethmallSellYn = "N"
		End If
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

	Function getOptionLimitEa(ino, isold)
		dim ret : ret = (ino - isold - 5)
		if (ret < 1) then ret=0
		If (ret >= 1000) Then ret = 999
		getOptionLimitEa = ret
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
			If (GetTenTenMargin < CMAXMARGIN) Then
				MustPrice = CStr(GetRaiseValue(Forgprice/10)*10)
			Else
				If (FSellCash < Round(FHmallPrice * 0.55, 0)) Then
					MustPrice = CStr(GetRaiseValue(Round(FHmallPrice * 0.55, 0)/10)*10)
				Else
					MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
				End If
			End If
		End If
	End Function

	Public Function getBasicImage()
		If IsNULL(FbasicImageNm) or (FbasicImageNm="") Then Exit function
		getBasicImage = FbasicImageNm
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

	Public Function getHmallContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_hmall.jpg'></p><br>"

		If ForderComment <> "" Then
			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
		End If

		If Fitemsize <> "" Then
			strRst = strRst & "- 사이즈 : " & Fitemsize & "<br>"
		End if

		If Fitemsource <> "" Then
			strRst = strRst & "- 재료 : " &  Fitemsource & "<br>"
		End If
		strRst = strRst & Replace(Replace(FItemContent,"",""),"","")

		'# 추가 상품 설명이미지 접수
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
				If rsget("imgType") = "1" Then
					strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400") & """><br>")
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		'#기본 상품 설명이미지
		If ImageExists(FmainImage) Then strRst = strRst & ("<img src=""" & FmainImage & """ ><br>")
		If ImageExists(FmainImage2) Then strRst = strRst & ("<img src=""" & FmainImage2 & """ ><br>")

		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_hmall.jpg"">")
		getHmallContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = ""
			strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_hmall.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_hmall.jpg"">")
			getHmallContParamToReg = strRst
		End If
		rsget.Close
	End Function

	Public Function getAttrInfo()
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		chkMultiOpt = false

		If FoptionCnt = 0 Then
			strRst = ""
			strRst = strRst & "					<uitmCombYn>N</uitmCombYn>"									'상품속성조합여부
			strRst = strRst & "					<uitm1AttrTypeNm><![CDATA[단일옵션]]></uitm1AttrTypeNm>"	'상품속성1속성유형명
			strRst = strRst & "					<uitm2AttrTypeNm></uitm2AttrTypeNm>"						'상품속성2속성유형명
			strRst = strRst & "					<uitm3AttrTypeNm></uitm3AttrTypeNm>"						'상품속성3속성유형명
			strRst = strRst & "					<uitm4AttrTypeNm></uitm4AttrTypeNm>"						'상품속성4속성유형명
			strRst = strRst & "					<uitmChocPossYn>N</uitmChocPossYn>"							'상품속성선택가능여부
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			If chkMultiOpt = True Then
				optionTypeName1 = ""
				optionTypeName2 = ""
				optionTypeName3 = ""
				optionTypeName4 = ""
				strSql = ""
				strSql = strSql & " SELECT typeseq, optionTypeName From db_item.[dbo].[tbl_item_option_Multiple] "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " GROUP BY typeseq, optionTypeName "
				strSql = strSql & " ORDER BY Typeseq "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.Eof Then
					Do until rsget.EOF
						Select Case rsget("typeseq")
							Case "1"		optionTypeName1 = rsget("optionTypeName")
							Case "2"		optionTypeName2 = rsget("optionTypeName")
							Case "3"		optionTypeName3 = rsget("optionTypeName")
							Case "4"		optionTypeName4 = rsget("optionTypeName")
						End Select
						rsget.MoveNext
					Loop
				End If
				rsget.close

				strRst = ""
				strRst = strRst & "					<uitmCombYn>Y</uitmCombYn>"									'상품속성조합여부
				strRst = strRst & "					<uitm1AttrTypeNm><![CDATA["&optionTypeName1&"]]></uitm1AttrTypeNm>"	'상품속성1속성유형명
				strRst = strRst & "					<uitm2AttrTypeNm><![CDATA["&optionTypeName2&"]]></uitm2AttrTypeNm>"						'상품속성2속성유형명
				strRst = strRst & "					<uitm3AttrTypeNm><![CDATA["&optionTypeName3&"]]></uitm3AttrTypeNm>"						'상품속성3속성유형명
				strRst = strRst & "					<uitm4AttrTypeNm><![CDATA["&optionTypeName4&"]]></uitm4AttrTypeNm>"						'상품속성4속성유형명
				strRst = strRst & "					<uitmChocPossYn>Y</uitmChocPossYn>"							'상품속성선택가능여부
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 optionTypeName "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " GROUP BY optionTypeName "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					strRst = ""
					strRst = strRst & "					<uitmCombYn>Y</uitmCombYn>"									'상품속성조합여부
					strRst = strRst & "					<uitm1AttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitm1AttrTypeNm>"	'상품속성1속성유형명
					strRst = strRst & "					<uitm2AttrTypeNm></uitm2AttrTypeNm>"						'상품속성2속성유형명
					strRst = strRst & "					<uitm3AttrTypeNm></uitm3AttrTypeNm>"						'상품속성3속성유형명
					strRst = strRst & "					<uitm4AttrTypeNm></uitm4AttrTypeNm>"						'상품속성4속성유형명
					strRst = strRst & "					<uitmChocPossYn>Y</uitmChocPossYn>"							'상품속성선택가능여부
				End If
				rsget.Close
			End If
		End If
		getAttrInfo = strRst
	End Function

	Public Function getOptSellUitmDtl
		Dim strSql, strRst, i, chkMultiOpt, j, chkqty
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, commaCount
		chkqty = ""
		chkMultiOpt = false

		If FoptionCnt = 0 Then
			buf = ""
			buf = buf & "	<Dataset id=""dsSellUitmDtl"">"											'#판매상품속성내역
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>INSERT</rowType>"
			buf = buf & "				<chk>1</chk>"												'속성존재여부 | 단품인경우만 사용 , Default :1  (단품등록이 제대로 안될때만 사용하세요, 일반적으로 chk없이도 단품 등록은 정상적으로 됩니다.)
			buf = buf & "				<bsitmCd></bsitmCd>"										'기준상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요), null 값
			buf = buf & "				<slitmCd></slitmCd>"										'상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , null 값
			buf = buf & "				<uitmCd></uitmCd>"											'상품속성코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , 속성코드(uitmTmpCd)값 입력
			buf = buf & "				<uitmTmpCd>0</uitmTmpCd>"									'#속성코드 | 단품/조합인경우 사용
			buf = buf & "				<uitm1AttrNm><![CDATA[단일옵션]]></uitm1AttrNm>"			'#상품속성1속성유형명 | 단품/조합인경우 사용
			buf = buf & "				<uitm2AttrNm><![CDATA[단일옵션]]></uitm2AttrNm>"			'#상품속성2속성유형명 | 단품/조합인경우 사용
			buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#상품속성3속성유형명 | 단품/조합인경우 사용
			buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#상품속성4속성유형명 | 단품/조합인경우 사용
			buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#판매시작일자 | 단품/조합인경우 사용
			buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#판매종료일자 | 단품/조합인경우 사용
			buf = buf & "				<uitmTotNm><![CDATA[단일옵션]]></uitmTotNm>"				'#상품속성전체명 | 단품/조합인경우 사용
			buf = buf & "				<addQty>0</addQty>"											'#추가수량 |	단품/조합인경우 사용
			buf = buf & "				<maxSellPossQty>"&getLimitEa()&"</maxSellPossQty>"			'#최대판매가능수량 | 단품/조합인경우 사용
			buf = buf & "				<sellGbcd>00</sellGbcd>"									'#판매구분코드 | 00 진행 <- (등록인경우 기본값으로 셋팅요망), 11 일시중단, 19 영구중단
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			buf = ""
			buf = buf & "	<Dataset id=""dsSellUitmDtl"">"											'#판매상품속성내역
			buf = buf & "		<rows>"

			If chkMultiOpt = True Then
				j = 0

				strSql = ""
				strSql = strSql & " SELECT optionname, optlimitno, optlimitsold "
				strSql = strSql & " FROM db_item.dbo.tbl_item_option "
				strSql = strSql & " WHERE itemid = " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						commaCount = Ubound(Split(("optionname"), ","))
						buf = buf & "			<row>"
						buf = buf & "				<rowType>INSERT</rowType>"
						buf = buf & "				<chk>1</chk>"												'속성존재여부 | 단품인경우만 사용 , Default :1  (단품등록이 제대로 안될때만 사용하세요, 일반적으로 chk없이도 단품 등록은 정상적으로 됩니다.)
						buf = buf & "				<bsitmCd></bsitmCd>"										'기준상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요), null 값
						buf = buf & "				<slitmCd></slitmCd>"										'상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , null 값
						buf = buf & "				<uitmCd></uitmCd>"											'상품속성코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , 속성코드(uitmTmpCd)값 입력
						buf = buf & "				<uitmTmpCd>"&j&"</uitmTmpCd>"								'#속성코드 | 단품/조합인경우 사용

						buf = buf & "				<uitm1AttrNm><![CDATA["&Split(rsget("optionname"), ",")(0)&"]]></uitm1AttrNm>"			'#상품속성1속성유형명 | 단품/조합인경우 사용
						buf = buf & "				<uitm2AttrNm><![CDATA["&Split(rsget("optionname"), ",")(1)&"]]></uitm2AttrNm>"			'#상품속성2속성유형명 | 단품/조합인경우 사용
						If Ubound(Split(rsget("optionname"), ",")) = 2 Then
							buf = buf & "				<uitm3AttrNm><![CDATA["&Split(rsget("optionname"), ",")(2)&"]]></uitm3AttrNm>"								'#상품속성3속성유형명 | 단품/조합인경우 사용
						Else
							buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#상품속성3속성유형명 | 단품/조합인경우 사용
						End If

						If Ubound(Split(rsget("optionname"), ",")) = 3 Then
							buf = buf & "				<uitm4AttrNm><![CDATA["&Split(rsget("optionname"), ",")(3)&"]]></uitm4AttrNm>"								'#상품속성3속성유형명 | 단품/조합인경우 사용
						Else
							buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#상품속성4속성유형명 | 단품/조합인경우 사용
						End If
						buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#판매시작일자 | 단품/조합인경우 사용
						buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#판매종료일자 | 단품/조합인경우 사용
						buf = buf & "				<uitmTotNm><![CDATA["& Replace(rsget("optionname"), ",", "/") &"]]></uitmTotNm>"				'#상품속성전체명 | 단품/조합인경우 사용
						buf = buf & "				<addQty>0</addQty>"											'#추가수량 |	단품/조합인경우 사용
						If FLimityn = "Y" Then
							If rsget("optlimitno") - rsget("optlimitsold") - 5 < 0 Then
								chkqty = 0
							Else
								chkqty = rsget("optlimitno") - rsget("optlimitsold") - 5
							End If
							buf = buf & "				<maxSellPossQty>"&chkqty&"</maxSellPossQty>"			'#최대판매가능수량 | 단품/조합인경우 사용
						Else
							buf = buf & "				<maxSellPossQty>9999</maxSellPossQty>"						'#최대판매가능수량 | 단품/조합인경우 사용
						End If
						buf = buf & "				<sellGbcd>00</sellGbcd>"									'#판매구분코드 | 00 진행 <- (등록인경우 기본값으로 셋팅요망), 11 일시중단, 19 영구중단
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				rsget.close
			Else
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						buf = buf & "			<row>"
						buf = buf & "				<rowType>INSERT</rowType>"
						buf = buf & "				<chk>1</chk>"												'속성존재여부 | 단품인경우만 사용 , Default :1  (단품등록이 제대로 안될때만 사용하세요, 일반적으로 chk없이도 단품 등록은 정상적으로 됩니다.)
						buf = buf & "				<bsitmCd></bsitmCd>"										'기준상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요), null 값
						buf = buf & "				<slitmCd></slitmCd>"										'상품코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , null 값
						buf = buf & "				<uitmCd></uitmCd>"											'상품속성코드 | 단품인경우만 사용(단품등록이 제대로 안될때만 사용하세요) , 속성코드(uitmTmpCd)값 입력
						buf = buf & "				<uitmTmpCd>"&j&"</uitmTmpCd>"								'#속성코드 | 단품/조합인경우 사용
						buf = buf & "				<uitm1AttrNm><![CDATA["&rsget("typename")&"]]></uitm1AttrNm>"			'#상품속성1속성유형명 | 단품/조합인경우 사용
						buf = buf & "				<uitm2AttrNm><![CDATA["&rsget("kindname")&"]]></uitm2AttrNm>"			'#상품속성2속성유형명 | 단품/조합인경우 사용
						buf = buf & "				<uitm3AttrNm></uitm3AttrNm>"								'#상품속성3속성유형명 | 단품/조합인경우 사용
						buf = buf & "				<uitm4AttrNm></uitm4AttrNm>"								'#상품속성4속성유형명 | 단품/조합인경우 사용
						buf = buf & "				<sellStrtDt>"&Replace(Date(), "-", "")&"</sellStrtDt>"		'#판매시작일자 | 단품/조합인경우 사용
						buf = buf & "				<sellEndDt>"&Replace(DateAdd("yyyy", 5, DATE()), "-", "")&"</sellEndDt>"	'#판매종료일자 | 단품/조합인경우 사용
						buf = buf & "				<uitmTotNm><![CDATA["&rsget("typename")&"/"&rsget("kindname")&"]]></uitmTotNm>"				'#상품속성전체명 | 단품/조합인경우 사용
						buf = buf & "				<addQty>0</addQty>"											'#추가수량 |	단품/조합인경우 사용
						If FLimityn = "Y" Then
							If rsget("limitno") - rsget("limitsold") - 5 < 0 Then
								chkqty = 0
							Else
								chkqty = rsget("limitno") - rsget("limitsold") - 5
							End If
							buf = buf & "				<maxSellPossQty>"&chkqty&"</maxSellPossQty>"			'#최대판매가능수량 | 단품/조합인경우 사용
						Else
							buf = buf & "				<maxSellPossQty>9999</maxSellPossQty>"						'#최대판매가능수량 | 단품/조합인경우 사용
						End If
						buf = buf & "				<sellGbcd>00</sellGbcd>"									'#판매구분코드 | 00 진행 <- (등록인경우 기본값으로 셋팅요망), 11 일시중단, 19 영구중단
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				rsget.close
			End If
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptSellUitmDtl = buf
	End Function

	Public Function getOptTypeMst(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false
		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'상품속성속성유형순번
			buf = buf & "				<uitmAttrTypeNm><![CDATA[단일옵션]]></uitmAttrTypeNm>"							'상품속성속성유형명
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			If chkMultiOpt = True Then
				buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
				buf = buf & "		<rows>"
				strSql = ""
				strSql = strSql & " SELECT typeseq, optionTypeName From db_item.[dbo].[tbl_item_option_Multiple] "
				strSql = strSql & " WHERE itemid = " & FItemid
				strSql = strSql & " GROUP BY typeseq, optionTypeName "
				strSql = strSql & " ORDER BY Typeseq "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.Eof Then
					Do until rsget.EOF
						Select Case rsget("typeseq")
							Case "1"		optionTypeName1 = rsget("optionTypeName")
							Case "2"		optionTypeName2 = rsget("optionTypeName")
							Case "3"		optionTypeName3 = rsget("optionTypeName")
							Case "4"		optionTypeName4 = rsget("optionTypeName")
						End Select
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmAttrTypeSeq>"&rsget("typeseq")&"</uitmAttrTypeSeq>"	'상품속성속성유형순번
						buf = buf & "				<uitmAttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitmAttrTypeNm>"							'상품속성속성유형명
						buf = buf & "			</row>"
						rsget.MoveNext
					Loop
				End If
				rsget.close
				buf = buf & "		</rows>"
				buf = buf & "	</Dataset>"
			Else
				strSql = ""
				strSql = strSql & " SELECT TOP 1 optionTypeName "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & FItemid
				strSql = strSql & " GROUP BY optionTypeName "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not(rsget.EOF or rsget.BOF) Then
					buf = buf & "	<Dataset id=""dsUitmAttrTypeMst"">"
					buf = buf & "		<rows>"
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'상품속성속성유형순번
					buf = buf & "				<uitmAttrTypeNm><![CDATA["&rsget("optionTypeName")&"]]></uitmAttrTypeNm>"							'상품속성속성유형명
					buf = buf & "			</row>"
					buf = buf & "		</rows>"
					buf = buf & "	</Dataset>"
				End If
				rsget.Close
			End If
		End If
		getOptTypeMst = buf
	End Function

	Public Function getOptAttrMst(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, j, rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false

		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsUitmAttrMst"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmTmpSeq>0</uitmTmpSeq>"									'속성순번
			buf = buf & "				<uitmAttrTypeSeq>1</uitmAttrTypeSeq>"						'상품속성속성유형순번
			buf = buf & "				<uitmAttrNm></uitmAttrNm>"									'상품속성명
			buf = buf & "				<uitmCreYn></uitmCreYn>"									'상품속성생성여부
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close

			buf = buf & "	<Dataset id=""dsUitmAttrMst"">"
			buf = buf & "		<rows>"
			j = 0
			strSql = ""
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
			rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
			rsget.Open strSql, dbget
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<uitmTmpSeq>"&j&"</uitmTmpSeq>"								'속성순번
					buf = buf & "				<uitmAttrTypeSeq>"&rsget("TypeSeq")&"</uitmAttrTypeSeq>"	'상품속성속성유형순번
					buf = buf & "				<uitmAttrNm><![CDATA["&rsget("kindname")&"]]></uitmAttrNm>"	'상품속성명
					buf = buf & "				<uitmCreYn></uitmCreYn>"									'상품속성생성여부
					buf = buf & "			</row>"
					j = j + 1
					rsget.MoveNext
				Loop
			End If
			rsget.Close
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptAttrMst = buf
	End Function

	Public Function getOptCombDtl(gbn)
		Dim strSql, strRst, i, chkMultiOpt
		Dim optionTypeName1, optionTypeName2, optionTypeName3, optionTypeName4
		Dim buf, j, rowType, tmpOption

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		chkMultiOpt = false

		buf = ""
		If FoptionCnt = 0 Then
			buf = buf & "	<Dataset id=""dsSellUitmCombDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<uitmTmpCd>0</uitmTmpCd>"									'속성코드
			buf = buf & "				<uitmTmpSeq>0</uitmTmpSeq>"									'속성순번
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql, dbget
		    if not rsget.Eof then
		    	chkMultiOpt = True
		    end if
		    rsget.close
			buf = buf & "	<Dataset id=""dsSellUitmCombDtl"">"
			buf = buf & "		<rows>"
			If chkMultiOpt = True Then
				'멀티옵션일 때 해야함..........
				i = 0
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get2] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						If j = 0 Then
							tmpOption = rsget("itemoption")
						End If

						If tmpOption <> rsget("itemoption") Then
							tmpOption = rsget("itemoption")
							i = i + 1
						End If
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmTmpCd>"&i&"</uitmTmpCd>"												'속성코드
						buf = buf & "				<uitmTmpSeq>"&rsget("rnum")&"</uitmTmpSeq>"									'속성순번
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				 rsget.close
			Else
				j = 0
				strSql = ""
				strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_OptionAttr_Get] " & FItemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType = adOpenStatic
				rsget.LockType = adLockOptimistic
				rsget.Open strSql, dbget
				If Not(rsget.EOF or rsget.BOF) then
					Do until rsget.EOF
						buf = buf & "			<row>"
						buf = buf & "				<rowType>"&rowType&"</rowType>"
						buf = buf & "				<uitmTmpCd>"&rsget("rnum")&"</uitmTmpCd>"									'속성코드
						buf = buf & "				<uitmTmpSeq>"&rsget("rnum")&"</uitmTmpSeq>"									'속성순번
						buf = buf & "			</row>"
						j = j + 1
						rsget.MoveNext
					Loop
				End If
				 rsget.close
			End If
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getOptCombDtl = buf
	End Function

    Public Function getCertOrganName(icertOrganName)
		Select Case icertOrganName
			Case Instr(icertOrganName, "FITI시험연구원") > 0				getCertOrganName = "8"
			Case Instr(icertOrganName, "한국화학융합시험연구원") > 0		getCertOrganName = "4"
			Case Instr(icertOrganName, "한국기계전기전자시험연구원") > 0	getCertOrganName = "10"
			Case Instr(icertOrganName, "KOTITI 시험연구원") > 0				getCertOrganName = "14"
			Case Instr(icertOrganName, "한국건설생활시험연구원") > 0		getCertOrganName = "5"
			Case Instr(icertOrganName, "한국의류시험연구원") > 0			getCertOrganName = "7"
			Case Instr(icertOrganName, "한국산업기술시험원") > 0			getCertOrganName = "3"
			Case Instr(icertOrganName, "한국건설생활환경시험연구원") > 0	getCertOrganName = "5"
		End Select
    End function

	Public Function getHmallItemSafeInfoToReg(gbcd, gbn)
		Dim buf
		Dim strSql, safetyDiv, certNum, certOrganName, modelName, certDate
		Dim safeCertLawGbcd, safeCertTypeGbcd, safeCertNo, safeCrtiGbcd, speCate
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		speCate = "N"
		strSql = ""
		strSql = strSql & " SELECT TOP 1 i.itemid, t.safetyDiv, isNull(t.certNum, '') as certNum, isNull(f.modelName, '') as modelName, isNull(f.certDate, '') as certDate, isNull(f.certOrganName, '') as certOrganName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " LEFT JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv	= rsget("safetyDiv")
			certNum		= rsget("certNum")
			certOrganName = rsget("certOrganName")
			modelName	= rsget("modelName")
			certDate	= rsget("certDate")
		End If
		rsget.Close

		Select Case safetyDiv
			Case "10"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "01"
				safeCertNo				= certNum
			Case "20"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "30"
				safeCertLawGbcd			= "20"
				safeCertTypeGbcd		= "03"
			Case "40"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "02"				'생활화학은 안전확인만됨..
				safeCertNo				= certNum
			Case "50"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "60"
				safeCertLawGbcd			= "30"
				safeCertTypeGbcd		= "03"
			Case "70"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "01"
				safeCertNo				= certNum
			Case "80"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "02"
				safeCertNo				= certNum
			Case "90"
				safeCertLawGbcd			= "10"
				safeCertTypeGbcd		= "03"

		End Select
		safeCrtiGbcd = getCertOrganName(certOrganName)

		Select Case FitemLCsfCd
			Case "R6"	
				safeCertLawGbcd = "10"
				speCate =  "Y"
		End Select

		If speCate="Y" and gbcd = "03" Then
			buf = ""
			buf = buf & "	<Dataset id=""dsItemSafeCertMngDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<safeCertLawGbcd>10</safeCertLawGbcd>"		'안전인증법구분코드 | 10 어린이안전특별법, 20 전기안전법, 30 생활화학안전법, 40 방송통신기자재법
			buf = buf & "				<safeCertTypeGbcd>03</safeCertTypeGbcd>"	'안전인증유형구분코드 | 01 안전인증, 02 안전확인, 03 공급자적합성확인, 04 안전기준준수대상, 05 적합인증, 06 적합등록, 07 잠정인증
			buf = buf & "				<safeCertDt></safeCertDt>"						'안전인증일자 | YYYYMMDD 형태
			buf = buf & "				<safeCertNo></safeCertNo>"					'안전인증번호
			buf = buf & "				<safeCrtiGbcd></safeCrtiGbcd>"				'안전인증기관구분코드 | 1 한국전기전자시험연구원, 2 한국전자파연구원, 3 한국산업기술시험원, 4 한국화학시험연구원, 5 한국생활환경시험연구원, 6 한국기기유화시험연구원, 7 한국의류시험연구원, 8 FITI시험연구원, 9 한국표준협회, 10 한국기계전기전자시험연구원, 11 식약청, 12 방송통신위원회, 13 국립전파연구회, 14 KOTITI 시험연구원, 
			buf = buf & "				<safeCertClasGbcd>1</safeCertClasGbcd>"						'안전인증항목구분코드 | 0 없음, 1 국가통합코드(KC), 2 공산품안전검사, 3 안전인증대상공산품, 4 자율안전확인대상공산품, 5 전기용품안전인증, 6 의료기기제조품목허가, 7 안전확인대상전기용품, 8 위해우려제품
			buf = buf & "				<safeCertImgNm></safeCertImgNm>"							'안전인증이미지명
			buf = buf & "				<certInfIdnfYn>Y</certInfIdnfYn>"							'인증정보확인여부 | Y 동의, N 비동의
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		Else
			buf = ""
			buf = buf & "	<Dataset id=""dsItemSafeCertMngDtl"">"
			buf = buf & "		<rows>"
			buf = buf & "			<row>"
			buf = buf & "				<rowType>"&rowType&"</rowType>"
			buf = buf & "				<safeCertLawGbcd>"&safeCertLawGbcd&"</safeCertLawGbcd>"		'안전인증법구분코드 | 10 어린이안전특별법, 20 전기안전법, 30 생활화학안전법, 40 방송통신기자재법
			buf = buf & "				<safeCertTypeGbcd>"&safeCertTypeGbcd&"</safeCertTypeGbcd>"	'안전인증유형구분코드 | 01 안전인증, 02 안전확인, 03 공급자적합성확인, 04 안전기준준수대상, 05 적합인증, 06 적합등록, 07 잠정인증
			buf = buf & "				<safeCertDt>"&certDate&"</safeCertDt>"						'안전인증일자 | YYYYMMDD 형태
			buf = buf & "				<safeCertNo>"&safeCertNo&"</safeCertNo>"					'안전인증번호
			buf = buf & "				<safeCrtiGbcd>"&safeCrtiGbcd&"</safeCrtiGbcd>"				'안전인증기관구분코드 | 1 한국전기전자시험연구원, 2 한국전자파연구원, 3 한국산업기술시험원, 4 한국화학시험연구원, 5 한국생활환경시험연구원, 6 한국기기유화시험연구원, 7 한국의류시험연구원, 8 FITI시험연구원, 9 한국표준협회, 10 한국기계전기전자시험연구원, 11 식약청, 12 방송통신위원회, 13 국립전파연구회, 14 KOTITI 시험연구원, 
			buf = buf & "				<safeCertClasGbcd>1</safeCertClasGbcd>"						'안전인증항목구분코드 | 0 없음, 1 국가통합코드(KC), 2 공산품안전검사, 3 안전인증대상공산품, 4 자율안전확인대상공산품, 5 전기용품안전인증, 6 의료기기제조품목허가, 7 안전확인대상전기용품, 8 위해우려제품
			buf = buf & "				<safeCertImgNm></safeCertImgNm>"							'안전인증이미지명
			buf = buf & "				<certInfIdnfYn></certInfIdnfYn>"							'인증정보확인여부 | Y 동의, N 비동의
			buf = buf & "			</row>"
			buf = buf & "		</rows>"
			buf = buf & "	</Dataset>"
		End If
		getHmallItemSafeInfoToReg = buf
	End Function

	Function getHmallItemInfoCdToReg(gbn)
		Dim strSql, buf
		Dim mallinfoCd,infoContent,infotype
		Dim itstCd, itstGbcd, itstTitl, itstCntn
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		buf = ""
		buf = buf & "	<Dataset id=""dsItstDtl"">"
		buf = buf & "		<rows>"

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
			    itstCd		= rsget("itstCd")
			    itstGbcd	= rsget("itstGbcd")
			    itstTitl	= rsget("itstTitl")
				itstCntn	= rsget("itstCntn")

			    If Not (IsNULL(itstCntn)) AND (itstCntn <> "") Then
			    	itstCntn = replace(itstCntn, chr(31), "")
				End If

				buf = buf & "			<row>"
				buf = buf & "				<rowType>"&rowType&"</rowType>"
				buf = buf & "				<itstCd>"&itstCd&"</itstCd>"											'상품기술서코드
				buf = buf & "				<itstGbcd>"&itstGbcd&"</itstGbcd>"										'상품기술서구분코드 | 10 정보고시용, 20 상품기술서용 (상품기술서용항목은 상품필수정보 조회 API를 이용하여 조회한다.)
				buf = buf & "				<itstTitl><![CDATA["&itstTitl&"]]></itstTitl>"							'상품기술서제목	String(200)	
				buf = buf & "				<itstCntn><![CDATA["&itstCntn&"]]></itstCntn>"							'상품기술서내용	String(4000)	
				buf = buf & "			</row>"
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		buf = buf & "		</rows>"
		buf = buf & "	</Dataset>"
		getHmallItemInfoCdToReg = buf
	End Function

	Public Function getHmallSectIdToReg(gbn)
		Dim buf, strSql
		Dim sectAttrGbcd, sectId1, sectId2
		Dim rowType

		Select Case gbn
			Case "I"			rowType = "INSERT"
			Case "U"			rowType = "UPDATE"
		End Select

		buf = ""
		buf = buf & "	<Dataset id=""dsDispItemDtl"">"
		buf = buf & "		<rows>"

		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_SpecialCategoryMapping_Get] '"& FtenCateLarge &"', '"& FtenCateMid &"' "
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				sectAttrGbcd = ""
				sectId1 = ""
				sectId2 = ""

			    sectAttrGbcd	= rsget("sectAttrGbcd")
			    sectId1			= rsget("sectId1")
			    sectId2			= rsget("sectId2")

				If sectId1 <> "" Then
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<sectAttrGbcd>"&sectAttrGbcd&"</sectAttrGbcd>"				'매장속성구분코드 | 01 일반매장
					buf = buf & "				<sectId>"&sectId1&"</sectId>"								'매장ID | 상품의 정상 노출을 위해서는 1개 이상의 활성화 매장 등록이 필요함. 상품에 활성화 매장이 등록되어 있지 않은 경우 해당 데이터셋을 통해 추가 등록 가능함
					buf = buf & "			</row>"
				End If

				If sectId2 <> "" Then
					buf = buf & "			<row>"
					buf = buf & "				<rowType>"&rowType&"</rowType>"
					buf = buf & "				<sectAttrGbcd>"&sectAttrGbcd&"</sectAttrGbcd>"				'매장속성구분코드 | 01 일반매장
					buf = buf & "				<sectId>"&sectId2&"</sectId>"								'매장ID | 상품의 정상 노출을 위해서는 1개 이상의 활성화 매장 등록이 필요함. 상품에 활성화 매장이 등록되어 있지 않은 경우 해당 데이터셋을 통해 추가 등록 가능함
					buf = buf & "			</row>"
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		buf = buf & "		</rows>"
		buf = buf & "	</Dataset>"
		getHmallSectIdToReg = buf
	End Function


	Function fngetOptionEditParam(iitemid)
		Dim sqlStr, regedOptArr, i, buf, optionArr, j
		Dim optionLimitNo, optsellYn, boolchk, optTypeName, isSingleOption, slashReplace
		boolchk = False
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemid, outmallOptCode, replace(outmallOptName, '&amp;', '&') as outmallOptName, outmallSellyn, outmalllimitno "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_hmall_regedOption] "
		sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"' "
		rsget.Open sqlStr,dbget
		IF not rsget.EOF THEN
			regedOptArr = rsget.getRows()
		END IF
		rsget.close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemid, isusing, optsellyn, optlimityn, optlimitno, optlimitsold, replace(optionTypeName, char(9), '') as optionTypeName, replace(optionname, char(9), '') as optionname, Len(optionTypeName) as typeLength "
		sqlStr = sqlStr & " FROM db_item.dbo.tbl_item_option "
		sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"' "
		rsget.Open sqlStr,dbget
		IF not rsget.EOF THEN
			optionArr = rsget.getRows()
		END IF
		rsget.close

		If IsArray(regedOptArr) Then
			If Ubound(regedOptArr,2) = 0 AND regedOptArr(2, 0) = "단일옵션" Then			'등록된 것이 단품일때
				If FLimitYn = "Y" Then
					optionLimitNo = getLimitEa()
				Else
					If FHmallPrice >= 300000 Then
						optionLimitNo = 30
					Else
						optionLimitNo = 999
					End If
				End If

				If optionLimitNo < 1 Then
					optionLimitNo = 1
				End If

				buf = ""
				buf = buf & "{"
				buf = buf & "  ""itemid"": """&iitemid&""","
				buf = buf & "  ""options"": ["
				buf = buf & "    {"
				buf = buf & "      ""uitmcd"": """& regedOptArr(1, 0) &""","
				buf = buf & "      ""maxSellPossQty"": "&optionLimitNo&","
				buf = buf & "      ""sellGbcd"": """&Chkiif(IsSoldOutLimit5Sell = "True", "11", "00")&""""
				buf = buf & "    }"
				buf = buf & "  ]"
				buf = buf & "}"
			Else
				If IsArray(optionArr) Then
	'				If optionArr(6, 0) = Split(regedOptArr(2, 0), "/")(0) Then
					If optionArr(6, 0) = LEFT(Trim(regedOptArr(2, 0)), Trim(optionArr(8, 0))) Then
						isSingleOption = "Y"
					End If
				End If

				buf = ""
				buf = buf & "{"
				buf = buf & "  ""itemid"": """&iitemid&""","
				buf = buf & "  ""options"": ["
				For i = 0 To Ubound(regedOptArr, 2)
					buf = buf & "    {"
					buf = buf & "      ""uitmcd"": """& regedOptArr(1, i) &""","
					If IsArray(optionArr) Then
						For j = 0 To Ubound(optionArr, 2)
							If isSingleOption = "Y" Then
								slashReplace = replace(Trim(regedOptArr(2, i)), Trim(optionArr(6, 0)) & "/", "")
								slashReplace = replace(slashReplace, "∼", "~")
								slashReplace = replace(slashReplace, "∼", "&")

								If Trim(slashReplace) = Trim(optionArr(7, j)) Then
									If FLimitYn = "Y" Then
										optionLimitNo = getOptionLimitEa(optionArr(4, j), optionArr(5, j))
									Else
										optionLimitNo = 999
									End If

									If (optionArr(1, j) <> "Y") OR (optionArr(2, j) <> "Y") Then
										optsellYn = "11"
									Else
										optsellYn = "00"
									End If

									If optionLimitNo < 1 Then
										optsellYn = "11"
										optionLimitNo = 1			'판매 안 함이라도 재고가 1은 되야 오류가 안 남 ㅡㅡ;;
									End If

									boolchk = true
									Exit For
								End If
							Else
								slashReplace = replace(regedOptArr(2, i), "/", ",")
								slashReplace = replace(slashReplace, "∼", "~")

								If Trim(slashReplace) = Trim(replace(optionArr(7, j), "/", ",")) Then
									If FLimitYn = "Y" Then
										optionLimitNo = getOptionLimitEa(optionArr(4, j), optionArr(5, j))
									Else
										If FHmallPrice >= 300000 Then
											optionLimitNo = 30
										Else
											optionLimitNo = 999
										End If
									End If

									If (optionArr(1, j) <> "Y") OR (optionArr(2, j) <> "Y")  Then
										optsellYn = "11"
									Else
										optsellYn = "00"
									End If

									If optionLimitNo < 1 Then
										optsellYn = "11"
										optionLimitNo = 1			'판매 안 함이라도 재고가 1은 되야 오류가 안 남 ㅡㅡ;;
									End If

									boolchk = true
									Exit For
								End If
							End If
						Next
					End If

					If boolchk = false Then
						optionLimitNo = 1
						optsellYn = "11"
					End If
					buf = buf & "      ""maxSellPossQty"": "&optionLimitNo&","
					buf = buf & "      ""sellGbcd"": """&optsellYn&""""
					buf = buf & "    }"&Chkiif(i = Ubound(regedOptArr, 2), "", ",")  &"  "
				Next
				buf = buf & "  ]"
				buf = buf & "}"
			End If
		End If
		fngetOptionEditParam = buf
	End Function

	Public Function getHmallItemConfirmParameter
		Dim strRst, tt
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>텐바이텐</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "	<Dataset id=""dsCond"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"
		strRst = strRst & "				<itemCsfDCd />"
		strRst = strRst & "				<venCd>20</venCd>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "</Root>"
		getHmallItemConfirmParameter = strRst
	End Function

	Public Function gethmallItemRegParameter
		Dim strRst, childItemYn
		'################################ 안전인증 항목 최초 호출 ###############################
		Dim CallSafe, CSafeyn, CSafeGbCd, gbnflag
		CallSafe = getSafetyParam()
		CSafeyn = Split(CallSafe, "|_|")(0)
		CSafeGbCd = Split(CallSafe, "|_|")(1)
		gbnflag = Split(CallSafe, "|_|")(2)
		If CSafeGbCd = "N" Then CSafeGbCd = "" End If

		If FitemLCsfCd = "R6" and CSafeGbCd <> "" Then
			childItemYn = "Y"
		ElseIf gbnflag = "child" Then
			childItemYn = "Y"
		Else
			childItemYn = "N"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"		''실제는 UTF-8로 하기@!@@@@@@@@@@@@@@
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>텐바이텐</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItemCsfCd"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#상품대분류코드
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#상품중분류코드
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#상품소분류코드
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#상품세분류코드
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#상품분류구분코드 | 40 현대Hmall
		strRst = strRst & "				<qaTrgtYn>N</qaTrgtYn>"										'#QA대상여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<safeCertTrgtYn>N</safeCertTrgtYn>"							'#안전인증대상여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<coreMngYn>N</coreMngYn>"									'#핵심관리여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<itstDlbrYn>N</itstDlbrYn>"									'#상품기술서심의여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<frdlvSellLimtYn>Y</frdlvSellLimtYn>"						'#해외배송판매제한여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItem"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<slitmCd></slitmCd>"										'#판매상품코드
		strRst = strRst & "				<slitmNm><![CDATA["&getItemNameFormat()&"]]></slitmNm>"		'#판매상품명
		strRst = strRst & "				<sitmCd></sitmCd>"											'상품코드
		strRst = strRst & "				<bsitmCd></bsitmCd>"										'기준상품코드
		strRst = strRst & "				<bsitmNm></bsitmNm>"										'기준상품명
'		strRst = strRst & "				<baseCmpsItemNm></baseCmpsItemNm>"							'기본구성상품명
'		strRst = strRst & "				<addCmpsItemNm></addCmpsItemNm>"							'추가구성상품명
'		strRst = strRst & "				<engItemNm></engItemNm>"									'영문상품명
'		strRst = strRst & "				<itemUrl></itemUrl>"										'상품URL
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#상품대분류코드
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#상품중분류코드
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#상품소분류코드
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#상품세분류코드
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#상품분류구분코드 | 40 현대Hmall
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'협력사상품코드
		strRst = strRst & "				<afcrItemCd></afcrItemCd>"									'제휴사상품코드
		strRst = strRst & "				<frgnDrctBuyYn>N</frgnDrctBuyYn>"							'해외직구여부 | N/Y
		strRst = strRst & "				<mkcoGbcd>30</mkcoGbcd>"									'#제조사구분코드 | 10 제조사, 20 수입원, 30 판매원
		strRst = strRst & "				<mkcoCd>2347</mkcoCd>"										'#제조사코드
		strRst = strRst & "				<mkcoNm>텐바이텐</mkcoNm>"									'제조사명
		strRst = strRst & "				<octyCnryGbcd>"&FoctyCnryGbcd&"</octyCnryGbcd>"				'#원산지국가구분코드
		strRst = strRst & "				<octyCnryNm>"&FoctyCnryNm&"</octyCnryNm>"					'#원산지국가명
		strRst = strRst & "				<prmrOrgCnryGbcd></prmrOrgCnryGbcd>"						'주요원료국가구분코드
		strRst = strRst & "				<prmrOrgCnryNm></prmrOrgCnryNm>"							'주요원료국가명
		strRst = strRst & "				<prmrOrgNm></prmrOrgNm>"									'주요원료명	| 주원료국가구분코드/명 없이 등록불가
		strRst = strRst & "				<itemMemo></itemMemo>"										'상품메모
		strRst = strRst & "				<asGdCntn></asGdCntn>"										'AS안내내용	| 배송정보의 AS처리기준 
		strRst = strRst & "				<itemWgt></itemWgt>"										'상품중량 | frdlvYn 속성을 Y(해외배송)로 등록하는 경우 필수 무게 30kg 이상은 해외배송 가능 상품으로 설정"
		strRst = strRst & "				<itemWdthLen></itemWdthLen>"								'상품가로길이 | frdlvYn 속성을 Y(해외배송)로 등록하는 경우 필수 해외배송일 경우 소수점 한자리까지 등록"
		strRst = strRst & "				<itemHghLen></itemHghLen>"									'상품높이길이 | frdlvYn 속성을 Y(해외배송)로 등록하는 경우 필수 해외배송일 경우 소수점 한자리까지 등록"
		strRst = strRst & "				<itemHghtLen></itemHghtLen>"								'상품세로길이 | frdlvYn 속성을 Y(해외배송)로 등록하는 경우 필수 해외배송일 경우 소수점 한자리까지 등록"
		strRst = strRst & "				<itemGbcd>00</itemGbcd>"									'#상품구분코드 | 00 일반상품 ,01 폐쇄몰상품 ,02 공동구매상품 ,03 견적대표상품 ,04 무형상품 ,05 PPL상품 ,07 별도결제상품 ,09 구성품 ,10 TREND-H상품 ,11 클릭H상품"
		strRst = strRst & "				<itemSellGbcd>00</itemSellGbcd>"							'#상품판매구분코드 | 00 진행, 11 일시중단, 19 영구중단
		strRst = strRst & "				<adltItemYn>"&CHKIIF(IsAdultItem= "Y", "Y", "N")&"</adltItemYn>"	'#성인용품여부 | Y or N
		strRst = strRst & "				<itemRegTcndAgrYn>Y</itemRegTcndAgrYn>"						'#상품등록약관동의여부 | Y or N -> 반드시 Y
		strRst = strRst & "				<jwlSvrtEnclYn>N</jwlSvrtEnclYn>"							'#보석감정서동봉여부 | Y or N (보석감정서동봉시 Y)
		strRst = strRst & "				<giftItemYn>N</giftItemYn>"									'#사은품상품여부 | 'N'값으로 설정.
		strRst = strRst & "				<tcommUseYn>N</tcommUseYn>"									'#T커머스사용여부 | Y or N
		strRst = strRst & "				<stckGdYn>N</stckGdYn>"										'#재고안내여부 | Y or N
		strRst = strRst & "				<hmallRsvSellYn>N</hmallRsvSellYn>"							'#HMALL예약판매여부 | Y or N
		strRst = strRst & "				<frgnBuyPrxyYn>N</frgnBuyPrxyYn>"							'#해외구매대행여부 | Y or N
		strRst = strRst & "				<basktUseNdmtYn>N</basktUseNdmtYn>"							'#장바구니사용불가여부 | Y or N
		strRst = strRst & "				<prsnMsgPossYn>N</prsnMsgPossYn>"							'#선물메시지가능여부 | Y or N
		strRst = strRst & "				<prsnPackPossYn>N</prsnPackPossYn>"							'#선물포장가능여부 | Y or N
		strRst = strRst & "				<addBuyOptUseYn>Y</addBuyOptUseYn>"							'#추가구매옵션사용여부 | Y or N
		strRst = strRst & "				<oshpVenAdrSeq>4</oshpVenAdrSeq>"							'#출고협력사주소순번 | Y or N
		strRst = strRst & "				<rtpExchVenAdrSeq>4</rtpExchVenAdrSeq>"						'#반품교환협력사주소순번 | Y or N
		strRst = strRst & "				<emgyExchVenAdrSeq></emgyExchVenAdrSeq>"					'긴급교환협력사주소순번
		strRst = strRst & "				<itntDispYn>Y</itntDispYn>"									'#인터넷전시여부
		strRst = strRst & "				<itemQnaExpsYn>Y</itemQnaExpsYn>"							'#상품QNA노출여부
		strRst = strRst & "				<webExpsPrmoNm><![CDATA[]]></webExpsPrmoNm>"				'웹노출프로모션명
'		strRst = strRst & "				<prmo2TxtCntn></prmo2TxtCntn>"								'프로모션2 문구내용 | 프로모션2 문구내용 입력시 프로모션노출 시작일자/일시 및 프로모션노출 종료일자/일시(prmoExpsStrtDtm, prmoExpsStrtTime, prmoExpsEndDtm, prmoExpsEndTime) 필수입력"
'		strRst = strRst & "				<prmoExpsStrtDtm></prmoExpsStrtDtm>"						'프로모션노출 시작일자 | ex) 20191118 입력 prmo2TxtCntn 입력 시 필수(Y)"
'		strRst = strRst & "				<prmoExpsStrtTime></prmoExpsStrtTime>"						'프로모션노출 시작일시 | 0~23시 시간단위까지 입력 (15시부터일 경우 프로모션 시작일 경우 15 입력) prmo2TxtCntn 입력 시 필수(Y)"
'		strRst = strRst & "				<prmoExpsEndDtm></prmoExpsEndDtm>"							'프로모션노출 종료일자 | ex) 20191118 입력 prmo2TxtCntn 입력 시 필수(Y)"
'		strRst = strRst & "				<prmoExpsEndTime></prmoExpsEndTime>"						'프로모션노출 종료일시 | 0~24시 시간단위까지 입력 (0~23시일경우 23:00:00 로 세팅, 24시일 경우 23:59:59로 자동 세팅됨) prmo2TxtCntn 입력 시 필수(Y)"
		strRst = strRst & "				<prmoTxtDcCopnYn>N</prmoTxtDcCopnYn>"						'#프로모션문구할인쿠폰여부 | Y or N
		strRst = strRst & "				<prmoTxtSpdcYn>N</prmoTxtSpdcYn>"							'#프로모션문구깜짝할인여부 | Y or N
		strRst = strRst & "				<prmoTxtSvmtYn>N</prmoTxtSvmtYn>"							'#프로모션문구적립금여부 | Y or N
		strRst = strRst & "				<prmoTxtFamtFxrtGbcd>2</prmoTxtFamtFxrtGbcd>"				'프로모션문구정액정률구분코드 | "1 정률, 2 정액"
		strRst = strRst & "				<prmoTxtSvmtPrdcYn>N</prmoTxtSvmtPrdcYn>"					'#프로모션문구적립금선할인여부 | default 'N'값으로 설정.
		strRst = strRst & "				<prmoTxtWintYn>N</prmoTxtWintYn>"							'#프로모션문구무이자여부 | default 'N'값으로 설정.
		strRst = strRst & "				<prmoTxtSpymDcYn>N</prmoTxtSpymDcYn>"						'#프로모션문구일시불할인여부 | default 'N'값으로 설정.
		strRst = strRst & "				<frdlvYn>N</frdlvYn>"										'#해외배송여부 | default 'N'값으로 설정.
		strRst = strRst & "				<packOpenRtpNdmtYn>Y</packOpenRtpNdmtYn>"					'#포장오픈반품불가여부 | Y or N
		strRst = strRst & "				<ostkYn>N</ostkYn>"											'#품절여부 | default 'N'값으로 설정.
		strRst = strRst & "				<dlvHopeDtDsntYn>N</dlvHopeDtDsntYn>"						'#배송희망일자지정여부 | default 'N'값으로 설정.
		strRst = strRst & "				<itstHtmlYn>Y</itstHtmlYn>"									'#상품기술서HTML여부 | Y or N
		strRst = strRst & "				<itstPhotoExpsYn>Y</itstPhotoExpsYn>"						'#상품기술서사진노출여부 | default 'N'값으로 설정.
		strRst = strRst & "				<dlvItemFormGbcd></dlvItemFormGbcd>"						'배송상품형태구분코드	String(2)	"00 일반, 10 행거, 20 냉동
		strRst = strRst & "				<qckDlvPossYn>N</qckDlvPossYn>"								'#퀵배송가능여부 | default 'N'값으로 설정.
		strRst = strRst & "				<dwtdYn>N</dwtdYn>"											'#직회수여부 | default 'N'값으로 설정.
		strRst = strRst & "				<lrpyYn>Y</lrpyYn>"											'#후환불여부 | "Y or N (해외배송인 경우 'Y'로 입력되어야 함) 해외배송 상품은 후환불 상품"
		strRst = strRst & "				<sameItemMxpkPossQty></sameItemMxpkPossQty>"				'동일상품합포장가능수량	Number	
		strRst = strRst & "				<mxpkYn>N</mxpkYn>"											'#합포장여부 | default 'N'값으로 설정.
		strRst = strRst & "				<packMagnGbcd>20</packMagnGbcd>"							'#포장주체구분코드	String(2)	"10 당사 20 협력사" 
		strRst = strRst & "				<dlvcGbcd>00</dlvcGbcd>"									'#배송비구분코드	String(2)	'00' 일반상품 으로 입력
		strRst = strRst & "				<dlvcPayGbcd>10</dlvcPayGbcd>"								'#배송비지불구분코드	String(2)	"00 무료, 10 선결제, 20 착불, 30 설치상품 (입력불가)"
		strRst = strRst & "				<arpayDlvGdCntn></arpayDlvGdCntn>"							'착불배송안내내용	String(400)	
		strRst = strRst & "				<cvstWtdwPossYn>N</cvstWtdwPossYn>"							'#편의점회수여부 | default 'N'값으로 설정.
		strRst = strRst & "				<prpyDlvCost></prpyDlvCost>"								'선급배송비용 | 00(무료)시  미입력(null) 10(선결제)시  값이 존재하면 상품별배송비값이 null 이면 묶음배송비(소액장바구니) 20(착불)시  미입력(null) 30(설치상품)시  미입력(null) 백화점협력사&백화점배송일 경우 상품별배송비 등록 불가
		strRst = strRst & "				<irgnAreaAddDlvCost></irgnAreaAddDlvCost>"					'도서지역추가배송비용
		strRst = strRst & "				<mngWhNo>990</mngWhNo>"										'#관리창고번호
		strRst = strRst & "				<sbctDlvcoCd>12</sbctDlvcoCd>"								'#도급배송사코드
		strRst = strRst & "				<dlvMagnGbcd>20</dlvMagnGbcd>"								'#배송주체구분코드 | 10 홈쇼핑,20 협력사,30 택배사"
		strRst = strRst & "				<dlvcChmgGbcd>20</dlvcChmgGbcd>"							'#배송비부담주체구분코드 | 10 홈쇼핑,20 협력사,30 택배사NL NULL"
		strRst = strRst & "				<dlvFormGbcd>40</dlvFormGbcd>"								'#배송형태구분코드 | 00 센터배송,10 백화점배송,20 현대홈직택배,30 협력사직택배,40 협력사직송,50 백화점명절배송"
		strRst = strRst & "				<rtpWdmtGbcd>2</rtpWdmtGbcd>"								'반품회수방법구분코드 | 1 고객직접반송,2 협력사회수"
		strRst = strRst & "				<rtpDlvCost>6000</rtpDlvCost>"								'반품배송비용
		strRst = strRst & "				<exchWdmtGbcd>2</exchWdmtGbcd>"								'교환회수방법구분코드 | 1 고객직접반송,2 협력사회수"
		strRst = strRst & "				<exchDlvCost>6000</exchDlvCost>"							'교환배송비용
		strRst = strRst & "				<custDlvcWdmtGbcd>4</custDlvcWdmtGbcd>"						'고객배송비회수방법구분코드 | 1 배송박스동봉,2 무통장입금,3 택배기사지불,4 반품 접수시 선결제"
		strRst = strRst & "				<stlmWayScopGbcd>10</stlmWayScopGbcd>"						'결제수단범위구분코드 | 10 전 결제수단가능,20 현금/법인카드만사용가능,30 상품권제외 모두 결제 가능"
		strRst = strRst & "				<pntStlmNdmtYn>N</pntStlmNdmtYn>"							'#포인트결제불가여부
		strRst = strRst & "				<ostkRishpSmsYn>N</ostkRishpSmsYn>"							'#품절재입고SMS여부 | default 'N'값으로 설정.
		strRst = strRst & "				<oshpSmsExcldYn>N</oshpSmsExcldYn>"							'#출고SMS제외여부 | Y or N
		strRst = strRst & "				<hmallItemSrchExcldYn>N</hmallItemSrchExcldYn>"				'HMALL상품검색제외여부 | Y or N
		strRst = strRst & "				<itemTypeGbcd>01</itemTypeGbcd>"							'#상품유형구분코드 | 00 백화점상품, 01 협력사상품, 02 상품권, 03 위탁상품권, 04 도서/공연(백화점), 05 도서/공연(협력사) 도서나 공연 상품 등록시 백화점 협력사일 경우 04, 일반협력사일 경우 05 등록
'		strRst = strRst & "				<intgItemGbcd></intgItemGbcd>"								'무형상품구분코드 | 01 렌탈, 02 보험, 03 여행, 04 휴대폰, 05 장기할부, 06 모바일 상품권, 07 초고가 상품, 08 보험(Hmall)  상품구분코드(itemGbcd)값이 무형상품(04)일 경우에만 필수 입력
'		strRst = strRst & "				<intgItemStlmYn>N</intgItemStlmYn>"							'#무형상품여부
		strRst = strRst & "				<prchMdaGbcd>40</prchMdaGbcd>"								'#매입매체구분코드 | 40 Hmall  20 홈쇼핑
		strRst = strRst & "				<frstRegMdaGbcd>02</frstRegMdaGbcd>"						'#최초등록매체구분코드 | 02 인터넷
		strRst = strRst & "				<vatRate><![CDATA[10]]></vatRate>"							'#부가세비율
		strRst = strRst & "				<itemTaxnYn>"&Chkiif(Fvatinclude="Y", "Y", "N")&"</itemTaxnYn>"	'#상품과세여부 | Y or N
		strRst = strRst & "				<venCd>002569</venCd>"										'#협력사코드
		strRst = strRst & "				<ven2Cd></ven2Cd>"											'2차협력사코드
		strRst = strRst & "				<prchMthdGbcd>33</prchMthdGbcd>"							'#매입방법구분코드 | 11 직영, 22 특정, 33 수수료
		strRst = strRst & "				<itemTaxnGbcd>"&Chkiif(Fvatinclude="Y", "001", "000")&"</itemTaxnGbcd>"		'#상품과세구분코드 | 000 면세, 001 과세, 002 영세
		strRst = strRst & "				<ringItemYn>N</ringItemYn>"									'#반지상품여부 | default 'N'값으로 설정.
		strRst = strRst & "				<brndGbcd>40</brndGbcd>"									'#브랜드구분코드 | 40 Hmall브랜드
		strRst = strRst & "				<brndCd>205390</brndCd>"									'#브랜드코드 | 등록된 브랜드 전체 조회(selectBrndList) 없을 경우 해당 MD에게 신규 등록 요청
		strRst = strRst & "				<itntBrndNm>텐바이텐</itntBrndNm>"							'#인터넷브랜드명
		strRst = strRst & "				<dptsPchCd></dptsPchCd>"									'백화점펀칭코드 | 백화점협력사만 입력
		strRst = strRst & "				<rsptMdCd>8048</rsptMdCd>"									'#담당자MD코드
		strRst = strRst & getAttrInfo()
		strRst = strRst & "				<ordMakeYn>"&FordMakeYn&"</ordMakeYn>"						'#주문제작여부 | Y or N (주문제작시 Y)
'		strRst = strRst & "				<baseSectId></baseSectId>"									'#기본매장ID | 판매될 기본전시장
'		strRst = strRst & "				<frdlvFormGbcd></frdlvFormGbcd>"							'해외배송형태구분코드  | 1 박스, 2 봉투 (해외배송인 경우 필수 값)-> 해외 배송에서 봉투 선택시 봉투 최대 사이즈는 가로 52cm x 세로 44cm를 초과할 수 없음)"
'		strRst = strRst & "				<hscd></hscd>"												'HS코드 | 해외배송인 경우 필수 HS코드 값 유효성 체크(Harmonized System)"
'		strRst = strRst & "				<frgnOrdPiupSrvYn></frgnOrdPiupSrvYn>"						'해외주문픽업서비스여부 | default 'N'값으로 설정.		
'		strRst = strRst & "				<frdlvNchgYn></frdlvNchgYn>"								'해외배송무료여부 | default 'N'값으로 설정.
		strRst = strRst & "				<childItemYn>"&childItemYn&"</childItemYn>"					'#어린이상품여부 | "Y 13세미만 어린이상품 N 13세이상 상품"
'		strRst = strRst & "				<hdmalItnlYn></hdmalItnlYn>"								'더현대몰 연동 여부 | 백화점 상품의 경우 연동여부 Y로 선택 가능
'		strRst = strRst & "				<chkExceptSafeCert></chkExceptSafeCert>"					'안전인증등록예외처리 | 안전인증업체 예외 처리 대상 협력사일 경우 안전인증대상 대중소세로 상품 등록시 Y 셋팅하면 안전인증 필수입력하지 않아도 됨(db저장값 아님)당사 HELP 구현프로세스 참조
		strRst = strRst & "				<inslItemYn>N</inslItemYn>"									'설치상품여부 | Y 설치상품, N 설치상품아님
'		strRst = strRst & "				<meatHisYn></meatHisYn>"									'육류이력표시여부 | Y 표시, N 미표시
		strRst = strRst & "				<safeMngTrgtYn>"&Chkiif(gbnflag="elec", "Y", "N")&"</safeMngTrgtYn>"					'전안법대상여부 | Y 전안법 대상 (전안법 대상 상품일 경우 safeCertTypeGbcd 값 필수), N 전안법 비대상
		strRst = strRst & "				<chemSafeTrgtYn>"&Chkiif(gbnflag="life", "Y", "N")&"</chemSafeTrgtYn>"							'생활화학제품 대상여부 | Y : 생활화학제품 대상, N : 생활화학제품 비대상
		strRst = strRst & "				<parlImprYn>N</parlImprYn>"									'병행수입여부 | Y 병행수입상품, N 병행수입상품 아님
		strRst = strRst & "				<itemYetaGbcd>00</itemYetaGbcd>"							'#상품연말정산구분코드 | 00 일반, 01 도서, 02 공연, itemTypeGbcd를 04 혹은 05로 선택할 경우 도서일 경우 01, 공연상품일 경우 02 선택해야 함 일반상품일 경우 00으로 처리 필요
'		strRst = strRst & "				<dawnDlvYn></dawnDlvYn>"									'새벽배송여부 | Y or N
'		strRst = strRst & "				<stpicPossYn></stpicPossYn>"								'스토어픽가능여부 | Y or N
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"							'당일픽업가능여부 | Y or N
		strRst = strRst & "				<areaDlvCostAddYn>Y</areaDlvCostAddYn>"						'지역배송비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddDlvCost>3000</jejuAddDlvCost>"						'제주도추가배송비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddDlvCost>3000</irgnAddDlvCost>"						'도서추가배송비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<areaRtpCostAddYn>Y</areaRtpCostAddYn>"						'지역반품비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddRtpCost>3000</jejuAddRtpCost>"						'제주도추가반품비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddRtpCost>3000</irgnAddRtpCost>"						'도서추가반품비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<areaExchCostAddYn>Y</areaExchCostAddYn>"					'지역교환비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddExchCost>3000</jejuAddExchCost>"					'제주도추가교환비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddExchCost>3000</irgnAddExchCost>"					'도서추가교환비용 | 천원이상 만원이하 가능
'		strRst = strRst & "				<harmItemYn></harmItemYn>"									'위해상품여부 | Y or N
'		strRst = strRst & "				<brodEqmtMngTrgtYn></brodEqmtMngTrgtYn>"					'방송장비기자재대상여부 | Y or N
		strRst = strRst & "				<bbprcCopnDcYn>Y</bbprcCopnDcYn>"							'혜택모음가표시 쿠폰할인 | Y or N
		strRst = strRst & "				<bbprcSpymDcYn>Y</bbprcSpymDcYn>"							'혜택모음가표시 일시불할인 | Y or N
		strRst = strRst & "				<bbprcSvmtPrdcYn>Y</bbprcSvmtPrdcYn>"						'혜택모음가표시 H.Point선할인 | Y or N
		strRst = strRst & "				<bbprcSpdcYn>Y</bbprcSpdcYn>"								'혜택모음가표시 깜짝할인 | Y or N
		strRst = strRst & "				<prcExpsBitVal1>1</prcExpsBitVal1>"							'가격비교노출가 쿠폰할인 | 0: 해당없음, 1: 해당 ※ 가격비교노출가 관련 항목들은 셋트값으로 전부 다 입력하던가 아니면 아예 모두가 입력을 안하던가 해야 함. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal2>2</prcExpsBitVal2>"							'가격비교노출가 일시불할인 | 0: 해당없음, 2: 해당 ※ 가격비교노출가 관련 항목들은 셋트값으로 전부 다 입력하던가 아니면 아예 모두가 입력을 안하던가 해야 함. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal4>4</prcExpsBitVal4>"							'가격비교노출가 H.Point선할인 | 0: 해당없음, 4: 해당 ※ 가격비교노출가 관련 항목들은 셋트값으로 전부 다 입력하던가 아니면 아예 모두가 입력을 안하던가 해야 함. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
		strRst = strRst & "				<prcExpsBitVal8>8</prcExpsBitVal8>"							'가격비교노출가 깜짝할인 | 0: 해당없음, 8: 해당 ※ 가격비교노출가 관련 항목들은 셋트값으로 전부 다 입력하던가 아니면 아예 모두가 입력을 안하던가 해야 함. (prcExpsBitVal1,prcExpsBitVal2,prcExpsBitVal4,prcExpsBitVal8)
'		strRst = strRst & "				<frgnDrctDlvYn></frgnDrctDlvYn>"							'해외직배송(통관부호불필요) | Y or N
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallSectIdToReg("I")
		If CSafeyn = "Y" Then
			strRst = strRst & getHmallItemSafeInfoToReg(CSafeGbCd, "I")
		End If
'		strRst = strRst & "	<Dataset id=""dsSlitmBcdDtl"">"											'바코드내역
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<bcdBrndGbcd></bcdBrndGbcd>"								'바코드브랜드구분코드 | 10 일반브랜드, 20 백화점브랜드
'		strRst = strRst & "				<shrtBcdVal></shrtBcdVal>"									'단축바코드값
'		strRst = strRst & "				<totBcdVal></totBcdVal>"									'전체바코드값
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsHmallRsvItemDtl"">"										'Hmall예약판매정보
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<sellStrtDt></sellStrtDt>"									'#판매시작일자
'		strRst = strRst & "				<sellEndDt></sellEndDt>"									'#판매종료일자
'		strRst = strRst & "				<dlvStrtDt></dlvStrtDt>"									'#배송시작일자
'		strRst = strRst & "				<dlvEndDt></dlvEndDt>"										'#배송종료일자
'		strRst = strRst & "				<dlvAdmGdCntn></dlvAdmGdCntn>"								'#배송관리자안내내용
'		strRst = strRst & "				<custGdCntn></custGdCntn>"									'#고객안내내용
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsSlitmAsVenHis"">"										'AS협력사정보
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<asVenNm></asVenNm>"										'#AS협력사명
'		strRst = strRst & "				<asgnrNm></asgnrNm>"										'#담당자명
'		strRst = strRst & "				<rgno></rgno>"												'#사업자등록번호
'		strRst = strRst & "				<postNo></postNo>"											'#우편번호
'		strRst = strRst & "				<venBaseAdr></venBaseAdr>"									'#협력사기본주소
'		strRst = strRst & "				<venPtcAdr></venPtcAdr>"									'#협력사상세주소
'		strRst = strRst & "				<tela></tela>"												'#전화지역번호
'		strRst = strRst & "				<tels></tels>"												'#전화국번호
'		strRst = strRst & "				<teli></teli>"												'#전화개별번호
'		strRst = strRst & "				<extsTel></extsTel>"										'#내선전화번호
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsItemDlvNdmtDtl"">"										'배송불가지역
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<apocGbcd></apocGbcd>"										'#전국구분코드 | 10 서울,수도권, 11 서울신배송불가지역, 12 서울,수도권 명절 배송불가지역, 20 지방(백화점명절배송불가) ※일반업체선택불가!, 21 지방불가, 22 도서/산간지역 불가, 30 제주불가
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsSlitmPrcAthzHis"">"										'#가격정보
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<prcAplyStrtDtm>"&Replace(Date(), "-", "")&"</prcAplyStrtDtm>"	'#가격적용시작일시
		strRst = strRst & "				<prcAthzGbcd>00</prcAthzGbcd>"								'#가격결재구분코드 | 00 MD승인대기
		strRst = strRst & "				<sellPrc>"&MustPrice()&"</sellPrc>"							'#판매가격 | 실판매가 , VAT 별도 포함하지 않습니다. 제시된 가격 그대로 판매됩니다.
		strRst = strRst & "				<mrgnRate>"&FMrgnRate&"</mrgnRate>"							'#마진비율 | 마진율 입력,  '%'는 입력하지 말아주세요.
		strRst = strRst & "				<dptsOpCd></dptsOpCd>"										'#백화점OP코드 | OP코드
		strRst = strRst & "				<dptsVenOpCd></dptsVenOpCd>"								'백화점협력사OP코드 | 백화점협력사인경우 필수로 입력
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'협력사상품코드 | 협력사상품코드 존재시 필수 입력
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsHpItemDtl"">"											'핸드폰상품정보
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<trfsCntn></trfsCntn>"										'#요금제내용
'		strRst = strRst & "				<stplMths></stplMths>"										'#약정개월수
'		strRst = strRst & "				<ccrgAmt></ccrgAmt>"										'#위약금금액
'		strRst = strRst & "				<teRealChrgAmt</teRealChrgAmt>"								'#단말기실부담금액
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getOptSellUitmDtl()
		strRst = strRst & getOptTypeMst("I")
		strRst = strRst & getOptAttrMst("I")
		strRst = strRst & getOptCombDtl("I")

'		strRst = strRst & "	<Dataset id=""dsAsctSlitmDtl"">"										'관련상품정보
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<asctItemGbcd></asctItemGbcd>"								'#관련상품구분코드 | 10 트렌드H
'		strRst = strRst & "				<asctSlitmCd></asctSlitmCd>"								'#관련판매상품코드
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallItemInfoCdToReg("I")

		strRst = strRst & "	<Dataset id=""dsHtmlItstMst"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<htmlItstGbcd>00</htmlItstGbcd>"							'HTML상품기술서구분코드 | 00 일반, 01 식품, 02 영문설명서
		strRst = strRst & "				<htmlItstCntn><![CDATA["&getHmallContParamToReg()&"]]></htmlItstCntn>"	'상품기술서구분코드 | HTML상품기술서내용
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsMdaSlitmDtl"">"											'매체판매내역
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"							'판매매체분류코드 | 02 : Hmall, 04 : 모바일 두개 모두 체크 필요
		strRst = strRst & "			</row>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>INSERT</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>04</sellMdaCsfCd>"							'판매매체분류코드 | 02 : Hmall, 04 : 모바일 두개 모두 체크 필요
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsCsmtMkcoSlitmDtl"">"									'화장품제조사판매상품내역
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"							'화장품제조사순번 | 화장품 상품인경우 필수로 입력하여야 합니다.
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

'		strRst = strRst & "	<Dataset id=""dsItemIntlAddSetupDtl"">"									'더현대몰 연동정보내역
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>INSERT</rowType>"
'		strRst = strRst & "				<storpiupExclItemYn></storpiupExclItemYn>"			'스토어픽전용상품(택배불가) | dsItem.hdmalIntlYn(더현대몰 연동여부)가 Y일 경우 필수입력 입력값 : Y/N (스토어픽가능 이 Y일 경우에만 Y 가능)"
'		strRst = strRst & "				<storpiupPossYn></storpiupPossYn>"					'스토어픽가능 | dsItem.hdmalIntlYn(더현대몰 연동여부)가 Y일 경우 필수입력 입력값 : Y/N"
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"					'당일픽업가능 | dsItem.hdmalIntlYn(더현대몰 연동여부)가 Y일 경우 필수입력 입력값 : Y/N"
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & "</Root>"
		gethmallItemRegParameter = strRst
	End Function

	Public Function gethmallItemEditParameter
		Dim strRst, childItemYn
		'################################ 안전인증 항목 최초 호출 ###############################
		Dim CallSafe, CSafeyn, CSafeGbCd, gbnflag
		CallSafe = getSafetyParam()
		CSafeyn = Split(CallSafe, "|_|")(0)
		CSafeGbCd = Split(CallSafe, "|_|")(1)
		gbnflag = Split(CallSafe, "|_|")(2)
		If CSafeGbCd = "N" Then CSafeGbCd = "" End If

		If FitemLCsfCd = "R6" and CSafeGbCd <> "" Then
			childItemYn = "Y"
		ElseIf gbnflag = "child" Then
			childItemYn = "Y"
		Else
			childItemYn = "N"
		End If

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""dsSession"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>텐바이텐</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItemCsfCd"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#상품대분류코드
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#상품중분류코드
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#상품소분류코드
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#상품세분류코드
		strRst = strRst & "				<itemCsfGbcd>40</itemCsfGbcd>"								'#상품분류구분코드 | 40 현대Hmall
		strRst = strRst & "				<qaTrgtYn>N</qaTrgtYn>"										'#QA대상여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<safeCertTrgtYn>N</safeCertTrgtYn>"							'#안전인증대상여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<coreMngYn>N</coreMngYn>"									'#핵심관리여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<itstDlbrYn>N</itstDlbrYn>"									'#상품기술서심의여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "				<frdlvSellLimtYn>Y</frdlvSellLimtYn>"						'#해외배송판매제한여부 | 상품분류조회 후 해당 데이터입력
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsItem"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"						'#판매상품코드
		strRst = strRst & "				<slitmNm><![CDATA["&getItemNameFormat()&"]]></slitmNm>"		'#판매상품명
		strRst = strRst & "				<venCd>002569</venCd>"										'#협력사코드
'		strRst = strRst & "				<ven2Cd></ven2Cd>"											'2차협력사코드
'		strRst = strRst & "				<addCmpsItemNm></addCmpsItemNm>"							'추가구성상품명
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"							'협력사상품코드
		If FmaySoldOut = "Y" OR IsMayLimitSoldout = "Y" Then
			strRst = strRst & "			<itemSellGbcd>11</itemSellGbcd>"							'#판매상태구분 |00 진행, 11 일시중단, 19 영구중단..매뉴얼에는 없는 데 이거 없으면 수정안됨
		Else
			strRst = strRst & "			<itemSellGbcd>00</itemSellGbcd>"							'#판매상태구분 |00 진행, 11 일시중단, 19 영구중단..매뉴얼에는 없는 데 이거 없으면 수정안됨
		End If
'		strRst = strRst & "				<giftEvntStrtDtm></giftEvntStrtDtm>"						'사은품이벤트시작일시
'		strRst = strRst & "				<giftEvntEndDtm>S4</giftEvntEndDtm>"						'사은품이벤트종료일시
'		strRst = strRst & "				<giftCntn>S4</giftCntn>"									'사은품내용
'		strRst = strRst & "				<giftImgNm>S4</giftImgNm>"									'사은품이미지명
'		strRst = strRst & "				<tcommUseYn>N</tcommUseYn>"									'T커머스사용여부
'		strRst = strRst & "				<webExpsPrmoNm></webExpsPrmoNm>"							'웹노출프로모션명
'		strRst = strRst & "				<prmo2TxtCntn></prmo2TxtCntn>"								'프로모션2 문구내용
'		strRst = strRst & "				<prmoExpsStrtDtm></prmoExpsStrtDtm>"						'프로모션노출 시작일자
'		strRst = strRst & "				<prmoExpsStrtTime></prmoExpsStrtTime>"						'프로모션노출 시작일시
'		strRst = strRst & "				<prmoExpsEndDtm></prmoExpsEndDtm>"							'프로모션노출 종료일자
'		strRst = strRst & "				<prmoExpsEndTime></prmoExpsEndTime>"						'프로모션노출 종료일시
		strRst = strRst & "				<itstHtmlYn>Y</itstHtmlYn>"									'상품기술서HTML여부
		strRst = strRst & "				<itstPhotoExpsYn>Y</itstPhotoExpsYn>"						'상품기술서사진노출여부
		strRst = strRst & getAttrInfo()
		strRst = strRst & "				<childItemYn>"&childItemYn&"</childItemYn>"					'#어린이상품여부 | Y : 13세미만 어린이상품, N : 13세이상 상품
'		strRst = strRst & "				<childItemYn>N</childItemYn>"								'#어린이상품여부 | Y : 13세미만 어린이상품, N : 13세이상 상품
'		strRst = strRst & "				<childItemYn>"&maybeChildYn&"</childItemYn>"				'#어린이상품여부 | Y : 13세미만 어린이상품, N : 13세이상 상품
		strRst = strRst & "				<safeCertTypeGbcd>"&CSafeGbCd&"</safeCertTypeGbcd>"			'인증유형구분코드 | 01 안전인증, 02 안전확인, 03 공급자 적합성 확인, 04 안전기준준수대상
		strRst = strRst & "				<itemLCsfCd>"&FitemLCsfCd&"</itemLCsfCd>"					'#상품대분류코드
		strRst = strRst & "				<itemMCsfCd>"&FitemMCsfCd&"</itemMCsfCd>"					'#상품중분류코드
		strRst = strRst & "				<itemSCsfCd>"&FitemSCsfCd&"</itemSCsfCd>"					'#상품소분류코드
		strRst = strRst & "				<itemDCsfCd>"&FitemCsfGbcd&"</itemDCsfCd>"					'#상품세분류코드
'		strRst = strRst & "				<frgnDrctBuyYn></frgnDrctBuyYn>"							'해외직구여부
		strRst = strRst & "				<safeMngTrgtYn>"&Chkiif(gbnflag="elec", "Y", "N")&"</safeMngTrgtYn>"					'전안법대상여부 | Y 전안법 대상 (전안법 대상 상품일 경우 safeCertTypeGbcd 값 필수), N 전안법 비대상
		strRst = strRst & "				<chemSafeTrgtYn>"&Chkiif(gbnflag="life", "Y", "N")&"</chemSafeTrgtYn>"							'생활화학제품 대상여부 | Y : 생활화학제품 대상, N : 생활화학제품 비대상
		strRst = strRst & "				<parlImprYn>N</parlImprYn>"									'병행수입여부 | Y 병행수입상품, N 병행수입상품 아님
'		strRst = strRst & "				<stpicPossYn></stpicPossYn>"								'스토어픽가능여부 | Y or N
'		strRst = strRst & "				<thdyPiupPossYn></thdyPiupPossYn>"							'당일픽업가능여부 | Y or N
'		strRst = strRst & "				<brodEqmtMngTrgtYn></brodEqmtMngTrgtYn>"					'방송장비기자재대상여부 | Y or N
		strRst = strRst & "				<dlvFormGbcd>40</dlvFormGbcd>"								'#배송형태구분코드 | 00 센터배송 ,10 백화점배송 ,20 현대홈직택배 ,30 협력사직택배 ,40 협력사직송 ,50 백화점명절배송
		strRst = strRst & "				<areaDlvCostAddYn>Y</areaDlvCostAddYn>"						'지역배송비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddDlvCost>3000</jejuAddDlvCost>"						'제주도추가배송비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddDlvCost>3000</irgnAddDlvCost>"						'도서추가배송비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<areaRtpCostAddYn>Y</areaRtpCostAddYn>"						'지역반품비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddRtpCost>3000</jejuAddRtpCost>"						'제주도추가반품비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddRtpCost>3000</irgnAddRtpCost>"						'도서추가반품비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<areaExchCostAddYn>Y</areaExchCostAddYn>"					'지역교환비용추가여부 | Y or N
		strRst = strRst & "				<jejuAddExchCost>3000</jejuAddExchCost>"					'제주도추가교환비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<irgnAddExchCost>3000</irgnAddExchCost>"					'도서추가교환비용 | 천원이상 만원이하 가능
		strRst = strRst & "				<bbprcCopnDcYn>Y</bbprcCopnDcYn>"							'혜택모음가표시 쿠폰할인 | Y or N
		strRst = strRst & "				<bbprcSpymDcYn>Y</bbprcSpymDcYn>"							'혜택모음가표시 일시불할인 | Y or N
		strRst = strRst & "				<bbprcSvmtPrdcYn>Y</bbprcSvmtPrdcYn>"						'혜택모음가표시 H.Point선할인 | Y or N
		strRst = strRst & "				<bbprcSpdcYn>Y</bbprcSpdcYn>"								'혜택모음가표시 깜짝할인 | Y or N
		strRst = strRst & "				<prcExpsBitVal1>1</prcExpsBitVal1>"							'가격비교노출가 쿠폰할인 | 0: 해당없음, 1: 해당
		strRst = strRst & "				<prcExpsBitVal2>2</prcExpsBitVal2>"							'가격비교노출가 일시불할인 | 0: 해당없음, 2: 해당
		strRst = strRst & "				<prcExpsBitVal4>4</prcExpsBitVal4>"							'가격비교노출가 H.Point선할인 | 0: 해당없음, 4: 해당
		strRst = strRst & "				<prcExpsBitVal8>8</prcExpsBitVal8>"							'가격비교노출가 깜짝할인 | 0: 해당없음, 8: 해당
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		If CSafeyn = "Y" Then
			strRst = strRst & getHmallItemSafeInfoToReg(CSafeGbCd, "U")
		End If
'		strRst = strRst & "	<Dataset id=""dsSlitmBcdDtl"">"
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>UPDATE</rowType>"
'		strRst = strRst & "				<bcdBrndGbcd></bcdBrndGbcd>"								'바코드브랜드구분코드 | 10 일반브랜드, 20 백화점브랜드
'		strRst = strRst & "				<shrtBcdVal></shrtBcdVal>"									'단축바코드값
'		strRst = strRst & "				<totBcdVal></totBcdVal>"									'전체바코드값
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getOptTypeMst("U")
		strRst = strRst & getOptAttrMst("U")
		strRst = strRst & getOptCombDtl("U")

'		strRst = strRst & "	<Dataset id=""dsAsctSlitmDtl"">"
'		strRst = strRst & "		<rows>"
'		strRst = strRst & "			<row>"
'		strRst = strRst & "				<rowType>UPDATE</rowType>"
'		strRst = strRst & "				<asctItemGbcd></asctItemGbcd>"								'관련상품구분코드 | 10 트렌드H
'		strRst = strRst & "				<asctSlitmCd></asctSlitmCd>"								'관련판매상품코드
'		strRst = strRst & "			</row>"
'		strRst = strRst & "		</rows>"
'		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallItemInfoCdToReg("U")

		strRst = strRst & "	<Dataset id=""dsHtmlItstMst"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<htmlItstGbcd>00</htmlItstGbcd>"							'HTML상품기술서구분코드 | 00 일반, 01 식품, 02 영문설명서
		strRst = strRst & "				<htmlItstCntn><![CDATA["&getHmallContParamToReg()&"]]></htmlItstCntn>"	'상품기술서구분코드 | HTML상품기술서내용
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & "	<Dataset id=""dsMdaSlitmDtl"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>02</sellMdaCsfCd>"								'판매매체분류코드
		strRst = strRst & "			</row>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<sellMdaCsfCd>04</sellMdaCsfCd>"								'판매매체분류코드
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"

		strRst = strRst & getHmallSectIdToReg("U")
		strRst = strRst & "</Root>"
		gethmallItemEditParameter = strRst
	End Function

	Public Function getHmallPriceParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"
		strRst = strRst & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "	<Dataset id=""sessionVO"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<userId>hs002569</userId>"
		strRst = strRst & "				<userNm>텐바이텐</userNm>"
		strRst = strRst & "				<userGbcd>20</userGbcd>"
		strRst = strRst & "				<userIp>192.168.1.72</userIp>"
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "	<Dataset id=""dsItemPrcHistTran"">"
		strRst = strRst & "		<rows>"
		strRst = strRst & "			<row>"
		strRst = strRst & "				<rowType>UPDATE</rowType>"
		strRst = strRst & "				<slitmCd>"&FHmallGoodNo&"</slitmCd>"							'#상품코드
		strRst = strRst & "				<prcAplyStrtDtm>"&Replace(Date(), "-", "")&"</prcAplyStrtDtm>"	'#가격적용일 | 20131204
		strRst = strRst & "				<prcAplyStrtTime></prcAplyStrtTime>"							'가격적용시작시간 | 가격적용 시-분 (오후4시 30분은 1630으로 작성)
		strRst = strRst & "				<prcDcEndDtm></prcDcEndDtm>"									'가격종료일	| 20131208
		strRst = strRst & "				<prcDcEndTime></prcDcEndTime>"									'세일종료시간 | 세일종료 시-분 (오후4시 30분은 1630으로 작성)
		strRst = strRst & "				<prcAthzGbcd>00</prcAthzGbcd>"									'#결재요청구분코드 | 00 : 요청, 41: 요청취소
		strRst = strRst & "				<sellPrc>"&MustPrice()&"</sellPrc>"								'#판매가격
		strRst = strRst & "				<mrgnRate>"&FMrgnRate&"</mrgnRate>"								'#마진율
		strRst = strRst & "				<dptsVenOpCd></dptsVenOpCd>"									'백화점협력사OP코드	String(2)	
		strRst = strRst & "				<venItemCd>"&FItemid&"</venItemCd>"								'협력사상품코드	String(20)	
		strRst = strRst & "				<prmoCopyYn>Y</prmoCopyYn>"										'프로모션복사여부 | 기존 프로모션정보(무이자, 쿠폰, 일시불할인, 적립금)를  신규 가격수정시 그대로 복사해서 사용할 경우 Y
		strRst = strRst & "			</row>"
		strRst = strRst & "		</rows>"
		strRst = strRst & "	</Dataset>"
		strRst = strRst & "</Root>"
		getHmallPriceParameter = strRst
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CHmall
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
	Public FRectMatchCate
	Public FRectMatchShipping
	Public FRectGosiEqual
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectExpensive10x10
	Public FRectdiffPrc
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

	Public Sub getHmallNotRegOnlyOneItem
		strSql = ""
		strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_Hmall_Reg_Get] " & FRectItemID
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
				FOneItem.FAdultType			= db2html(rsget("adultType"))
				FOneItem.Fvatinclude		= rsget("vatinclude")
				FOneItem.FordMakeYn			= rsget("ordMakeYn")
		End If
		rsget.Close
	End Sub

	Public Sub getHmallNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			addSql = addSql & " and i.itemid not in ("
			addSql = addSql & "		SELECT itemid FROM ("
			addSql = addSql & "			SELECT itemid"
			addSql = addSql & " 		,count(*) as optCNT"
			addSql = addSql & " 		,sum(CASE WHEN optAddPrice>0 then 1 ELSE 0 END) as optAddCNT"
			addSql = addSql & " 		,sum(CASE WHEN (optsellyn='N') or (optlimityn='Y' and (optlimitno-optlimitsold<1)) then 1 ELSE 0 END) as optNotSellCnt"
			addSql = addSql & " 		FROM db_item.dbo.tbl_item_option"
			addSql = addSql & " 		WHERE itemid in (" & FRectItemID & ")"
			addSql = addSql & " 		and isusing='Y'"
			addSql = addSql & " 		GROUP BY itemid"
			addSql = addSql & " 	) T"
            addSql = addSql & " where optAddCNT>0"
            addSql = addSql & " or (optCnt-optNotSellCnt<1)"
            addSql = addSql & " )"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum "
		strSql = strSql & "	, isNULL(R.hmallStatCD,-9) as hmallStatCD, isNull(R.hmallPrice, 0) as hmallPrice "
		strSql = strSql & "	, UC.socname_kor "
		strSql = strSql & "	,(SELECT [db_etcmall].[dbo].[getHmallMargin] (" & FRectItemID & ")) as mrgnRate"
		strSql = strSql & "	,S.octyCnryGbcd, S.octyCnryNm"
		strSql = strSql & "	,LEFT(am.CateKey, 2) as itemLCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 4) as itemMCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 6) as itemSCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 8) as itemCsfGbcd "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_hmall_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_hmall_sourceCodeName] (" & FRectItemID & ") as S on i.itemid = S.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_category as tm on am.CateKey = tm.CateKey "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_hmall_regItem as R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c as UC on i.makerid = UC.userid"
		strSql = strSql & " Where i.isusing='Y' "
		strSql = strSql & " and i.isExtUsing='Y' "
		strSql = strSql & " and UC.isExtUsing<>'N' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
'		strSql = strSql & " and ((i.deliveryType<>9) or ((i.deliveryType=9) and (i.sellcash>=10000)))"
		strSql = strSql & " and i.sellyn='Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "		'플라워/화물배송/해외직구
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv<50 and i.itemdiv<>'08' "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.cate_large<>'999' "
		strSql = strSql & " and i.sellcash>0 "
		strSql = strSql & " and ((i.LimitYn='N') or ((i.LimitYn='Y') and (i.LimitNo - i.LimitSold > "&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.

'		strSql = strSql & " and (i.sellcash<>0 and convert(int, ((i.sellcash-i.buycash)/i.sellcash)*100) >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((sellcash - buycash)/ sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((orgprice - orgsuplycash)/ orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN sailyn = 'N' AND (Round(((sellcash - buycash)/ sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun = '"&CMALLNAME&"') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun = '"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and isnull(R.hmallStatCD,0) < 3  "
		strSql = strSql & " and cm.mapCnt is Not Null "		'카테고리 매칭 상품만
		strSql = strSql & " and i.itemdiv not in ('06') "	'주문제작문구 상품 제외
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
		End If
		rsget.Close
	End Sub

	Public Sub getHmallEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & " ,m.hmallGoodNo, m.hmallSellyn, m.regImageName, isNull(m.hmallprice, 0) as hmallprice "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or ((i.sailyn <> 'Y') and (i.sellcash + round(i.orgprice * 0.5, 0) < m.hmallprice)) "	'할인이 아니고 정상가의 50%이상인 판매가가 hmall에 등록된 경우
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & " 	or i.itemdiv in ('06') "		''주문제작문구 상품 품절처리
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "
		strSql = strSql & "		or ((i.sellcash < 50000) AND (i.itemname like '%무료배송%' or i.itemname like '%무료 배송%')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & "	,(SELECT [db_etcmall].[dbo].[getHmallMargin] (" & FRectItemID & ")) as mrgnRate"
		strSql = strSql & "	,S.octyCnryGbcd, S.octyCnryNm"
		strSql = strSql & "	,LEFT(am.CateKey, 2) as itemLCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 4) as itemMCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 6) as itemSCsfCd "
		strSql = strSql & "	,LEFT(am.CateKey, 8) as itemCsfGbcd "
		strSql = strSql & " ,IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_hmall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_hmall_sourceCodeName] (" & FRectItemID & ") as S on i.itemid = S.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_Hmall_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.hmallGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CHmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Ficon2Image		= "http://webimage.10x10.co.kr/image/icon2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("icon2Image")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FHmallGoodNo		= rsget("hmallGoodNo")
				FOneItem.FHmallSellYn		= rsget("hmallSellYn")
				FOneItem.FMrgnRate			= rsget("mrgnRate")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName		= rsget("regImageName")
				FOneItem.FHmallprice		= rsget("hmallprice")

				FOneItem.FoctyCnryGbcd		= rsget("octyCnryGbcd")
				FOneItem.FoctyCnryNm		= rsget("octyCnryNm")
				FOneItem.FitemLCsfCd		= rsget("itemLCsfCd")
				FOneItem.FitemMCsfCd		= rsget("itemMCsfCd")
				FOneItem.FitemSCsfCd		= rsget("itemSCsfCd")
				FOneItem.FitemCsfGbcd		= rsget("itemCsfGbcd")
				FOneItem.Fitemsize			= db2html(rsget("itemsize"))
				FOneItem.Fitemsource		= db2html(rsget("itemsource"))
				FOneItem.Fordercomment		= db2html(rsget("ordercomment"))
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

'Hmall 상품코드 얻기
Function getHmallGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ISNULL(hmallgoodno, '') as hmallgoodno FROM db_etcmall.dbo.tbl_hmall_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getHmallGoodno = rsget("hmallgoodno")
	End If
	rsget.Close
End Function

'Hmall 상품코드 얻기
Function getHmallGoodno2(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 ISNULL(hmallgoodno, '') as hmallgoodno FROM db_etcmall.dbo.tbl_hmall_regitem WHERE itemid = '"&iitemid&"' and APIaddImg = 'Y' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getHmallGoodno2 = rsget("hmallgoodno")
	End If
	rsget.Close
End Function

'텐바이텐 기본 이미지 얻기
Function getTenBasicImage(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT basicimage " & VBCRLF
	strSql = strSql & " FROM db_item.dbo.tbl_item  " & VBCRLF
	strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getTenBasicImage = rsget("basicimage")
	End If
	rsget.Close
End Function
%>