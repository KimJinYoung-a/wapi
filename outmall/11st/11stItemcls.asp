<%
CONST CMAXMARGIN = 18
CONST CMALLNAME = "11st1010"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST APIURL = "http://api.11st.co.kr/rest"
CONST APISSLURL = "https://api.11st.co.kr/rest"
CONST APIkey = "a2319e071dbc304243ee60abd07e9664"
CONST CDEFALUT_STOCK = 99999

Class C11stItem
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
	Public FSt11StatCD
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
	Public FDepthCode
	Public Fcdmkey
	Public Fcddkey
	Public FSt11GoodNo
	Public FSt11price
	Public FSt11SellYn

	Public FSafeDiv
	Public FIsNeed
	Public FDepth1Code
	Public FAdultType
	Public FOrderMaxNum
	Public FOutmallstandardMargin

	'// 품절여부
	public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	end function

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
		Dim ownItemCnt
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
			If (GetTenTenMargin < outmallstandardMargin) Then
				tmpPrice = Forgprice
			Else
				tmpPrice = FSellCash
			End If
		End If
		MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	End Function

	'최대 구매 수량
	Public Function getLimit11stEa()
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
		getLimit11stEa = ret
	End Function

	'// 11st 판매여부 반환
	Public Function get11stSellYn()
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				get11stSellYn = "Y"
			Else
				get11stSellYn = "N"
			End If
		Else
			get11stSellYn = "N"
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

    Public Function GetSourcearea()
		If IsNULL(Fsourcearea) or (Fsourcearea="") then
			GetSourcearea = "."
		Else
			GetSourcearea = Fsourcearea
		End if
    End function

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
' o11st.FOneItem.FLimityn
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

	'// 상품등록: 상품설명 파라메터 생성(상품등록용)
	Public Function get11stContParamToReg()
		Dim strRst, strSQL,strtextVal
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_11st.jpg'></p><br>"

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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_11st.jpg"">")
		get11stContParamToReg = strRst

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
			strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_11st.jpg'></p><br>"
			strRst = strRst & Replace(Replace(strtextVal,"",""),"","")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_11st.jpg"">")
			get11stContParamToReg = strRst
		End If
		rsget.Close
	End Function

	'// 검색어
	Public Function getItemKeyword()
		Dim arrRst, arrRst2, q, Keyword1, strRst
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
			getItemKeyword = LeftB(arrRst(0), 20) &","&LeftB(arrRst(1), 20) &","& LeftB(arrRst(2), 20) &","& LeftB(arrRst(3), 20) &","& LeftB(arrRst(4), 20)
		Else
			For q = 0 to Ubound(arrRst)
				Keyword1 = Keyword1&LeftB(arrRst(q), 20) &","
			Next
			If Right(keyword1,1) = "," Then
				keyword1 = Left(keyword1,Len(keyword1)-1)
			End If
			getItemKeyword = keyword1
		End If
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

	Public Function get11stAddImageParam()
		Dim strRst, strSQL, i
		strRst = ""
		strRst = strRst & "	<prdImage01>"&FbasicImage&"</prdImage01>"					'#대표 이미지 URL | 이미지는 11번가 서버가 다운로드하여 300 x 300 사이즈로 리사이징 한뒤 11번가 이미지서버에 저장 됩니다. 이미지 확장자는 gif, jpg, jpeg, png 만 사용가능합니다. 이미지 url 호출시 "Content-Type" 이 정의가 되어있지 않으면 이미지 다운로드가 이루어 지지않습니다.
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "	<prdImage0"&i+1&">http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"</prdImage0"&i+1&">"					'추가 이미지 1 URL
				End If
				rsget.MoveNext
				If i>=3 Then Exit For
			Next
		End If
		rsget.Close
'		strRst = strRst & "	<prdImage05/>"					'목록이미지 | 검색 결과 페이지나 카테고리 리스트 페이지에서 노출되는 이미지입니다.
'		strRst = strRst & "	<prdImage09/>"					'카드뷰이미지 | 쇼킹딜/기획전용 카드뷰 형태 이미지입니다.
'		strRst = strRst & "	<prdImage01Src/>"				'이미지 바이트 코드 | 바이트 코드로 변환하여 보내셔야 합니다.
		get11stAddImageParam = strRst
	End Function

	Public Function get11stSafeParam()
		Dim strRst, certTypeCd, strSql, arrRows, notarrRows, nlp, newDiv, newCertNo
		If FSafeDiv = "2" Then

			strSql = ""
			strSql = strSql & " SELECT TOP 5 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				arrRows = rsget.getRows()
			Else
				notarrRows = "Y"
			End If
			rsget.Close

			If notarrRows = "" Then		'전안법 적용된 데이터라면 제대로 꼽기
				If FsafetyYn = "Y" Then
					For nLp =0 To UBound(arrRows,2)
				    	newDiv = ""
						Select Case arrRows(1,nLp)
							Case "10"		newDiv = "102"		'전기용품 > 안전인증
							Case "20"		newDiv = "104"		'전기용품 > 안전확인 신고
							Case "30"		newDiv = "127"		'전기용품 > 공급자 적합성 확인
							Case "40"		newDiv = "101"		'생활제품 > 안전인증
							Case "50"		newDiv = "103"		'생활제품 > 자율안전확인
							Case "60"		newDiv = "124"		'생활제품 > 안전품질표시
							Case "70"		newDiv = "128"		'어린이제품 > 안전인증
							Case "80"		newDiv = "129"		'어린이제품 > 안전확인
							Case "90"		newDiv = "130"		'어린이제품 > 공급자 적합성 확인
						End Select

						newCertNo = arrRows(0,nLp)
						If newCertNo = "x" Then
							newCertNo = ""
						End If

						strRst = strRst & "	<ProductCert>"
						strRst = strRst & "		<certTypeCd>"&newDiv&"</certTypeCd>"
						strRst = strRst & "		<certKey><![CDATA["&newCertNo&"]]></certKey>"				'인증번호
						strRst = strRst & "	</ProductCert>"
					Next
				Else
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>132</certTypeCd>"								'#132 : [전기용품/생활용품] 상품상세설명 참조
					strRst = strRst & "		<certKey/>"
					strRst = strRst & "	</ProductCert>"

				End If
			Else
				If FsafetyYn = "Y" AND FSafetyNum <> "" Then
					Select Case FsafetyDiv
						Case "10"	certTypeCd = "101"													'[공산품] 안전인증
						Case "20"	certTypeCd = "102"													'[전기용품] 안전인증
						Case "30"	certTypeCd = "124"													'[공산품] 안전/품질표시
						Case "40"	certTypeCd = "103"													'[공산품] 자율안전확인
						Case "50"	certTypeCd = "123"													'[공산품] 어린이보호포장
					End Select
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>"&certTypeCd&"</certTypeCd>"
					strRst = strRst & "		<certKey><![CDATA["&FSafetyNum&"]]></certKey>"				'인증번호
					strRst = strRst & "	</ProductCert>"
				Else
					strRst = strRst & "	<ProductCert>"
					strRst = strRst & "		<certTypeCd>132</certTypeCd>"							'#132 : [전기용품/생활용품] 상품상세설명 참조
					strRst = strRst & "		<certKey/>"
					strRst = strRst & "	</ProductCert>"
				End If
			End If
		Else
			'strRst = strRst & "	<ProductCert>"
			'strRst = strRst & "		<certTypeCd>131</certTypeCd>"								'#131 : 해당없음(대상이 아닌 경우)..개편된듯 131코드 사라짐
			'strRst = strRst & "		<certKey/>"
			'strRst = strRst & "	</ProductCert>"
		End If
		get11stSafeParam = strRst
	End Function

	Public Function get11stSafeNewParam()
		Dim strRst, certTypeCd, strSql, arrRows, notarrRows, nlp, newDiv, newCertNo, i, crtfGrpObjClfCd
		Dim crtfGrpTypCd, certNum, safetyDiv, certKey
		If FSafeDiv = "2" Then
			strSql = ""
			strSql = strSql & " SELECT TOP 1 certNum, safetyDiv " & vbcrlf
			strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg " & vbcrlf
			strSql = strSql & " WHERE itemid='"&FItemID&"' " & vbcrlf
			strSql = strSql & " ORDER BY regdate DESC "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				certNum = rsget("certNum")
				safetyDiv = rsget("safetyDiv")
				If certNum = "x" Then
					certNum = ""
				End If
			Else
				certNum = ""
				safetyDiv = ""
			End If
			rsget.Close

			strRst = ""
			For i = 1 to 4
				Select Case i
					Case "1"			crtfGrpTypCd = "01"				'01: 전기용품/생활용품 KC인증
					Case "2"			crtfGrpTypCd = "02"				'02: 어린이제품 KC인증
					Case "3"			crtfGrpTypCd = "03"				'03 : 방송통신기자재 KC인증
					Case "4"			crtfGrpTypCd = "04"				'04 : 생활화학 및 살생물제품
				End Select

				crtfGrpObjClfCd = ""
				'crtfGrpObjClfCd(01) : KC인증대상, crtfGrpObjClfCd(03 or 05) : 03 KC인증대상아님 / 05 생활화학 및 살생물제품 대상 아님
				Select Case crtfGrpTypCd
					Case "01"
						If (safetyDiv = "10") OR (safetyDiv = "20") OR (safetyDiv = "40") OR (safetyDiv = "50") Then
							crtfGrpObjClfCd = "01"
						End If
					Case "02"
						If (safetyDiv = "70") OR (safetyDiv = "80") Then
							crtfGrpObjClfCd = "01"
						End If
					Case "04"
						crtfGrpObjClfCd = "05"
					Case Else
						crtfGrpObjClfCd = "03"
				End Select

				'만약 생화학이 아니고 인증번호가 없으면 03 처리
				If crtfGrpTypCd <> "04" AND certNum = "" Then
					crtfGrpObjClfCd = "03"
				End If

				If crtfGrpObjClfCd = "" Then
					crtfGrpObjClfCd = "03"
				End If

				Select Case safetyDiv
					Case "10"		newDiv = "102"		'전기용품 > 안전인증
					Case "20"		newDiv = "104"		'전기용품 > 안전확인 신고
					Case "30"		newDiv = "127"		'전기용품 > 공급자 적합성 확인
					Case "40"		newDiv = "101"		'생활제품 > 안전인증
					Case "50"		newDiv = "103"		'생활제품 > 자율안전확인
					Case "60"		newDiv = "124"		'생활제품 > 안전품질표시
					Case "70"		newDiv = "128"		'어린이제품 > 안전인증
					Case "80"		newDiv = "129"		'어린이제품 > 안전확인
					Case "90"		newDiv = "130"		'어린이제품 > 공급자 적합성 확인
				End Select

				If crtfGrpTypCd = "01" AND ((safetyDiv = "10") OR (safetyDiv = "20") OR (safetyDiv = "40") OR (safetyDiv = "50")) Then
					certTypeCd = newDiv
					certKey = certNum
				ElseIf crtfGrpTypCd = "02" AND ((safetyDiv = "70") OR (safetyDiv = "80")) Then
					certTypeCd = newDiv
					certKey = certNum
				Else
					certTypeCd = ""
					certKey = ""
				End If

				strRst = strRst & "	<ProductCertGroup>"												'인증정보그룹
				strRst = strRst & "		<crtfGrpTypCd>"&crtfGrpTypCd&"</crtfGrpTypCd>"				'인증정보그룹번호 | 인증정보그룹번호가 존재하지 않는 식품 인증의 경우, 해당 값을 입력하지 않습니다. 전기용품/생활용품, 어린이제품, 방송통신기자재, 생활화학 및및 살생물제품에 대한 인증정보 입력이 필수인 카테고리일 경우 01, 02, 03, 04의 인증정보를 모두 입력해주세요. → 01 : 전기용품/생활용품 KC인증 → 02 : 어린이제품 KC인증 → 03 : 방송통신기자재 KC인증 → 04 : 생활화학 및 살생물제품
				strRst = strRst & "		<crtfGrpObjClfCd>"&crtfGrpObjClfCd&"</crtfGrpObjClfCd>"		'KC인증대상여부 | 인증정보그룹번호가 01, 02, 03, 04인 경우 인증대상여부 값을 필수 입력해야 합니다. (인증정보그룹번호 01 : 인증대상여부 01, 02, 03 택 1 사용 가능 / 인증정보그룹번호 02 : 인증대상여부 01, 03 택 1 사용 가능 / 인증정보그룹번호 03 : 인증대상여부 01, 03 택 1 사용 가능 / 인증정보그룹번호 04 : 인증대상여부 04, 05 택 1 사용 가능) → 01 : KC인증대상 → 02 : KC면제대상 → 03 : KC인증대상 아님 → 04 : 생활화학 및 살생물제품 대상 → 05 : 생활화학 및 살생물제품 대상 아님
	'			strRst = strRst & "		<crtfGrpExptTypCd></crtfGrpExptTypCd>"						'KC면제유형 | KC인증대상여부가 02인 경우 KC면제유형 값을 필수 입력해야 합니다. → 02 : 구매대행면제대상 → 03 : 병행수입면제대상
				strRst = strRst & "		<ProductCert>"												'인증정보 | 인증정보는 최대 100개 까지만 제공합니다.
				If certTypeCd <> "" Then
					strRst = strRst & "			<certTypeCd>"&certTypeCd&"</certTypeCd>"			'인증유형
				Else
					strRst = strRst & "			<certTypeCd></certTypeCd>"							'인증유형
				End If

				If certKey <> "" Then
					strRst = strRst & "			<certKey><![CDATA["&certKey&"]]></certKey>"			'인증번호
				Else
					strRst = strRst & "			<certKey></certKey>"								'인증번호
				End If
				strRst = strRst & "		</ProductCert>"
				strRst = strRst & "	</ProductCertGroup>"
			Next
		End If
		get11stSafeNewParam = strRst
	End Function

	Public Function get11stOptionParam()
		Dim strSql, strRst1, strRst2, strRst3, i, optNm, optDc, chkMultiOpt, optLimit, arrList1, arrList2, optCode, j, optionname, optaddprice, tmpoName
		Dim voptsellyn, visusing
		strRst1 = ""
		strRst2 = ""
		strRst3 = ""
		chkMultiOpt = false
		If FoptionCnt > 0 Then
			strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	        rsget.CursorLocation = adUseClient
			rsget.CursorType = adOpenStatic
			rsget.LockType = adLockOptimistic
		    rsget.Open strSql,dbget,1
		    if not rsget.Eof then
		    	chkMultiOpt = true
		        arrList1 = rsget.getRows()
		    end if
		    rsget.close
		End If
		strRst1 = strRst1 & "	<optSelectYn>Y</optSelectYn>"									'선택형 옵션 여부 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
		strRst1 = strRst1 & "	<txtColCnt>1</txtColCnt>"										'고정값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 옵션을 등록하실 경우 1 고정값을 주셔야 합니다.
'		If chkMultiOpt = True Then
'		strRst1 = strRst1 & "	<optionAllQty>"&getLimit11stEa&"</optionAllQty>"				'멀티옵션 일괄재고수량 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 재고 수량 설정이 api 에서는 불가합니다. 일괄설정만 가능.
'		strRst1 = strRst1 & "	<optionAllAddPrc>0</optionAllAddPrc>"							'멀티옵션 옵션가 0원 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 옵션가 설정이 api 에서는 불가합니다. 0원만 입력 가능
'		strRst1 = strRst1 & "	<optionAllAddWght/>"											'멀티옵션 일괄옵션추가무게 설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "상품상세 옵션값 노출 방식 선택"을 생략하실 경우 등록순 옵션이 노출됩니다. "멀티옵션" 방식이 아닌 "싱글옵션" 방식 일 경우는 Element는 생략해주셔야 합니다. 멀티옵션은 옵션별 옵션무게 설정이 api 에서는 불가합니다. 일괄설정만 가능.
'		End If
		strRst1 = strRst1 & "	<prdExposeClfCd>00</prdExposeClfCd>"							'상품상세 옵션값 노출 방식 선택 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 00 : 등록순, 01 : 옵션값 가나다순, 02 : 옵션값 가나다 역순, 03 : 옵션가격 낮은 순, 04 : 옵션가격 높은 순
		strRst1 = strRst1 & "	<optUpdateYn>Y</optUpdateYn>"
'		strRst1 = strRst1 & "	<optMixYn/>"													'전체옵션 조합여부 |"옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 정의한 전체 옵션값이 조합되어 멀티옵션으로 등록, N : 옵션 매핑Key에 존재하는(선택 된) 값으로만 멀티 옵션 등록
		If chkMultiOpt = True Then
			strSql = "select typeseq, optionTypeName, optionKindName, optaddPrice from db_item.[dbo].[tbl_item_option_Multiple] where itemid = " & FItemid & " ORDER BY Typeseq, kindSeq "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If not rsget.Eof Then
				arrList2 = rsget.getRows()
			End If
			rsget.close

			For i = 0 To Ubound(arrList1, 2)
				strRst2 = strRst2 & "	<ProductRootOption>"											'ProductRootOption |"옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
				strRst2 = strRst2 & "		<colTitle>"&arrList1(2,i)&"</colTitle>"						'옵션명 | 40Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다.
				For j = 0 to Ubound(arrList2, 2)
					If arrList1(1,i) = arrList2(0,j) then

						arrList2(2,j) = replace(arrList2(2,j), "&", "+")			'2017-06-05 김진영..혜리 대리 요청으로 &->+로 수정

						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<colOptPrice>0</colOptPrice>"					'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "			<colValue0>"&arrList2(2,j)&"</colValue0>"		'옵션값 | 50Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다. 한 상품안에서 옵션값은 중복이 될수 없습니다.
						strRst2 = strRst2 & "		</ProductOption>"
					end if
				Next
				strRst2 = strRst2 & "	</ProductRootOption>"
			Next

			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice, isUsing, optsellyn "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
'			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "		'2017-05-17 김진영 수정
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				strRst2 = strRst2 & "	<ProductOptionExt>"												'ProductOptionExt | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    tmpoName	= db2html(rsget("optionname"))
				    visUsing= rsget("isUsing")
				    voptsellyn= rsget("optsellyn")

					optionname = ""
					For i = 0 To Ubound(arrList1, 2)
						If Ubound(Split(tmpoName, ",")) > 0 Then				'2017-06-15 김진영 추가
							optionname = optionname & arrList1(2,i) &":"&Split(tmpoName, ",")(i) &"†"
						End If
					Next

					If Right(optionname,1) = "†" Then
						optionname = Left(optionname, Len(optionname) - 1)
					End If

				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
			    	If voptsellyn <> "Y" Then optLimit = 0	'2017-05-17 김진영 수정
			    	If visUsing <> "Y" Then optLimit = 0	'2017-05-17 김진영 수정

			'		If optLimit > 0 Then		'2017-05-17 김진영 수정
						strRst2 = strRst2 & "		<ProductOption>"
						strRst2 = strRst2 & "			<useYn>"&chkiif(optLimit>0,"Y","N")&"</useYn>"
						strRst2 = strRst2 & "			<colOptPrice>"&optaddprice&"</colOptPrice>"				'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "			<colCount>"&optLimit&"</colCount>"			'옵션재고수량 | 멀티옵션일 경우는 일괄설정이 되므로 입력하시면 안됩니다. 옵션상태(useYn)가 N일 때만 0 입력 가능합니다.
						strRst2 = strRst2 & "			<colSellerStockCd>"&optCode&"</colSellerStockCd>"		'셀러재고번호 | 셀러가 사용하는 재고번호
						strRst2 = strRst2 & "			<optionMappingKey>"&optionname&"</optionMappingKey>"	'옵션매핑Key | 멀티옵션의 조합된 옵션을 매핑하기 위한 Key(예: 옵션명1:옵션값1†옵션명2:옵션값2)
						strRst2 = strRst2 & "		</ProductOption>"
			'		End If
					rsget.MoveNext
				Loop
				strRst2 = strRst2 & "	</ProductOptionExt>"
			End If
			rsget.Close
		Else
			strSql = "Select itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where itemid=" & FItemid
			strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) Then
				If db2Html(rsget("optionTypeName"))<>"" Then
					optNm = Replace(db2Html(rsget("optionTypeName")),":","")
				Else
					optNm = "옵션"
				End If
				strRst1 = strRst1 & "	<colTitle>"&optNm&"</colTitle>"
				Do until rsget.EOF
					optCode		= FItemid&"_"&rsget("itemoption")
					optaddprice = rsget("optaddprice")
				    optLimit	= rsget("optLimit")
				    optionname	= db2html(rsget("optionname"))
				    optionname = replace(optionname, "&", "+")			'2017-06-05 김진영..혜리 대리 요청으로 &->+로 수정
				    optionname = replace(optionname, ",", "+")
				    optLimit = optLimit - 5
				    If (optLimit < 1) Then optLimit = 0
				    If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					If optLimit > 0 Then
						strRst2 = strRst2 & "	<ProductOption>"								'ProductOption | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
						strRst2 = strRst2 & "		<useYn>Y</useYn>"							'옵션상태 | 멀티옵션일 경우는 지원하지 않는 기능입니다. Y : 사용함, N : 품절
						strRst2 = strRst2 & "		<colOptPrice>"&optaddprice&"</colOptPrice>"	'옵션가 | 기본 판매가의 +100%/-50%까지 설정하실 수 있습니다. 옵션가격이 0원인 상품이 반드시 1개 이상 있어야 합니다.
						strRst2 = strRst2 & "		<colValue0>"&optionname&"</colValue0>"		'옵션값 | 50Byte 까지만 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다. 한 상품안에서 옵션값은 중복이 될수 없습니다
						strRst2 = strRst2 & "		<colCount>"&optLimit&"</colCount>"			'옵션재고수량 | 멀티옵션일 경우는 일괄설정이 되므로 입력하시면 안됩니다. 옵션상태(useYn)가 N일 때만 0 입력 가능합니다.
						strRst2 = strRst2 & "		<colSellerStockCd>"&optCode&"</colSellerStockCd>"	'셀러재고번호 | 셀러가 사용하는 재고번호
						strRst2 = strRst2 & "	</ProductOption>"
					End If
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If

		If FitemDiv = "06" Then
			strRst3 = strRst3 & "	<ProductCustOption>"										'옵션등록 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 구매자작성형 옵션의 등록 최대 5개까지 등록 가능
			strRst3 = strRst3 & "		<colOptName>텍스트를 입력하세요</colOptName>"				'구매자 작성형 옵션명 | 옵션명 최대 공백포함 한글10자/영문(숫자)20자)까지 입력가능하며 특수 문자[',",%,&,<,>,#,†]는 입력할 수 없습니다.
			strRst3 = strRst3 & "		<colOptUseYn>Y</colOptUseYn>"							'옵션 사용 여부 | Y : 사용함, N : 사용안함
			strRst3 = strRst3 & "	</ProductCustOption>"
		End If
		get11stOptionParam = strRst1 & strRst2 & strRst3
'rw get11stOptionParam
'response.end
	End Function

	'기본정보 등록 XML
	Public Function get11stItemRegParameter
		Dim strRst
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8"" standalone=""yes""?>"
		strRst = strRst & "<Product>"
'		strRst = strRst & "	<abrdBuyPlace/>"												'해외상품코드
'		strRst = strRst & "	<abrdSizetableDispYn/>"											'해외사이즈 조견표 노출여부
'		strRst = strRst & "	<selMnbdNckNm><![CDATA[텐바이텐]]></selMnbdNckNm>"				'닉네임 | 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요. 닉네임을 입력하지 않으면 대표 닉네임이 자동으로 등록됩니다. @ 텐바이텐 입력하면 오류발생
		strRst = strRst & "	<selMthdCd>01</selMthdCd>"										'#판매방식 | 01 : 고정가판매, 02 : 사용안함, 03 : 사용안함, 04 : 예약판매, 05 : 중고판매
		strRst = strRst & "	<dispCtgrNo>"&FDepthCode&"</dispCtgrNo>"						'#카테고리번호 | 최하위 카테고리만 입력가능합니다. 세카테고리를 입력하셔야 하며 세카테고리가 없는 경우 소카테고리를 입력하셔야 합니다. 카테고리번호 조회 서비스를 이용하여 실시간 조회가 가능합니다. 카테고리 수정은 세카테고리 (혹은 소카테고리) 까지만 가능합니다. 대카테고리, 중카테고리를 변경하고자 할 경우 상품을 새로 등록해주세요.
		strRst = strRst & "	<prdTypCd>01</prdTypCd>"										'#서비스 상품 코드 | 여행 카테고리 선택 시 타입을 구분하여 등록할 수 있습니다. 여행 제휴사 상품은 여행 상품등록처리 API를 이용해 주시기 바랍니다. 01 : 일반배송상품, 13 : 제휴사 여행상품
'		strRst = strRst & "	<hsCode/>"														'?#H.S Code | 대한민국 관세청에 신고되는 HSCode 입니다. 해외쇼핑 통합배송 상품과, 이태리직배송상품인 경우에만 필수 설정합니다. SO 혹은 PO의 상품등록 카테고리에서 기본 HSCode를 확인 하실 수 있습니다. 상품별 성격에 맞는 HSCode를 선택하셔야 하며, 잘못 등록하여 문제가 발생할 경우 셀러분께서 해결하셔야 합니다. 아래 첨부파일을 참조 부탁 드립니다.
		strRst = strRst & "	<prdNm><![CDATA["&getItemNameFormat&"]]></prdNm>"				'#상품명 | 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요. 글자수는 50Byte로 제한 될 예정입니다. 한글 25자, 영문/숫자 50자 이내로 입력을 권장합니다. 입력이 불가한 특수문자가 포함될 경우, 해당 문자는 상품명에서 자동 미노출처리 됩니다. [자세히보기]
'		strRst = strRst & "	<prdNmEng/>"													'영문 상품명
'		strRst = strRst & "	<engDispYn/>"													'11번가 영문 노출 | Y : 노출, N : 비노출
'		strRst = strRst & "	<advrtStmt/>"													'상품홍보문구 | 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요. 글자수는 40Byte 로 제한됩니다. 한글 20자, 영문/숫자 40자 이내로 입력해 주십시오.
		strRst = strRst & "	<brand><![CDATA[텐바이텐]]></brand>"								'브랜드 | 브랜드를 정확히 입력하면 해당 상품의 검색 노출이 더 많아집니다. 브랜드는 텍스트 형태로만 입력 브랜드 관련 서비스에 전시를 위해서는 브랜드명을 정확히 입력해 주셔야 합니다. 특히 스펠링에 유의해주세요.
		strRst = strRst & "	<rmaterialTypCd>05</rmaterialTypCd>"							'#원재료 유형 코드 | 01 : 농산물, 02 : 수산물, 03 : 가공품, 04 : 원산지 의무 표시대상 아님, 05 : 상품별 원산지는 상세설명 참조
		strRst = strRst & "	<orgnTypCd>03</orgnTypCd>"										'원산지 코드 | 01 : 국내. 국내원산지 코드를 같이 입력해야 합니다, 02 : 해외. 해외원산지 코드를 같이 입력해야 합니다, 03 : 기타. 원산지명을 입력해야합니다
'		strRst = strRst & "	<orgnTypDtlsCd/>"												'원산지 지역 코드 | 원산지 코드가 "국내", "해외"일 경우 원산지 지역 코드 값을 입력하셔야 합니다.
		strRst = strRst & "	<orgnNmVal><![CDATA["&GetSourcearea&"]]></orgnNmVal>"			'원산지명 | 원산지 코드가 "기타"일 경우 원산지명을 입력하셔야 합니다.
'		strRst = strRst & "	<ProductRmaterial>"												'원재료 정보 | 원재료 유형이 가공품(03)일 경우 원재료 정보를 입력하셔야 합니다. 상품은 최대 10개, 상품의 원재료 성분 정보는 최대 5개까지 등록 가능합니다.
'		strRst = strRst & "		<rmaterialNm/>"												'원재료 상품명
'		strRst = strRst & "		<ingredNm/>"												'원료명
'		strRst = strRst & "		<orgnCountry/>"												'원산지
'		strRst = strRst & "		<content/>"													'함량
'		strRst = strRst & "	</ProductRmaterial>"
		strRst = strRst & "	<beefTraceStat>03</beefTraceStat>"								'#축산물 이력번호 | 01 : 이력번호 표시대상 제품, 02 : 이력번호 표시대상 아님, 03 : 상세설명 참조
'		strRst = strRst & "	<beefTraceNo/>"													'이력번호 표시대상 제품 | 01 : 이력번호 표시대상제품 선택시 이력번호 표시대상 제품(xxxx)에 들어갈 내용을 입력합니다. 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요. 글자수는 20Byte 로 제한됩니다. 한글 10자, 영문/숫자 20자 이내로 입력해 주십시오.
		strRst = strRst & "	<sellerPrdCd>"&FItemid&"</sellerPrdCd>"							'판매자 상품코드 | 중복이 가능하며 본 코드값으로 11번가 상품 조회 등이 가능합니다. 필수값이 아니며 생략 가능합니다.
		strRst = strRst & "	<suplDtyfrPrdClfCd>"&CHKIIF(FVatInclude="N","02","01")&"</suplDtyfrPrdClfCd>"	'#부가세/면세상품코드 | 면세상품 선택시, 세무/법률적 책임은 판매자님께 있습니다. 01 : 과세상품, 02 : 면세상품, 03 : 영세상품
'		strRst = strRst & "	<forAbrdBuyClf></forAbrdBuyClf>"								'#해외구매대행상품 여부 | SellerOffice 가입 시에 글로벌셀러로 가입한 경우에만 사용할 수 있습니다. 일반 셀러인 경우 생략해주세요. 해외거주 글로벌셀러는 일반판매상품(01)로만 등록가능합니다. 문의사항이 있으시면 11번가 담당MD와 상담해 주세요.01 : 일반판매상품, 02 : 해외판매대행상품
	If FItemdiv = "06" OR FItemdiv = "16" Then
		strRst = strRst & "	<prdStatCd>10</prdStatCd>"										'#상품상태 | 주문제작상품으로 등록하시면 구매자의 취소/반품/교환이 불가능하여 클레임이 발생할 수 있으니 신중하게 등록해주시기 바랍니다. Open API로 주문제작상품 등록 시, 판매자가 위 내용에 대해 숙지한 후 동의한 것으로 간주됩니다. 01 : 새상품, 02 : 중고상품 (판매방식이 "중고판매"인 경우만 선택가능합니다.) 03 : 재고상품 04 : 리퍼상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 05 : 전시(진열)상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 07 : 희귀소장품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 08 : 반품상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 09 : 스크래치상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 10 : 주문제작상품
	Else
		strRst = strRst & "	<prdStatCd>01</prdStatCd>"										'#상품상태 | 주문제작상품으로 등록하시면 구매자의 취소/반품/교환이 불가능하여 클레임이 발생할 수 있으니 신중하게 등록해주시기 바랍니다. Open API로 주문제작상품 등록 시, 판매자가 위 내용에 대해 숙지한 후 동의한 것으로 간주됩니다. 01 : 새상품, 02 : 중고상품 (판매방식이 "중고판매"인 경우만 선택가능합니다.) 03 : 재고상품 04 : 리퍼상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 05 : 전시(진열)상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 07 : 희귀소장품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 08 : 반품상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 09 : 스크래치상품(판매방식이 "중고판매"인 경우만 선택가능합니다.) 10 : 주문제작상품
	End If
'		strRst = strRst & "	<useMon/>"														'사용개월수 | 판매방식이 중고판매인 경우 반드시 입력해 주셔야 합니다.
'		strRst = strRst & "	<paidSelPrc/>"													'구입당시 판매가 | 판매방식이 중고판매인 경우 반드시 입력해 주셔야 합니다.
'		strRst = strRst & "	<exteriorSpecialNote/>"											'외관/기능상 특이사항 | 판매방식이 중고판매인 경우 반드시 입력해 주셔야 합니다.
		strRst = strRst & "	<minorSelCnYn>"&Chkiif(IsAdultItem() = "Y", "N", "Y")&"</minorSelCnYn>"		'#미성년자 구매가능 | 미성년자 구매불가를 선택하시면, 미성년자 회원에게 상품이미지가 노출되지 않으며 '19금'으로 표시됩니다. 구매불가 상품을 구매가능으로 표시한 경우, 판매금지 처리 될 수 있습니다. Y : 가능, N : 불가능
		strRst = strRst & get11stAddImageParam()
		strRst = strRst & "	<htmlDetail><![CDATA["&get11stContParamToReg&"]]></htmlDetail>"	'상세설명 | iframe 사용은 가능하지만 권장하지 않습니다. html 을 입력하실 경우 <![CDATA[ ]]> 로 묶어 주세요. 외부로의 링크는 제한되며 자세한 사항은 상세설명 물음표를 참조해 주세요. html을 입력하는 경우 일부 스크립트 및 스타일 태그는 제한되니 SellerOffice 상품등록에서 상세설명 html 미리보기를 반드시 테스트해 주세요. html guide를 준수하여 입력하면, 구매 고객이 옵션 찾기가 편리해집니다. Guide를 이용하여 상세설명을 등록해 보세요.
'		strRst = strRst & get11stSafeParam()		''2022-07-04 개편 전
		strRst = strRst & get11stSafeNewParam()		''2022-07-04 개편 후
'		strRst = strRst & "	<ProductMedical>"												'의료기기 품목허가
'		strRst = strRst & "		<MedicalKey/>"												'의료기기 품목허가번호
'		strRst = strRst & "		<MedicalRetail/>"											'의료기기 판매업신고 기관 및 번호
'		strRst = strRst & "		<MedicalAd/>"												'의료기기사전광고심의번호
'		strRst = strRst & "	</ProductMedical>"
		strRst = strRst & "	<reviewDispYn>Y</reviewDispYn>"									'상품리뷰/후기 전시여부
		strRst = strRst & "	<reviewOptDispYn>Y</reviewOptDispYn>"							'상품리뷰/후기 옵션 노출여부
		strRst = strRst & "	<selTermUseYn>N</selTermUseYn>"									'#판매기간 | Y : 설정함. 판매기간 설정이 가능합니다., N : 설정안함(고정가판매의 경우만). 즉시 영구판매가 이루어 집니다.
'		strRst = strRst & "	<selPrdClfCd/>"													'판매기간코드/예약기간코드 | 0:100 : 판매기간 직접입력. 판매방식 - "고정가판매" 일 경우만 사용가능, 3:101 : 3일, 5:102 : 5일, 7:103 : 7일, 15:104 : 15일, 30:105 : 30일(1개월), 60:106 : 60일(2개월), 90:107 : 90일(3개월), 120:108 : 120일(4개월), 3:401 : 3일, 5:402 : 5일, 7:403 : 7일, 15:404 : 15일, 30:405 : 30일(1개월), 60:406 : 60일(2개월), 90:407 : 90일(3개월), 0:400 : 예약기간 직접입력
'		strRst = strRst & "	<aplBgnDy/>"													'판매 시작일/예약 시작일
'		strRst = strRst & "	<aplEndDy/>"													'판매 종료일/예약 종료일 | 판매기간/예약기간 직접 입력일 경우만 입력. 나머지는 자동 계산. 생략해주세요
		strRst = strRst & "	<setFpSelTermYn>N</setFpSelTermYn>"								'고정가 판매기간 설정 | Y : 설정함, N : 설정안함..예약판매일 때만 설정 가능
'		strRst = strRst & "	<selPrdClfFpCd/>"												'판매기간코드 | 고정가 판매기간 설정 - Y인 경우만 사용가능. 0:100 : 설정안함, 3:101 : 3일, 5:102 : 5일, 7:103 : 7일, 15:104 : 15일, 30:105 : 30일(1개월), 60:106 : 60일(2개월), 90:107 : 90일(3개월), 120:108 : 120일(4개월)
'		strRst = strRst & "	<wrhsPlnDy/>"													'입고예정일 | 입고예정일은 판매종료일과 같은 날, 혹은 그 이후로 설정해 주셔야 하며, 주문처리 시, 최대 15일에 한해서 1회 연장할 수 있습니다. 입력하신 입고예정일은 상품상세 페이지에 안내되며, 입고예정일 지연 시, 신용점수가 차감되오니, 유의해 주십시오.
'		strRst = strRst & "	<contractCd/>"													'약정코드 | 휴대폰 카테고리에 상품 등록시 필수로 설정하셔야 합니다. 01 : 일반 약정 단말기, 02 : 요금제 약정 단말기
'		strRst = strRst & "	<chargeCd/>"													'요금제 코드 | 요금제 약정 단말기인 경우 반드시 입력하셔야 합니다.
'		strRst = strRst & "	<periodCd/>"													'약정기간 코드 | 01 : 무약정, 02 : 12개월, 03 : 18개월, 04 : 24개월, 05 : 30개월, 06 : 36개월
'		strRst = strRst & "	<phonePrc/>"													'단말기 출고 가격 | ,없이 숫자만 입력하세요. 60,000(X) 60000(O)
		If FDepth1Code = "2967" Then		'도서카테고리
		strRst = strRst & "	<maktPrc>"&ForgPrice&"</maktPrc>"								'정가 | 카테고리가 도서인 경우 반드시 입력해 주셔야 합니다.(음반, DVD/블루레이제외) 도서 정가제 관련 규정을 준수하여 등록하셔야 합니다. 도서정가제 규정을 위반하거나 정가를 허위로 등록할 경우 법적 책임은 판매자에게 있으며, 판매중지 등의 불이익을 받을 수 있습니다. 도서 정가제: 18개월 미만의 신간도서의 경우 정가의 10% 이내 할인, 적립율은 판매가의 최대 10% 이내 적용 가능합니다.(마일리지, OK캐쉬백 등 합산 적립)
		End If
		strRst = strRst & "	<selPrc>"&MustPrice&"</selPrc>"									'#판매가 | 판매가는 10원 단위로, 최대 10억 원 미만으로 입력 가능합니다. 판매가 정보 수정 시, 최대 50% 인상/80% 인하까지 수정하실 수 있습니다. 서비스이용료는 카테고리/판매가에 따라 다르게 적용될 수 있습니다.
'		strRst = strRst & "	<cuponcheck/>"													'기본즉시할인 설정여부 | "기본즉시할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. S : 기존값 유지(상품수정) 는 상품수정시에만 입력가능합니다. 쿠폰에 대한 수정이 일어나지 않습니다. Y : 설정함, N : 설정안함, S : 기존값 유지(상품수정)
'		strRst = strRst & "	<dscAmtPercnt/>"												'할인수치 | "기본즉시할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 판매가에서(xxxx)에 들어갈 수치를 입력합니다.
'		strRst = strRst & "	<cupnDscMthdCd/>"												'할인단위 코드 | "기본즉시할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : 원, 02 : %
'		strRst = strRst & "	<cupnUseLmtDyYn/>"												'할인 적용기간 설정여부 | "기본즉시할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "할인 적용기간 설정"을 하실 경우에만 Element를 입력해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<cupnIssEndDy/>"												'할인적용기간 종료일 | "기본즉시할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "할인 적용기간 설정"을 하실 경우에만 Element를 입력해 주세요.
'		strRst = strRst & "	<ocbYN/>"														'OK캐쉬백 지급 설정여부 | "OK캐쉬백 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<ocbValue/>"													'적립수치 | "OK캐쉬백 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 판매가에서(xxxx) 에 들어갈 수치를 입력합니다.
'		strRst = strRst & "	<ocbWyCd/>"														'적립단위 코드 | "OK캐쉬백 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : %, 02 : 원
'		strRst = strRst & "	<mileageYN/>"													'마일리지 지급 설정여부 | "마일리지 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<mileageValue/>"												'적립수치 | "마일리지 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 판매가에서(xxxx) 에 들어갈 수치를 입력합니다.
'		strRst = strRst & "	<mileageWyCd/>"													'적립단위 코드 | "마일리지 지급"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : %, 02 : 원
'		strRst = strRst & "	<intFreeYN/>"													'무이자 할부 제공 설정여부 | "무이자 할부 제공"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<intfreeMonClfCd/>"												'개월수| "무이자 할부 제공"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 05 : 2개월, 01 : 3개월, 06 : 4개월, 07 : 5개월, 02 : 6개월, 08 : 7개월, 09 : 8개월, 03 : 9개월, 10 : 10개월, 11 : 11개월, 04 : 12개월
'		strRst = strRst & "	<pluYN/>"														'복수구매할인 설정 여부 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<pluDscCd/>"													'복수구매할인 설정 기준 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : 수량기준, 02 : 금액기준
'		strRst = strRst & "	<pluDscBasis/>"													'복수구매할인 기준 금액 및 수량 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
'		strRst = strRst & "	<pluDscAmtPercnt/>"												'복수구매할인 금액/율 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
'		strRst = strRst & "	<pluDscMthdCd/>"												'복수구매할인 구분코드 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : %, 02 : 원
'		strRst = strRst & "	<pluUseLmtDyYn/>"												'복수구매할인 적용기간 설정 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<pluIssStartDy/>"												'복수구매할인 적용기간 시작일 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "할인 적용기간 설정"을 하실 경우에만 Element를 입력해 주세요.
'		strRst = strRst & "	<pluIssEndDy/>"													'복수구매할인 적용기간 종료일 | "복수 구매할인"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. "할인 적용기간 설정"을 하실 경우에만 Element를 입력해 주세요.
'		strRst = strRst & "	<hopeShpYn/>"													'희망후원 지급 설정 여부 | "희망후원"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. Y : 설정함, N : 설정안함
'		strRst = strRst & "	<hopeShpPnt/>"													'적립수치 | "희망후원"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
'		strRst = strRst & "	<hopeShpWyCd/>"													'적립단위 코드 | "희망후원"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : %, 02 : 원
	If (FOptionCnt > 0) OR (FItemdiv = "06") Then
		strRst = strRst & get11stOptionParam()
	End If
'		strRst = strRst & "	<useOptCalc/>"													'계산형옵션 설정여부 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 등록 : 계산형 옵션 정보 입력 시 사용, 미입력 사용안함 수정 : Y : 사용, N : 삭제 계산형 옵션은 가구/수납가구/학생가구, 침구/커튼/카페트, 홈/인테리어/DIY 카테고리에서만 사용 가능합니다. 계산형 옵션은 조합형 옵션을 최소 1개 이상 함께 등록해야 사용 가능합니다. 계산형 옵션은 작성형 옵션과 동시에 사용할 수 없습니다. 독립형 옵션을 사용하면 계산형 옵션을 사용할 수 없습니다. 판매최소값, 판매최대값, 단가기준값, 판매단위-숫자는 숫자로 입력하세요. 초기 개발시에는 옵션등록과 SellerOffice 상품등록과 반드시 병행하시면서 개발해주셔야 합니다.
'		strRst = strRst & "	<optCalcTranType/>"												'계산형 옵션 타입설정 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. reg : 등록, upd : 수정
'		strRst = strRst & "	<optTypCd/>"													'계산형옵션구분값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
'		strRst = strRst & "	<optItem1Nm/>"													'첫번째 계산형 옵션명 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 최대 20byte, 초과 내용은 삭제
'		strRst = strRst & "	<optItem1MinValue/>"											'첫번째 계산형 옵션 판매최소값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 1~1,000,000
'		strRst = strRst & "	<optItem1MaxValue/>"											'첫번째 계산형 옵션 판매최대값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 1~1,000,000
'		strRst = strRst & "	<optItem2Nm/>"													'두번째 계산형 옵션명 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 최대 20byte, 초과 내용은 삭제
'		strRst = strRst & "	<optItem2MinValue/>"											'두번째 계산형 옵션 판매최소값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 1~1,000,000
'		strRst = strRst & "	<optItem2MaxValue/>"											'두번째 계산형 옵션 판매최대값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 1~1,000,000
'		strRst = strRst & "	<optUnitPrc/>"													'?단가기준값 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 0.001~1,000,000
'		strRst = strRst & "	<optUnitCd/>"													'?기준 단위코드 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 01 : mm, 02 : cm, 03 : m
'		strRst = strRst & "	<optSelUnit/>"													'?판매단위-숫자 | "옵션등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요. 입력 숫자 범위 1~1,000,000
'		strRst = strRst & "	<ProductComponent>"												'?추가구성상품 | "추가구성상품등록"을 설정 하지 않을시에는 Element를 모두 삭제해 주세요.
'		strRst = strRst & "		<addPrdGrpNm/>"
'		strRst = strRst & "		<compPrdNm/>"
'		strRst = strRst & "		<sellerAddPrdCd/>"
'		strRst = strRst & "		<addCompPrc/>"
'		strRst = strRst & "		<compPrdQty/>"
'		strRst = strRst & "		<compPrdVatCd/>"
'		strRst = strRst & "		<addUseYn/>"
'		strRst = strRst & "		<addPrdWght/>"
'		strRst = strRst & "	</ProductComponent>"
		strRst = strRst & "	<prdSelQty>"&getLimit11stEa&"</prdSelQty>"						'재고수량 | 재고 수량은 반드시 입력하셔야 하며 옵션이 있을 경우 입력값과 상관없이 옵션수량의 총합으로 자동계산 되어 반영됩니다. 재고는 0으로 입력할 수 없습니다. 상품 판매 중단을 원하시면 판매중지 처리하시기 바랍니다.
'		strRst = strRst & "	<selMinLimitTypCd/>"											'최소구매수량 설정코드 | "최소구매수량" 서비스를 이용하지 않으신다면 Element를 생략해 주세요. 자동 "제한 안한(00)"으로 설정됩니다. 구매 제한 해당 제한 기간은 한달(30일)입니다. | 00 : 제한 안함, 01 : 1회 제한
'		strRst = strRst & "	<selMinLimitQty/>"												'최소구매수량 개수 | "최소구매수량" 서비스를 이용하지 않으신다면 Element를 생략해 주세요. 자동 "제한 안한(00)"으로 설정됩니다.
		strRst = strRst & "	<selLimitTypCd>01</selLimitTypCd>"								'최대구매수량 설정코드 | "최대구매수량" 서비스를 이용하지 않으신다면 "설정코드"와 상관없이 Element를 생략해 주세요. 자동 "제한 안한(00)"으로 설정됩니다. 00 : 제한 안함, 01 : 1회 제한, 02 : 기간 제한
		strRst = strRst & "	<selLimitQty>"& FOrderMaxNum &"</selLimitQty>"					'최대구매수량 개수 | "최대구매수량" 서비스를 이용하지 않으신다면 "설정코드"와 상관없이 Element를 생략해 주세요. 자동 "제한 안한(00)"으로 설정됩니다.
'		strRst = strRst & "	<townSelLmtDy/>"												'최대구매수량 재구매기간 | "최대구매수량" 서비스를 이용하지 않으신다면 "설정코드"와 상관없이 Element를 생략해 주세요. 자동 "제한 안한(00)"으로 설정됩니다.
		strRst = strRst & "	<useGiftYn>N</useGiftYn>"										'사은품 정보 사용여부 | Y : 사용함, N : 사용안함
'		strRst = strRst & "	<ProductGift>"
'		strRst = strRst & "		<giftInfo/>"												'사은품 정보
'		strRst = strRst & "		<giftNm/>"
'		strRst = strRst & "		<aplBgnDt/>"
'		strRst = strRst & "		<aplEndDt/>"
'		strRst = strRst & "	</ProductGift>"
		strRst = strRst & "	<gblDlvYn>N</gblDlvYn>"											'#전세계배송 이용여부 | 선택하지 않을 경우, 기본 '이용안함(N)으로 세팅되며, 아래와 같은 조건이 모두 충족되어야 가능합니다. 1. 셀러회원정보에 전세계배송 이용여부가 "노출 또는 이용"으로 되어있고 2. 등록하려는 상품카테고리의 전세게배송 이용여부가 "이용(Y)"으로 되어있고 카테고리별 전세계배송 가능여부확인 3. 상품옵션이 "독립형"이 아니어야 하고 4. 상품의 배송방법이 '택배' 또는 '우편(소포/등기)'로 되어있고 5. 상품의 배송비설정이 '무료' 또는 결제방법이 '선결제가능' 혹은 '선결제 필수'이어야 하고 6. 통관용으로 사용될 생산지 국가를 반드시 선택해야 하고 7. 상품무게는 반드시 입력해야 하고 8. 상품의 출고지가 "국내주소" 인경우만 '전세계배송'이 가능. Y : 이용, N : 이용안함
'		strRst = strRst & "	<gblHsCode/>"													'전세계배송 HSCode | 전세계배송 "이용" 상품인경우 필수로 입력하셔야 합니다.
		strRst = strRst & "	<dlvCnAreaCd>02</dlvCnAreaCd>"									'#배송가능지역 코드 | 01 : 전국, 02 : *전국(제주 도서산간지역 제외), 03 : 서울, 04 : 인천, 05 : 광주, 06 : 대구, 07 : 대전, 08 : 부산, 09 : 울산, 10 : 경기, 11 : 강원, 12 : 충남, 13 : 충북, 14 : 경남, 15 : 경북, 16 : 전남, 17 : 전북, 18 : 제주, 19 : 서울/경기, 20 : 서울/경기/대전, 21 : 충북/충남, 22 : 경북/경남, 23 : 전북/전남, 24 : 부산/울산, 25 : 서울/경기/제주도서산간 제외지역, 26 : 일부지역불가
		strRst = strRst & "	<dlvWyCd>01</dlvWyCd>"											'#배송방법 | "전세계배송 상품" 인경우 '택배 또는 우편(소포/등기)'만 입력가능합니다. 01 : 택배, 02 : 우편(소포/등기), 03 : 직접전달(화물배달), 04 : 퀵서비스, 05 : 배송필요없음
		'2019-05-23 10:465 dlvSendCloseTmpltNo 추가
		strRst = strRst & "	<dlvSendCloseTmpltNo>570949</dlvSendCloseTmpltNo>"				'#발송마감 템플릿번호 | 발송마감 템플릿번호 (오늘발송, 일반발송) 1개 등록 가능하며 선택입력 정보입니다. 기본적으로 배송방법이 택배인 상품에 한하여 유효하며 해외직구 상품, 예약판매상품, 주문제작상품, 셀러위탁배송 상품은 반영 대상에서 제외됩니다.
		strRst = strRst & "	<dlvCstInstBasiCd>07</dlvCstInstBasiCd>"						'#배송비 종류 | 01 : 무료, 02 : 고정 배송비, 03 : 상품 조건부 무료, 04 : 수량별 차등, 05 : 1개당 배송비, 07 : 판매자 조건부 배송비 2010.08.20 06->07 로 변경, 08 : 출고지 조건부 배송비 2010.10.08 추가, 09 : 11번가 통합 출고지 배송비, 10 : 11번가해외배송조건부배송비 (11번가 해외 배송을 사용하는 경우)
'		strRst = strRst & "	<dlvCst1/>"														'배송비 | 상품 조건부 무료(03), 고정 배송비(02)
'		strRst = strRst & "	<dlvCst4/>"														'배송비 | 1개당 배송비(05)
'		strRst = strRst & "	<dlvCst3/>"														'배송비 | 수량별 차등(04) "수량별 차등"은 조건추가에 따라 배송비를 최대 10개까지 설정 가능
'		strRst = strRst & "	<dlvCstInfoCd/>"												'배송비 | 고정 배송비(02) 배송비 추가 안내 적용조건 : 전세계배송(N), 선결제불가(02) 01 : (상품상세참고), 02 : (상품별 차등 적용), 03 : (지역별 차등 적용), 04 : (상품/지역별 차등), 06 : (서울/경기 무료, 이외 추가비용)
'		strRst = strRst & "	<PrdFrDlvBasiAmt/>"												'상품조건부 무료 상품기준금액 | 상품조건부 무료(03)
'		strRst = strRst & "	<dlvCnt1/>"														'수량별 차등 기준 ~이상 수량 | 수량별 차등(04) "수량별 차등"은 조건추가에 따라 기준 수량를 최대 10개까지 설정 가능
'		strRst = strRst & "	<dlvCnt2/>"														'수량별 차등 기준 ~이하 수량 | 수량별 차등(04) "수량별 차등"은 조건추가에 따라 기준 수량를 최대 9개까지 설정 가능
		strRst = strRst & "	<bndlDlvCnYn>Y</bndlDlvCnYn>"									'#묶음배송 여부 | Y : 가능, N : 불가
		strRst = strRst & "	<dlvCstPayTypCd>03</dlvCstPayTypCd>"							'#결제방법 | 01 : 선결제가능, 02 : 선결제불가, 03 : 선결제필수
		strRst = strRst & "	<jejuDlvCst>3000</jejuDlvCst>"									'#제주 | 제주 추가 배송비
		strRst = strRst & "	<islandDlvCst>3000</islandDlvCst>"								'#도서산간 | 도서산간 추가 배송비
		strRst = strRst & "	<addrSeqOut>2</addrSeqOut>"										'#출고지 주소 코드 | 우선 SellerOffice 상품등록에서 출고지 주소가 등록이 되어있어야 합니다. 등록된 출고지 주소를 Api 조회 서비스를 통해 주소 시퀀스코드를 조회합니다. 출고지 주소조회에서 조회한 시퀀스코드를 입력하시면 됩니다. 만일 출고지 주소를 생략하실 경우 기본주소로 자동 설정이 됩니다. 하여 상품수정을 하실경우 수정당시의 기본주소로 재설정이 됩니다. 상품수정시 기본주소 변동으로 인한 이슈사항을 줄이기 위해 출고지 코드 입력을 권장합니다. 전세계배송이 되려면 출고지가 국내여야만 한다.
'		strRst = strRst & "	<outsideYnOut/>"												'출고지 주소 해외 여부 | "출고지 주소 해외 여부"는 출고지 주소가 해외일 경우에만 입력하시고 국내일 경우는 생략해 주세요. Y : 출고지 해외, N : 출고지 국내
'		strRst = strRst & "	<addrSeqOutMemNo/>"												'통합 ID 회원 번호 | 출고지용 "통합 ID 회원 번호 (출고지용)"는 통합 출고지 사용하는 경우에만 입력해 주세요.
		strRst = strRst & "	<addrSeqIn>3</addrSeqIn>"										'#반품/교환지 주소 코드 | 우선 SellerOffice 상품등록에서 반품/교환지 주소가 등록이 되어있어야 합니다. 등록된 반품/교환지 주소를 Api 조회 서비스를 통해 주소 시퀀스코드를 조회합니다. 반품/교환지 주소조회에서 조회한 시퀀스코드를 입력하시면 됩니다. 만일 반품/교환지 주소를 생략하실 경우 기본주소로 자동 설정이 됩니다. 하여 상품수정을 하실 경우 수정당시의 기본주소로 재설정이 됩니다. 상품수정시 기본주소 변동으로 인한 이슈사항을 줄이기 위해 출고지 코드 입력을 권장합니다.
'		strRst = strRst & "	<outsideYnIn/>"													'반품/교환지 주소 해외 여부 | "반품/교환지 주소 해외 여부"는 반품/교환지 주소가 해외일 경우에만 입력하시고 국내일 경우는 생략해 주세요. Y : 반품/교환지 해외, N : 반품/교환지 국내
'		strRst = strRst & "	<addrSeqInMemNo/>"												'통합 ID 회원 번호 | 반품지용 "통합 ID 회원 번호 (반품지용)"는 통합 반품지 사용하는 경우에만 입력해 주세요.
'		strRst = strRst & "	<abrdCnDlvCst/>"												'해외취소 배송비 | 배송비는 10원단위로 입력하셔야 합니다. 3000(O), 2999(X), 2,900(X) 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다.

''''''''교환/반품비 금액 변경..2020-01-20 김진영
'		strRst = strRst & "	<rtngdDlvCst>2500</rtngdDlvCst>"								'#반품 배송비 | 배송비는 10원단위로 입력하셔야 합니다. 3000(O), 2999(X), 2,900(X)
'		strRst = strRst & "	<exchDlvCst>5000</exchDlvCst>"									'#교환 배송비(왕복) | 배송비는 10원단위로 입력하셔야 합니다. 3000(O), 2999(X), 2,900(X)
		strRst = strRst & "	<rtngdDlvCst>3000</rtngdDlvCst>"								'#반품 배송비 | 배송비는 10원단위로 입력하셔야 합니다. 3000(O), 2999(X), 2,900(X)
		strRst = strRst & "	<exchDlvCst>6000</exchDlvCst>"									'#교환 배송비(왕복) | 배송비는 10원단위로 입력하셔야 합니다. 3000(O), 2999(X), 2,900(X)
''''''''교환/반품비 금액 변경..2020-01-20 김진영 끝

		strRst = strRst & "	<rtngdDlvCd>01</rtngdDlvCd>"									'초기배송비 무료시 부과방법 | 초기배송비 무료 시 부과방법 구분코드를 입력하지 않을 경우 편도반품배송비가 교환배송비보다 크거나 같은 경우에는 ’02’ 편도, 편도반품배송비가 교환배송비보다 작은 경우에는 ’01’왕복 코드가 자동 등록됩니다. 01 : 왕복(편도x2), 02 : 편도
		strRst = strRst & "	<asDetail><![CDATA[텐바이텐 고객행복센터 1644-6035]]></asDetail>"	'#A/S 안내 | 반드시 입력하셔야 하며 입력할 내용이 없으시면 . 이라도 입력해주셔야 합니다. 공백은 안됩니다. 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요.
		strRst = strRst & "	<rtngExchDetail><![CDATA[텐바이텐 고객행복센터 1644-6035]]></rtngExchDetail>"	'#반품/교환 안내 | 상품상세 페이지에 안내되는 내용으로, 반품/교환 문의를 줄이실 수 있습니다. 반드시 입력하셔야 하며 입력할 내용이 없으시면 . 이라도 입력해주셔야 합니다. 공백은 안됩니다. 특수문자 등이 포함되어 있을 경우 <![CDATA[ ]]> 로 묶어 주세요.
		strRst = strRst & "	<dlvClf>02</dlvClf>"											'#배송 주체 | 출고지에 따라 결정이 됩니다. 상품의 출고지가 판매자 출고지 인 경우: 업체 배송, 11번가 통합 ID의 출고지인 경우: 11번가 배송, 11번가 해외 통합 출고지인 경우: 11번가 해외배송, 01 : 11번가 배송 (통합 ID의 출고지를 사용하는 경우), 02 : 업체배송 (셀러가 배송을 처리하는 경우), 03 : 11번가 해외 배송 (11번가 해외 통합 출고지를 사용하는 경우) 지정하지 않는 경우 default로 02(업체배송)으로 처리됩니다.
'		strRst = strRst & "	<abrdInCd/>"													'#11번가 해외 입고 유형 | 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다. 상품의 출고지가 11번가 무료 픽업 가능 지역인 경우: 11번가 무료 픽업, 판매자 직접 발송인 경우: 판매자발송, 구매 대행인 경우: 구매 대행 01 : 11번가 무료 픽업, 02 : 판매자발송, 03 : 구매 대행
'		strRst = strRst & "	<prdWght/>"														'#상품 무게 | g 단위로 입력 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 또는 전세계배송 상품" 인경우 필수입니다.
'		strRst = strRst & "	<ntShortNm/>"													'#생산지국가(통관용) | 대표상품의 생산지국가를 선택하시면 됩니다. 전세계 배송상품인 경우 필수입니다. 아래 첨부파일을 참조 부탁 드립니다.
'		strRst = strRst & "	<globalOutAddrSeq/>"											'#판매자 해외 출고지 주소 | 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다. 우선 SellerOffice 상품등록에서 출고지 주소(해외)가 등록이 되어있어야 합니다. 등록된 출고지 주소를 Api 조회 서비스를 통해 주소 시퀀스코드를 조회합니다. 출고지 주소조회에서 조회한 시퀀스코드를 입력하시면 됩니다. 상품수정시 기본주소 변동으로 인한 이슈사항을 줄이기 위해 출고지 코드 입력을 권장합니다.
'		strRst = strRst & "	<mbAddrLocation05/>"											'#판매자 해외 출고지 지역 정보 | 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다. 해외 코드로 입력하시기 바랍니다. 01 : 국내, 02 : 해외
'		strRst = strRst & "	<globalInAddrSeq/>"												'#판매자 반품/교환지 주소 | 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다. 우선 SellerOffice 상품등록에서 반품/교환지 주소(해외)가 등록이 되어있어야 합니다. 등록된 반품/교환지 주소를 Api 조회 서비스를 통해 주소 시퀀스코드를 조회합니다. 반품/교환지 주소조회에서 조회한 시퀀스코드를 입력하시면 됩니다. 상품수정시 기본주소 변동으로 인한 이슈사항을 줄이기 위해 반품/교환지 코드 입력을 권장합니다.
'		strRst = strRst & "	<mbAddrLocation06>01</mbAddrLocation06>"						'#판매자 반품/교환지 지역 정보 | 배송 주체(dlvClf) 코드가 03(11번가 해외 배송)인 경우 필수입니다. 01 : 국내, 02 : 해외
'		strRst = strRst & "	<mnfcDy/>"														'제조일자
'		strRst = strRst & "	<eftvDy/>"														'유효일자
		strRst = strRst & get11stItemInfoCdParameter
		strRst = strRst & "	<company><![CDATA["&CStr(FMakerName)&"]]></company>"			'제조사 | 제조사는 텍스트 형태로만 입력하며 제조사가 없을 시 "없음"으로 입력합니다.
'		strRst = strRst & "	<modelNm/>"														'모델명 | 모델명은 텍스트 형태로만 입력하며 모델명이 없을 시 "없음"으로 입력합니다. (예시 모델명 : 기본 라인 셔츠 SQBAB9401)
'		strRst = strRst & "	<modelCd/>"														'모델코드 | 등록하실 상품 모델의 고유한 식별정보로, 영문+숫자, 숫자, 영문 등으로 조합된 모델번호를 입력해주십시오. (예시 모델코드 : SQBAB9401)
'		strRst = strRst & "	<mnfcDy/>"														'출판/출시일자 | 카테고리가 도서/음반/DVD인 경우 입력해 주셔야 합니다. 출판일자를 허위로 기재할경우, 도서문화진흥법에 의해 처벌 받을 수 있습니다.
'		strRst = strRst & "	<mainTitle/>"													'원제 | 카테고리가 도서인 경우 선택하여 입력하실 수 있습니다. 원제명은 한글50자, 영문/숫자 100자 이내로 입력해 주십시오.
'		strRst = strRst & "	<artist/>"														'아티스트/감독(배우) | 카테고리가 음반/DVD인 경우 반드시 입력해 주셔야 합니다.(음반,음반(TAPE),DVD/비디오) 아티스트/감독(배우)명은 한글 50자,영문/숫자 100자 이내로 입력해 주십시오.
'		strRst = strRst & "	<mudvdLabel/>"													'음반 라벨 | 카테고리가 도서/음반/DVD>음반, 음반[TAPE]인 경우 선택하여 입력하실 수 있습니다. 레이블은 한글 100자,영문/숫자 200자 이내로 입력해 주십시오.
'		strRst = strRst & "	<maker/>"														'제조사 | 카테고리가 도서/음반/DVD>음반, 음반[TAPE],DVD 인 경우 반드시 입력 하셔야 합니다.(음반,음반(TAPE),DVD/비디오) 레이블은 한글 100자,영문/숫자 200자 이내로 입력해 주십시오.
'		strRst = strRst & "	<albumNm/>"														'앨범명 | 카테고리가 도서/음반/DVD>음반, 음반[TAPE] 인 경우 반드시 입력 하셔야 합니다.(음반,음반(TAPE)) 앨범명은 한글 100자,영문/숫자 200자 이내로 입력해 주십시오
'		strRst = strRst & "	<dvdTitle/>"													'DVD 타이틀 | 카테고리가 도서/음반/DVD>DVD 인 경우 반드시 입력 하셔야 합니다.(DVD/비디오) DVD 타이틀은 한글 100자,영문/숫자 200자 이내로 입력해 주십시오.
'		strRst = strRst & "	<bcktExYn/>"													'장바구니 담기 제한 | 장바구니 담기 제한은 Y/N으로 입력 됩니다.
		strRst = strRst & "	<prcCmpExpYn>Y</prcCmpExpYn>"									'가격비교 사이트 등록 여부 | 가격비교사이트 등록은 선택사항이며 "등록함"을 권장합니다. Y : 등록함, N : 등록안함
		strRst = strRst & "</Product>"
		get11stItemRegParameter = strRst
'response.write strRst
'response.end
	End Function

	'상품정보제공고시
    Public Function get11stItemInfoCdParameter()
		Dim strSql, buf
		Dim mallinfoCd, infoContent, mallinfodiv, vType
		strSql = ""
		strSql = strSql & " SELECT top 100 M.* , "
		strSql = strSql & " CASE WHEN (M.infoCd='00001') THEN '상세정보 별도표기' "
		strSql = strSql & " 	 WHEN (M.infoCd='00002') THEN '상세페이지 참고' "
		strSql = strSql & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' "
		strSql = strSql & " 	 WHEN (M.infoCd='21011') AND Len(isNull(F.infocontent, '')) < 2 THEN I.itemname "
		strSql = strSql & " 	 WHEN c.infotype='P' THEN '텐바이텐 고객행복센터 1644-6035' "
		strSql = strSql & " 	WHEN LEN( isNull(F.infocontent, '')) < 2 THEN '상품 상세 참고' " & vbcrlf
		strSql = strSql & " ELSE F.infocontent + isNULL(F2.infocontent,'') END AS infocontent "
		strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F2 on M.infocdAdd = F2.infocd and F2.itemid='"&FItemID&"' "
		strSql = strSql & " WHERE M.mallid = '11st' and IC.itemid='"&FItemID&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = CInt(rsget("mallinfodiv"))
			vType = 891010 + mallinfodiv
			If mallinfodiv = "47" Then
				vType = "1149547"
			ElseIf mallinfodiv = "48" Then
				vType = "1149546"
			End If

			buf = buf & "	<ProductNotification>"												'상품정보제공고시
			buf = buf & "		<type>"&vType&"</type>"											'#유형코드
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
			    If Not (IsNULL(infoContent)) AND (infoContent <> "") Then
			    	infoContent = replaceRst(replace(infoContent, chr(31), ""))
				End If
				buf = buf & "			<item>"													'#항목정보 | 유형에 해당하는 항목정보
				buf = buf & "				<code><![CDATA["&mallinfoCd&"]]></code>"			'항목코드
				buf = buf & "				<name><![CDATA["&infoContent&"]]></name>"			'항목값 | 날짜입력 방식은 YYYY/MM/DD (년/월/일) 형식으로 입력해야 합니다.
				buf = buf & "			</item>"
				rsget.MoveNext
			Loop
			buf = buf & "	</ProductNotification>"
		End If
		rsget.Close
		get11stItemInfoCdParameter = buf
    End Function
End Class

Class C11st
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
	Public Sub get11stNotRegOneItem
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
		strSql = strSql & "	, isNULL(R.st11StatCD,-9) as st11StatCD "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, tm.safeDiv, tm.isNeed, tm.depth1Code, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " INNER JOIN (  "
		strSql = strSql & "		SELECT tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt "
		strSql = strSql & "		FROM db_etcmall.dbo.tbl_11st_cate_mapping "
		strSql = strSql & "		GROUP BY tenCateLarge, tenCateMid, tenCateSmall "
		strSql = strSql & " ) as cm on cm.tenCateLarge = i.cate_large and cm.tenCateMid = i.cate_mid and cm.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_regItem as R on i.itemid = R.itemid"
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
		strSql = strSql & " and i.itemdiv in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
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
'		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and ( "
		strSql = strSql & "			convert(varchar(6), (i.cate_large + i.cate_mid)) not in ( "
		strSql = strSql & "				SELECT convert(varchar(6), cdl+cdm)  "
		strSql = strSql & "				FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category]  "	'2023-06-23 김진영 / 등록제외 카테고리라도 특정브랜드는 판매되도록
		strSql = strSql & "				WHERE mallgubun='"&CMALLNAME&"' "
		strSql = strSql & "		) or i.makerid in ( "
		strSql = strSql & "			'heidi2022', "
		strSql = strSql & "			'luna2022', "
		strSql = strSql & "			'uand2051', "
		strSql = strSql & "			'wpc001', "
		strSql = strSql & "			'lifeshop0510', "
		strSql = strSql & "			'bijou2023', "
		strSql = strSql & "			'blesscompany', "
		strSql = strSql & "			'JINNYSTAR01', "
		strSql = strSql & "			'sportsconnection', "
		strSql = strSql & "			'ithinkso', "
		strSql = strSql & "			'greenh03', "
		strSql = strSql & "			'kingkongoutlet', "
		strSql = strSql & "			'shoemiz', "
		strSql = strSql & "			'goldn', "
		strSql = strSql & "			'funnyfun', "
		strSql = strSql & "			'doran1020', "
		strSql = strSql & "			'gabangpop1010', "
		strSql = strSql & "			'osjarak' "
		strSql = strSql & "		) "
		strSql = strSql & "	)  "
		strSql = strSql & "	and i.itemid not in (SELECT itemid FROM db_etcmall.dbo.tbl_11st_regItem WHERE st11StatCD >= 3) "	''등록완료이상은 등록안됨.	'11st등록상품 제외
		strSql = strSql & " and cm.mapCnt is Not Null "'	카테고리 매칭 상품만
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
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
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FSafetyyn			= rsget("safetyyn")
				FOneItem.FSafetyNum			= rsget("safetyNum")
				FOneItem.FSafetyDiv			= rsget("safetyDiv")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")
				FOneItem.Fdeliverytype		= rsget("deliverytype")
				FOneItem.FDepthCode			= rsget("depthCode")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

	Public Sub get11stEditOneItem
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
		strSql = strSql & "	, m.st11GoodNo, m.st11price, m.st11SellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, UC.socname_kor, am.depthCode, isNULL(m.st11StatCD,-9) as st11StatCD, tm.safeDiv, tm.isNeed, tm.depth1Code, isNull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000)) "
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & "		or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & " 	or i.itemdiv not in ('01', '06', '16', '07') "		'01 : 일반, 06 : 주문제작(문구), 16 : 주문제작, 07 : 구매제한
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
'		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or (( "
		strSql = strSql & "			convert(varchar(6), (i.cate_large + i.cate_mid)) in ( "
		strSql = strSql & "				SELECT convert(varchar(6), cdl+cdm)  "
		strSql = strSql & "				FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category]  "	'2023-06-23 김진영 / 등록제외 카테고리라도 특정브랜드는 판매되도록
		strSql = strSql & "				WHERE mallgubun='11st1010') "
		strSql = strSql & "			) and ( "
		strSql = strSql & "				i.makerid not in ( "
		strSql = strSql & "					'heidi2022', "
		strSql = strSql & "					'luna2022', "
		strSql = strSql & "					'uand2051', "
		strSql = strSql & "					'wpc001', "
		strSql = strSql & "					'lifeshop0510', "
		strSql = strSql & "					'bijou2023', "
		strSql = strSql & "					'blesscompany', "
		strSql = strSql & "					'JINNYSTAR01', "
		strSql = strSql & "					'sportsconnection', "
		strSql = strSql & "					'ithinkso', "
		strSql = strSql & "					'greenh03', "
		strSql = strSql & "					'kingkongoutlet', "
		strSql = strSql & "					'shoemiz', "
		strSql = strSql & "					'goldn', "
		strSql = strSql & "					'funnyfun', "
		strSql = strSql & "					'doran1020', "
		strSql = strSql & "					'gabangpop1010', "
		strSql = strSql & "					'osjarak' "
		strSql = strSql & "				) "
		strSql = strSql & "			) "
		strSql = strSql & "		) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_cate_mapping as am on am.tenCateLarge=i.cate_large and am.tenCateMid=i.cate_mid and am.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_11st_category as tm on am.depthCode = tm.depthCode "
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.st11GoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new C11stItem
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
				FOneItem.FSt11GoodNo		= rsget("st11GoodNo")
				FOneItem.FSt11price			= rsget("st11price")
				FOneItem.FSt11SellYn		= rsget("st11SellYn")

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
                FOneItem.FDepthCode			= rsget("depthCode")
                FOneItem.Fvatinclude        = rsget("vatinclude")
				FOneItem.FSt11StatCD		= rsget("st11StatCD")
				FOneItem.Fdeliverfixday		= rsget("deliverfixday")

				FOneItem.FSafeDiv 			= rsget("safeDiv")
				FOneItem.FIsNeed 			= rsget("isNeed")
				FOneItem.FDepth1Code 		= rsget("depth1Code")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FOutmallstandardMargin	= rsget("outmallstandardMargin")
		End If
		rsget.Close
	End Sub

End Class

'11번가 상품코드 얻기
Function get11stGoodno(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 st11goodno FROM db_etcmall.dbo.tbl_11st_regitem WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		get11stGoodno = rsget("st11goodno")
	End If
	rsget.Close
End Function

'11번가 상품코드/상품가 얻기
Function get11stGoodno2(iitemid, ist11goodno, byref MustPrice)
	Dim strRst, strSql
	Dim sellcash, orgprice, buycash, saleyn, tmpPrice, vdeliverytype, ispecialPrice, outmallstandardMargin
	Dim GetTenTenMargin, st11goodno, ownItemCnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.sellcash, i.buycash, i.orgprice, i.sailyn, r.st11goodno, i.deliverytype, isnull(mi.mustPrice, 0) as specialPrice, isnull(f.outmallstandardMargin, "& CMAXMARGIN &") as outmallstandardMargin "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_outmall_mustPriceItem] as mi "
	strSql = strSql & " 	on i.itemid = mi.itemid "
	strSql = strSql & " 	and mi.mallgubun = '11st1010' "
	strSql = strSql & " 	and (GETDATE() >= mi.startDate and GETDATE() <= mi.endDate ) "
	strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		sellcash	= rsget("sellcash")
		orgprice	= rsget("orgprice")
		buycash		= rsget("buycash")
		saleyn		= rsget("sailyn")
		st11goodno	= rsget("st11goodno")
		vdeliverytype = rsget("deliverytype")
		ispecialPrice = rsget("specialPrice")
		outmallstandardMargin = rsget("outmallstandardMargin")
	Else
		get11stGoodno2 = ""
		Exit Function
		response.end
	End If
	rsget.close

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as CNT "
	strSql = strSql & " FROM db_item.dbo.tbl_item i "
	strSql = strSql & " JOIN db_partner.dbo.tbl_partner p on i.makerid = p.id "
	strSql = strSql & " WHERE p.purchaseType in (3, 5, 6) "		'3 : PB, 5 : ODM, 6 : 수입
	strSql = strSql & " and i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		ownItemCnt = rsget("CNT")
	End If
	rsget.Close

	If ispecialPrice <> "0" Then
		tmpPrice = ispecialPrice
	ElseIf ownItemCnt > 0 Then
		tmpPrice = orgprice
	Else
		GetTenTenMargin = CLng((10000 - buycash / sellcash * 100 * 100) / 100)
		If (GetTenTenMargin < outmallstandardMargin) Then
			tmpPrice = orgprice
		Else
			tmpPrice = sellcash
		End If
	End If
	MustPrice = CStr(GetRaiseValue(tmpPrice/10)*10)
	ist11goodno = st11goodno
End Function

'11번가 상품코드, 옵션 수 얻기
Function get11stGoodno3(iitemid, ist11goodno, byref opCnt)
	Dim strSql, st11goodno, optioncnt

	strSql = ""
	strSql = strSql & " SELECT TOP 1 i.optioncnt, r.st11goodno "
	strSql = strSql & " FROM db_item.dbo.tbl_item as i "
	strSql = strSql & " JOIN db_etcmall.dbo.tbl_11st_regitem as r on i.itemid = r.itemid "
	strSql = strSql & " WHERE i.itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		opCnt		= rsget("optioncnt")
		ist11goodno	= rsget("st11goodno")
	Else
		get11stGoodno3 = ""
		Exit Function
		response.end
	End If
	rsget.close
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
%>
