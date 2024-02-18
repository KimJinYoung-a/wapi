<%
CONST CMAXMARGIN = 10
CONST CMALLGUBUN = "naverep"
CONST CMALLNAME = "nvstoregift"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CDEFALUT_STOCK = 9999

Class CNvstoregiftItem
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
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FNvstoregiftGoodNo
	Public FNvstoregiftprice
	Public FNvstoregiftSellYn
	Public FregedOptCnt
	Public FAccFailCNT
	Public FMaySoldOut
	Public Fregitemname
	Public FLastErrStr
	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FNvstoregiftStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FSocname_kor
	Public FAPIaddImg
	Public FbasicimageNm
	Public FRegImageName
	Public FCateKey
	Public FNeedCert
	Public FAdultType
	Public FNvstorefarmid
	Public FOrderMaxNum
	Public FPurchasetype

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "9999999999" Then
			getOrderMaxNum = 9999999999
		End If
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function fnIsSpecialDate
		Dim sqlStr, specialPrice, cnt, cnt2
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
'		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			cnt = rsget("CNT")
		End If
		rsget.Close

		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as CNT "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			cnt2 = rsget("CNT")
		End If
		rsget.Close

		If cnt > 0 and cnt2 > 0 Then
			fnIsSpecialDate = "YY"
		ElseIf cnt > 0 and cnt2 = 0 Then
			fnIsSpecialDate = "YN"
		Else
			fnIsSpecialDate = "NN"
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, sqlStr, specialPrice
		' sqlStr = ""
		' sqlStr = sqlStr & " SELECT mustPrice "
		' sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		' sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		' sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		' sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		' rsget.Open sqlStr,dbget,1
		' If Not(rsget.EOF or rsget.BOF) Then
		' 	specialPrice = rsget("mustPrice")
		' 	fnIsSpecialDate = "Y"
		' End If
		' rsget.Close

		' If specialPrice <> "" Then
		' 	MustPrice = specialPrice
		' Else
		If FPurchasetype = "3" OR FPurchasetype = "5" OR FPurchasetype = "6" Then
			MustPrice = Forgprice
		Else
			GetTenTenMargin = CLng((10000 - Fbuycash / FSellCash * 100 * 100) / 100)
			If GetTenTenMargin < CMAXMARGIN Then
				MustPrice = Forgprice
			Else
				MustPrice = FSellCash
			End If
		End If
	End Function

    public function getBasicImage()
		Dim uploadBasicImage, strSQL
		if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 IMAGENAME "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_nvstorefarm_uploadimage "
		strSQL = strSQL & " WHERE ITEMID = '"& Fitemid &"' "
		strSQL = strSQL & " AND GUBUN = 1 "
		strSQL = strSQL & " ORDER BY GUBUN ASC "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			uploadBasicImage = rsget("IMAGENAME")
		End If
		rsget.Close

		If uploadBasicImage = "" Then
	        getBasicImage = FbasicImageNm
		Else
			getBasicImage = uploadBasicImage
		End If
    end function

	'// 스토어팜 판매여부 반환
	Public Function getNvstoregiftSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FSellYn="Y" and FIsUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getNvstoregiftSellYn = "Y"
			Else
				getNvstoregiftSellYn = "N"
			End If
		Else
			getNvstoregiftSellYn = "N"
		End If
	End Function

	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold < CMAXLIMITSELL))
	End Function

	Public Function IsMayLimitSoldout
		If FOptionCnt = 0 Then
			Exit Function
		End If
		Dim sqlStr, optLimit, limitYCnt, optaddprice, optAddpriceHalfOverCnt
		optAddpriceHalfOverCnt = 0
		sqlStr = ""
		sqlStr = sqlStr & " SELECT itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
		sqlStr = sqlStr & " FROM [db_item].[dbo].tbl_item_option "
		sqlStr = sqlStr & " WHERE isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				optLimit = rsget("optLimit")
				optaddprice = rsget("optaddprice")
				optLimit = optLimit-5
				If (optLimit < 1) Then optLimit = 0
				If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK

				If (optLimit <> 0) Then
					limitYCnt =  limitYCnt + 1
				End If

				'2020-01-31 김진영.. 옵션추가금액이 판매가보다 50%이상 비싸면 품절처리
				If optaddprice >= (MustPrice * 0.5) Then
					optAddpriceHalfOverCnt = optAddpriceHalfOverCnt + 1
				End If
				rsget.MoveNext
			Loop
		End If
		rsget.Close

		If (limitYCnt = 0) OR (optAddpriceHalfOverCnt > 0) Then
			IsMayLimitSoldout = "Y"
		Else
			IsMayLimitSoldout = "N"
		End If
	End Function


	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function getLimitNvstoregiftEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold - 5
			If ret > 10000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimitNvstoregiftEa = ret
	End Function

	Public Function getSalePrice
		Dim sqlStr, mustCnt
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as cnt "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			mustCnt = rsget("cnt")
		End If
		rsget.Close

		If mustCnt > 0 Then
			getSalePrice = Forgprice
		Else
			getSalePrice = Clng(GetRaiseValue(MustPrice/10)*10)
		End If
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

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

	Function getItemNameFormat()
		Dim buf
		buf = "[텐바이텐]"&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
	End Function

	Public Function getModelName
		Dim strSql, modelName, isRegCert, safetyDiv, safetyId
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN '121' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN '72' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN '1042' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN '51' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN '1020' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN '58' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN '1040' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN '1041' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN '1042' end as safetyId "
		strSql = strSql & " ,f.modelName "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_tenReg] as t on i.itemid = t.itemid "
		strSql = strSql & " JOIN db_item.[dbo].[tbl_safetycert_info] as f on t.itemid = f.itemid "
		strSql = strSql & " WHERE i.itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			safetyDiv		= rsget("safetyDiv")
			safetyId		= rsget("safetyId")
			modelName		= rsget("modelName")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		If isRegCert = "Y" and safetyDiv = "70" OR safetyDiv = "80" OR safetyDiv = "90" Then
			getModelName = "					<shop:ModelName>"&modelName&"</shop:ModelName>"
		Else
			getModelName = ""
		End If
	End Function

	'주문 제작 정보
    Public Function getzCostomMadeInd()
		Dim buf, CustomMade, EstimatedDeliveryTime
		If (Fitemdiv="06" or Fitemdiv="16") Then
			If (FrequireMakeDay > 5) Then
				EstimatedDeliveryTime = FrequireMakeDay
			ElseIf (FrequireMakeDay < 1) Then
				EstimatedDeliveryTime = 5
			Else
				EstimatedDeliveryTime = FrequireMakeDay + 1
			End If
			CustomMade = "Y"
		Else
			CustomMade = "N"
		End If

		buf = ""
		If CustomMade = "Y" Then
'			buf = buf & "				<shop:CustomMade>Y</shop:CustomMade>"		'# 주문 제작 상품 여부 Y or N | Y: EstimatedDeliveryTime입력 필수, "N": EstimatedDeliveryTime 입력 불가
'			buf = buf & "				<shop:UseReturnCancelNotification>Y</shop:UseReturnCancelNotification>"		'주문 제작 상품 반품/취소 제한 안내 여부
			buf = buf & "				<shop:EstimatedDeliveryTime>"&EstimatedDeliveryTime&"</shop:EstimatedDeliveryTime>"
		Else
'			buf = buf & "				<shop:CustomMade>N</shop:CustomMade>"		'# 주문 제작 상품 여부 Y or N | Y: EstimatedDeliveryTime입력 필수, "N": EstimatedDeliveryTime 입력 불가
			'buf = buf & "				<shop:EstimatedDeliveryTime></shop:EstimatedDeliveryTime>"
		End If
		getzCostomMadeInd = buf
    End Function

	'원산지 정보
	Public Function getOriginAreaType
		Dim buf
		buf = ""
		buf = buf & "				<shop:OriginArea>"													'#원산지 정보
		If Fsourcearea = "한국" OR Fsourcearea = "대한민국" OR Fsourcearea = "국산" Then
			buf = buf & "					<shop:Code>00</shop:Code>"									'#원산지 상세 지역 | 00 : 국산, 01 : 원양산, 02 : 수입산, 03 : 상세설명에 표시, 04 : 직접입력
'			buf = buf & "					<shop:Importer></shop:Importer>"							'수입사명 | 수입산인 경우 필수
			buf = buf & "					<shop:Plural>N</shop:Plural>"								'복수 원산지 | Y or N
'			buf = buf & "					<shop:Content></shop:Content>"								'원산지 표시 내용 | Code가 "기타:직접 입력"인 경우 필수
		Else
			buf = buf & "					<shop:Code>04</shop:Code>"									'#원산지 상세 지역 | 00 : 국산, 01 : 원양산, 02 : 수입산, 03 : 상세설명에 표시, 04 : 직접입력
'			buf = buf & "					<shop:Importer></shop:Importer>"							'수입사명 | 수입산인 경우 필수
			buf = buf & "					<shop:Plural>N</shop:Plural>"								'복수 원산지 | Y or N
			buf = buf & "					<shop:Content><![CDATA["&Fsourcearea&"]]></shop:Content>"	'원산지 표시 내용 | Code가 "기타:직접 입력"인 경우 필수
		End If
		buf = buf & "				</shop:OriginArea>"
		getOriginAreaType = buf
	End Function

	Public Function getNvstoregiftCertInfo
		Dim buf, strSql, safetyDiv, safetyId, certNum, certOrganName, certmakerName, isRegCert
		strSql = ""
		strSql = strSql & " select top 1 i.itemid, t.safetyDiv "
		strSql = strSql & " ,Case When t.safetyDiv = '10' THEN '121' "
		strSql = strSql & " 	When t.safetyDiv = '20' THEN '72' "
		strSql = strSql & " 	When t.safetyDiv = '30' THEN '1042' "
		strSql = strSql & " 	When t.safetyDiv = '40' THEN '51' "
		strSql = strSql & " 	When t.safetyDiv = '50' THEN '1020' "
		strSql = strSql & " 	When t.safetyDiv = '60' THEN '58' "
		strSql = strSql & " 	When t.safetyDiv = '70' THEN '1040' "
		strSql = strSql & " 	When t.safetyDiv = '80' THEN '1041' "
		strSql = strSql & " 	When t.safetyDiv = '90' THEN '1042' end as safetyId "
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
			certmakerName	= rsget("makerName")
			isRegCert		= "Y"
		Else
			isRegCert		= "N"
		End If
		rsget.Close

		If isRegCert = "Y" Then
			buf = buf & "				<shop:CertificationList>"												'인증 정보 목록	| 선택이나, 입력할 경우 하단 #은 필수
			buf = buf & "					<shop:Certification>"
			buf = buf & "						<shop:Id>"&safetyId&"</shop:Id>"								'인증 유형 ID
			buf = buf & "						<shop:Name>"&certOrganName&"</shop:Name>"						'#인증 기관
			buf = buf & "						<shop:Number>"&certNum&"</shop:Number>"							'#인증 번호
			buf = buf & "						<shop:Mark>Y</shop:Mark>"										'인증 마크 사용 여부 | Y or N, 미입력시 N
			If safetyDiv = "70" OR safetyDiv = "80" OR safetyDiv = "90" Then
				buf = buf & "					<shop:CompanyName>"&certmakerName&"</shop:CompanyName>"			'인증 상호 | 인증 유형이 방송통신기자재 적합인증/적합등록/잠정인증인 경우 필수
				buf = buf & "					<shop:KindType>CHI</shop:KindType>"								'인증정보 종류 미입력시 ETC | KC : KC 인증, CHI : 어린이제품 인증, GRN : 친환경인증, PARALLEL_IMPORT : 병행수입(면제대상), ETC : 기타 인증
			Else
				buf = buf & "					<shop:KindType>KC</shop:KindType>"								'인증정보 종류 미입력시 ETC | KC : KC 인증, CHI : 어린이제품 인증, GRN : 친환경인증, PARALLEL_IMPORT : 병행수입(면제대상), ETC : 기타 인증
			End If
			buf = buf & "					</shop:Certification>"
			buf = buf & "				</shop:CertificationList>"
'			If safetyDiv = "70" OR safetyDiv = "80" OR safetyDiv = "90" Then
'				buf = buf & "				<shop:ChildCertifiedProductExclusion>N</shop:ChildCertifiedProductExclusion>"
'			End If
			getNvstoregiftCertInfo = buf
		Else
			getNvstoregiftCertInfo = ""
		End If
	End Function

	'// 상품등록: 상품추가이미지 파라메터 생성
	Public Function getImageType()
		Dim buf, strSql, arrRows, i, basicimgStr, addimgStr
		addimgStr	= ""
		basicimgStr	= ""
		strSql = ""
		strSql = strSql & " SELECT TOP 10 imgType, storefarmURL FROM db_etcmall.[dbo].[tbl_nvstoregift_Image] WHERE itemid = '"&FItemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			arrRows = rsget.getRows()
		End If
		rsget.Close

		If isArray(arrRows) then
			For i = 0 To UBound(arrRows, 2)
				If arrRows(0, i) = "1" Then
					basicimgStr = arrRows(1,i)																		'대표 이미지
				Else
					addimgStr = addimgStr & "						<shop:Optional>"								'추가 이미지
					addimgStr = addimgStr & "							<shop:URL>"&arrRows(1,i)&"</shop:URL>"
					addimgStr = addimgStr & "						</shop:Optional>"
				End If
			Next
		End If

		buf = ""
		buf = buf & "				<shop:Image>"
		buf = buf & "					<shop:Representative>"
		buf = buf & "						<shop:URL>"&basicimgStr&"</shop:URL>"
		buf = buf & "					</shop:Representative>"
		If addimgStr <> "" Then
		buf = buf & "					<shop:OptionalList>"
		buf = buf & addimgStr
		buf = buf & "					</shop:OptionalList>"
		End If
		buf = buf & "				</shop:Image>"
		getImageType = buf
	End Function

	Public Function getDeliveryType
		Dim buf, deliveryPay
		deliveryPay = "50000"
		buf = ""
		buf = buf & "				<shop:Delivery>"
		buf = buf & "					<shop:Type>1</shop:Type>"												'#배송 방법 | 1 : 택배, 소포, 등기, 2 : 직접 배송(화물 배달)
		buf = buf & "					<shop:BundleGroupAvailable>Y</shop:BundleGroupAvailable>"				'#묶음 배송 가능 여부 | Y or N..묶음 배송 그룹 코드가 존재할 경우 자동으로 Y로 설정된다.
'		buf = buf & "					<shop:BundleGroupId></shop:BundleGroupId>"								'묶음 배송 그룹 코드 | 묶음 배송 가능이 Y고 묶음 배송 그룹 코드가 Null이면 기본 그룹으로 저장된다.
'		buf = buf & "					<shop:VisitAddressId></shop:VisitAddressId>"							'방문 수령 주소 코드
'		buf = buf & "					<shop:QuickServiceAreaList>"											'퀵서비스 배송 지역 코드 목록
'		buf = buf & "						<shop:QuickServiceAreaCode></shop:QuickServiceAreaCode>"
'		buf = buf & "					</shop:QuickServiceAreaList>"
		buf = buf & "					<shop:FeeType>2</shop:FeeType>"											'#배송비 유형 | 1 : 무료. 2 : 조건부 무료, 3 : 유료, 4 : 수량별 부과 - 반복구간, 5 : 수량별 부과 - 구간 직접 설정
		buf = buf & "					<shop:BaseFee>3000</shop:BaseFee>"										'#기본 배송비
		buf = buf & "					<shop:FreeConditionalAmount>"&deliveryPay&"</shop:FreeConditionalAmount>"	'무료 조건 금액 | 배송비 유형이 '조건부 무료'일 경우 입력한다.
'		buf = buf & "					<shop:RepeatQuantity></shop:RepeatQuantity>"							'반복 수량 | 배송비 유형이 수량별 부과 - 반복구간일 경우 입력한다.
'		buf = buf & "					<shop:SecondBaseQuantity></shop:SecondBaseQuantity>"					'2구간 최소수량 | 배송비 유형이 수량별 부과 - 구간직접 설정 일 경우 입력한다.
'		buf = buf & "					<shop:SecondExtraFee></shop:SecondExtraFee>"							'2구간 추가 배송비 | 배송비 유형이 수량별 부과 - 구간직접 설정 일 경우 입력한다.
'		buf = buf & "					<shop:ThirdBaseQuantity></shop:ThirdBaseQuantity>"						'3구간 최소수량 | 배송비 유형이 수량별 부과 - 구간직접 설정 일 경우 입력한다.
'		buf = buf & "					<shop:ThirdExtraFee></shop:ThirdExtraFee>"								'3구간 추가 배송비 | 배송비 유형이 수량별 부과 - 구간직접 설정 일 경우 입력한다.
		buf = buf & "					<shop:PayType>"&Chkiif(FdeliveryType = "7", "1", "2")&"</shop:PayType>"	'배송비 결제 방식 | 1 : 착불, 2 : 선결제, 3 : 착불 또는 선결제
'		buf = buf & "					<shop:AreaType></shop:AreaType>"										'지역별 추가 배송 권역 | 2 : 2권역 - 내륙/제주 및 도서 산간 지역으로 구분, 3 : 3권역 - 내륙/제주 외 도서 산간 지역으로 구분..묶음 배송 가능이 Y인 경우에는 해당 값은 무시 된다.
'		buf = buf & "					<shop:Area2ExtraFee></shop:Area2ExtraFee>"								'2권역 배송비 |  묶음 배송 가능이 Y인 경우에는 해당 값은 무시 된다.
'		buf = buf & "					<shop:Area3ExtraFee></shop:Area3ExtraFee>"								'3권역 배송비 |  묶음 배송 가능이 Y인 경우에는 해당 값은 무시 된다.
		buf = buf & "					<shop:ReturnDeliveryCompanyPriority>0</shop:ReturnDeliveryCompanyPriority>"	'#반품/교환 택배사 | 0 : 기본 반품 택배사, 1 : 보조 반품 택배사1, 2 :보조 반품 택배사2..부터 9까지
		buf = buf & "					<shop:ReturnFee>3000</shop:ReturnFee>"									'#반품 배송비
		buf = buf & "					<shop:ExchangeFee>6000</shop:ExchangeFee>"								'#교환 배송비
		buf = buf & "					<shop:DeliveryInsurance>N</shop:DeliveryInsurance>"						'반품안심케어 설정 | Y 또는 N 대상이 되는 경우에만 해당된다
'		buf = buf & "					<shop:ShippingAddressId></shop:ShippingAddressId>"						'출고지 주소 번호
'		buf = buf & "					<shop:ReturnAddressId></shop:ReturnAddressId>"							'반품/교환지 주소 번호
'		buf = buf & "					<shop:DifferentialFee></shop:DifferentialFee>"							'지역별 차등 배송비 정보
'		buf = buf & "					<shop:InstallationFee></shop:InstallationFee>"							'별도 설치비 유무 | Y or N
'		buf = buf & "					<shop:IndividualCustomUniqueCode></shop:IndividualCustomUniqueCode>"	'개인통관 고유부호 수집 여부 Y or N | 미입력시 Y로 저장
		buf = buf & "					<shop:AttributeType>NORMAL</shop:AttributeType>"						'배송속성타입코드 | NORMAL : 일반배송, TODAY : 오늘출발, OPTION_TOPDAY : 옵션별 오늘출발
		buf = buf & "					<shop:DeliveryCompany>HYUNDAI</shop:DeliveryCompany>"					'택배사 코드 | 배송 방법(Type) 필드값이 1(택배, 소포, 등기)이면 반드시 입력해야 한다
	If FItemdiv = "06" OR FItemdiv = "16" Then
		Dim mayDeliverDay
'		buf = buf & "					<shop:ExpectedDeliveryPeriodType>ETC</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
		Select Case FCateKey
			Case "50000980", "50000979", "50003503", "50003509", "50003505", "50003504", "50003510", "50003506", "50003502", "50003508", "50003501", "50003507", "50003334", "50003357", "50003330", "50003347", "50003332", "50003343", "50003336", "50003333", "50006876", "50003345", "50003355", "50003356", "50003344", "50003352", "50003335", "50003338", "50003340", "50003331", "50003341", "50003337", "50003353", "50003354", "50003349", "50006875", "50003346", "50003348", "50003350", "50003351", "50003339", "50003342", "50003474", "50003473", "50003471", "50003475", "50003476", "50003477", "50003478", "50003472", "50003470", "50000976", "50000973", "50000974", "50000969", "50000972", "50000975", "50000970", "50003480", "50000971", "50003479", "50003481", "50003482", "50003499", "50003500", "50003489", "50003493", "50003497", "50003485", "50003488", "50003492", "50003496", "50003484", "50003487", "50003491", "50003495", "50003483", "50003490", "50003494", "50003498", "50003486", "50000977", "50000983", "50000981", "50000982", "50003511", "50003513", "50003512", "50000984", "50000852", "50000864", "50000857", "50000861", "50000862", "50003514", "50003515", "50003516", "50000860", "50000859", "50000865", "50000853", "50000866", "50000863", "50000858", "50000867", "50006168", "50003643", "50003644", "50003647", "50003641", "50003521", "50003649", "50003648", "50003526", "50003651", "50003652", "50003645", "50003517", "50003524", "50003522", "50003523", "50003527", "50003518", "50003525", "50003642", "50003646", "50003640", "50003656", "50003655", "50003654", "50003519", "50003653", "50003520", "50003650", "50001066", "50001064", "50003307", "50003308", "50003309", "50003314", "50003321", "50001061", "50003310", "50003311", "50003315", "50003316", "50003312", "50003313", "50003317", "50003318", "50003319", "50003320", "50001062", "50003322", "50003684", "50003685", "50003686", "50003687", "50003688", "50003689", "50001067", "50001063", "50003328", "50003691", "50003327", "50003690", "50003325", "50003694", "50003324", "50003693", "50003696", "50001065", "50003329", "50003326", "50003692", "50003323"
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>TEN</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "10"
			Case "50001311", "50003242", "50003251", "50003254", "50003245", "50003248", "50003250", "50003253", "50003247", "50003249", "50003243", "50003246", "50003244", "50001310", "50003677", "50003294", "50003680", "50003683", "50003681", "50003266", "50003291", "50003295", "50003682", "50003264", "50003675", "50003265", "50003676", "50003695", "50003678", "50003679", "50003293", "50003292", "50003263", "50003304", "50003306", "50001347", "50003298", "50003300", "50003299", "50003296", "50003297", "50003302", "50003301", "50003305", "50003303", "50001346", "50001323", "50001322", "50001324", "50001321", "50001330", "50001320", "50001326", "50001327", "50001328", "50001329", "50001319", "50001325", "50001339", "50001342", "50001344", "50001340", "50001345", "50001337", "50001341", "50001336", "50001332", "50001333", "50001334", "50003671", "50001335", "50003672", "50003673", "50003262", "50001338", "50001059", "50001058", "50001055", "50001054", "50001056", "50001057", "50001317", "50001318", "50001313", "50001314", "50001315", "50003258", "50003255", "50003260", "50003256", "50003257", "50003261", "50001316", "50001308", "50003241", "50003670", "50003224", "50003212", "50003219", "50003231", "50003221", "50003237", "50003229", "50003240", "50003223", "50003211", "50003222", "50003210", "50003227", "50003217", "50003216", "50003232", "50003228", "50003238", "50003218", "50003234", "50003215", "50003233", "50003225", "50003213", "50003226", "50003214", "50003239", "50003230", "50006200", "50003236", "50003235", "50003220", "50001309", "50001307", "50001060"
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>FOURTEEN</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "15"
			Case "50002542", "50002539", "50002545", "50002547", "50002543", "50006838", "50002514", "50002541", "50002544", "50002513", "50002512", "50002519", "50001851", "50006201", "50006836", "50006837", "50002511", "50002538", "50002516", "50002540", "50002515", "50002517", "50002546", "50002518", "50006848", "50001853", "50001521", "50001852", "50006835", "50002523", "50002522", "50002521", "50002526", "50002549", "50002527", "50002548", "50002525", "50002524", "50003207", "50002535", "50002534", "50002529", "50002554", "50002532", "50002536", "50002530", "50002531", "50002553", "50002550", "50002556", "50002557", "50002537", "50002533", "50002528", "50002555", "50002552", "50002551", "50000264", "50006370", "50000262", "50006371", "50006369", "50000258", "50000260", "50000259", "50000255", "50000254", "50000257", "50000253", "50000256", "50000252", "50000263", "50000261"
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>SEVEN</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "7"
			Case "50000846", "50000847", "50000848", "50003808", "50003810", "50003804", "50003805", "50003807", "50003809", "50003806", "50000845", "50000831", "50000844", "50000836", "50000833", "50000843", "50000838", "50000837", "50000840", "50000834", "50000835", "50000832", "50006328", "50000839", "50000841", "50000830", "50000842", "50000774", "50000775", "50000772", "50000773", "50000777", "50003813", "50003811", "50000776", "50003812", "50000771", "50000769", "50000750", "50000768", "50000757", "50000770", "50000755", "50000766", "50000749", "50000761", "50000748", "50000753", "50000767", "50000752", "50000760", "50000759", "50000756", "50000762", "50000754", "50000751", "50006828", "50000758", "50000763", "50000747", "50000764", "50000765", "50000828", "50000823", "50000825", "50003636", "50000829", "50003816", "50003639", "50003637", "50003638", "50000827", "50000826", "50003627", "50003631", "50003632", "50003624", "50003635", "50003630", "50003626", "50003625", "50003628", "50003634", "50000824", "50000805", "50000812", "50000822", "50000810", "50000804", "50000808", "50000807", "50000821", "50000815", "50000814", "50000811", "50000816", "50000817", "50000809", "50000806", "50000778", "50000813", "50000818", "50000803", "50000820", "50000819", "50000651", "50000650", "50000646", "50000648", "50000649", "50000647", "50000652", "50000787", "50000784", "50000785", "50000792", "50003863", "50003860", "50000789", "50003861", "50003862", "50000788", "50003854", "50003855", "50003857", "50003858", "50003859", "50000790", "50003856", "50000783", "50000666", "50000791", "50000665", "50000541", "50003991", "50003990", "50000548", "50000550", "50000545", "50003993", "50003992", "50000549", "50006868", "50000543", "50000546", "50000542", "50000544", "50000547", "50000429", "50000431", "50000427", "50006173", "50000428", "50000430", "50000540", "50000539", "50003988", "50003989", "50000554", "50000555", "50000558", "50000559", "50000557", "50000556", "50000432", "50000433", "50004140", "50004142", "50000435", "50006188", "50004141", "50004143", "50005245", "50004147", "50004146", "50004144", "50004145", "50000426", "50000425", "50000671", "50000672", "50000669", "50000670", "50000667", "50000668", "50003997", "50003994", "50004001", "50003999", "50004005", "50004002", "50003995", "50004000", "50004006", "50003998", "50004004", "50004003", "50003996", "50000644", "50000639", "50000641", "50003976", "50003977", "50000642", "50000640", "50000643", "50000645", "50003824", "50003835", "50003818", "50003820", "50003836", "50003822", "50003825", "50003821", "50003826", "50003838", "50000779", "50003819", "50003833", "50003839", "50003831", "50003827", "50003840", "50003828", "50003829", "50003830", "50003817", "50003832", "50003837", "50003823", "50003834", "50004190", "50004191", "50003844", "50003852", "50003853", "50003847", "50003842", "50003841", "50003846", "50003843", "50003848", "50003849", "50003845", "50000780", "50003850", "50003851", "50000781", "50000659", "50000653", "50000654", "50000656", "50000657", "50000658", "50000655", "50000660", "50005464", "50003978", "50003980", "50003981", "50003979", "50000552", "50000553", "50000551", "50000436", "50004158", "50004168", "50004149", "50004181", "50004159", "50004169", "50004150", "50004182", "50006873", "50004192", "50004176", "50004157", "50004167", "50004148", "50004180", "50006871", "50006870", "50004187", "50004189", "50004160", "50004170", "50004151", "50006176", "50006171", "50006172", "50006169", "50006170", "50004162", "50004172", "50004153", "50004184", "50004178", "50004161", "50004171", "50004152", "50004183", "50004193", "50004177", "50004165", "50004175", "50004156", "50004188", "50004163", "50004173", "50004154", "50004185", "50006872", "50004164", "50004174", "50004155", "50006175", "50004186", "50004194", "50004179", "50004166", "50000664", "50000661", "50000663", "50000538", "50000662", "50003982", "50003985", "50003984", "50003987", "50003986", "50003983", "50000537", "50000434", "50000566", "50000565", "50000571", "50000567", "50000572", "50000564", "50000568", "50000573", "50000570", "50004139", "50004138", "50004137", "50000574", "50000569", "50004021", "50004023", "50004027", "50004019", "50004017", "50004026", "50004025", "50004020", "50004018", "50004029", "50004022", "50004028", "50004024", "50004014", "50004015", "50004012", "50004016", "50004010", "50004011", "50004013", "50004008", "50004009", "50000562", "50000560", "50000563", "50000561", "50001733", "50002771", "50002775", "50002772", "50002774", "50002770", "50001736", "50002773", "50001734", "50001735", "50002013", "50002009", "50002010", "50002011", "50002012", "50001511", "50001498", "50001499", "50001622", "50001500", "50001501", "50001503", "50001502", "50001504", "50001505", "50001506", "50001507", "50001508", "50001509", "50001510", "50001623", "50001624", "50001497", "50003152", "50003151", "50003150", "50003149", "50003153", "50000151", "50001587", "50001580", "50001586", "50001581", "50002923", "50002924", "50001582", "50001583", "50001584", "50001585", "50001598", "50001597", "50001595", "50002961", "50002962", "50002963", "50002965", "50002964", "50002966", "50001596", "50001599", "50002975", "50002972", "50002976", "50002974", "50002978", "50002970", "50002977", "50002979", "50002973", "50002971", "50002969", "50002968", "50000153", "50001593", "50001590", "50001592", "50006203", "50001591", "50001979", "50001980", "50001983", "50001984", "50001981", "50002373", "50002369", "50003156", "50002370", "50002368", "50001978", "50002371", "50002372", "50002374", "50001985", "50001982", "50002357", "50002344", "50002352", "50002362", "50006368", "50002345", "50002358", "50002364", "50002360", "50002350", "50002356", "50002351", "50002347", "50002361", "50002354", "50002365", "50002346", "50002343", "50002366", "50002359", "50002349", "50002355", "50002367", "50002353", "50002348", "50002363", "50001515", "50001516", "50001517", "50001513", "50001514", "50001512", "50001518", "50002119", "50005246", "50002138", "50002125", "50002127", "50002118", "50002126", "50002140", "50002128", "50002129", "50002139", "50002144", "50002137", "50002136", "50002120", "50002121", "50002122", "50002123", "50002124", "50002133", "50002142", "50002143", "50002141", "50002134", "50002116", "50002135", "50002114", "50002115", "50002117", "50002130", "50002131", "50002132", "50002014", "50001973", "50001963", "50002145", "50002328", "50002327", "50002319", "50002146", "50003154", "50002326", "50002318", "50002147", "50002317", "50001970", "50002323", "50002324", "50003155", "50002325", "50001987", "50001864", "50001992", "50001994", "50001986", "50001988", "50001993", "50001989", "50001850", "50001997", "50001996", "50001991", "50001995", "50002379", "50002375", "50002377", "50002378", "50002376", "50001990", "50002744", "50002766", "50002747", "50002733", "50002734", "50002761", "50002762", "50002765", "50002752", "50002758", "50002759", "50002731", "50002732", "50002763", "50002748", "50002760", "50002750", "50002749", "50002753", "50002751", "50002742", "50002743", "50002741", "50002740", "50002764", "50002736", "50002745", "50002735", "50002737", "50002746", "50002757", "50002739", "50002738", "50002756", "50002754", "50002755", "50001727", "50001618", "50001600", "50001601", "50001619", "50001616", "50001602", "50003097", "50003100", "50003098", "50003099", "50003093", "50003096", "50001617", "50003088", "50003090", "50003094", "50003089", "50003092", "50003095", "50003091", "50001615", "50002574", "50001855", "50002576", "50002575", "50001711", "50001726", "50002566", "50002572", "50006372", "50002567", "50001854", "50002573", "50002568", "50006373", "50002571", "50002569", "50002570", "50001719", "50001861", "50001863", "50001862", "50001723", "50001713", "50001718", "50001721", "50001724", "50001720", "50001707", "50002562", "50002563", "50002561", "50002558", "50002565", "50002560", "50002559", "50001857", "50001712", "50001714", "50006202", "50001856", "50002578", "50002579", "50002577", "50001710", "50001722", "50001860", "50001706", "50001708", "50001709", "50001725", "50001715", "50001858", "50002582", "50002723", "50002590", "50002589", "50002588", "50002584", "50002583", "50002591", "50002585", "50001859", "50002581", "50002580", "50002724", "50003208", "50002586", "50002587", "50002592", "50002725", "50002726", "50002727", "50002730", "50001716", "50002728", "50002729", "50001717", "50002927", "50002932", "50002928", "50002933", "50002934", "50002929", "50002930", "50002931", "50006840", "50005481", "50005482", "50002951", "50002943", "50002944", "50002941", "50002952", "50002953", "50002945", "50002954", "50002946", "50002948", "50002937", "50002940", "50002938", "50001594", "50002935", "50002936", "50002949", "50002955", "50002956", "50002942", "50002947", "50002957", "50002939", "50000266", "50000267", "50002094", "50004638", "50004605", "50004629", "50002113", "50004637", "50002097", "50004639", "50002081", "50002107", "50004614", "50004610", "50004611", "50004609", "50004612", "50004621", "50002108", "50002098", "50004606", "50004630", "50004627", "50004632", "50002075", "50004636", "50004635", "50002087", "50002088", "50002089", "50000268", "50002074", "50002078", "50002079", "50002109", "50002099", "50002103", "50002096", "50004607", "50002093", "50004634", "50004623", "50004620", "50004631", "50002085", "50002083", "50004633", "50002086", "50002084", "50002082", "50002095", "50004616", "50004618", "50004617", "50002112", "50002100", "50002102", "50002111", "50002110", "50004608", "50004613", "50004622", "50002104", "50002105", "50004604", "50002080", "50002077", "50004615", "50002106", "50004640", "50002073", "50002076", "50004626", "50002101", "50002091", "50002090", "50002092", "50004624", "50004625", "50004619", "50004628", "50000265", "50000152", "50002925", "50002926", "50001589", "50001588", "50001732", "50002767", "50002769"
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>FIVE</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "5"
			Case "50002768", "50001730", "50001728", "50001729", "50001731", "50004590", "50004594", "50004593", "50004603", "50004591", "50006839", "50004595", "50004599", "50004596", "50004602", "50004600", "50004601", "50004592", "50004598", "50004597", "50001737", "50001739", "50001738", "50003106", "50003105", "50001620", "50003109", "50003102", "50003101", "50003107", "50003103", "50003108", "50003104", "50003128", "50003115", "50003121", "50003136", "50001621", "50003133", "50003127", "50003141", "50003148", "50003114", "50003137", "50003122", "50003138", "50003147", "50003142", "50003123", "50003209", "50003139", "50003124", "50003126", "50003125", "50003146", "50003140", "50003143", "50003144", "50003145", "50003110", "50003134", "50003116", "50003117", "50003135", "50003118", "50003111", "50003129", "50003130", "50003119", "50003131", "50003112", "50003113", "50003132", "50003120", "50001579", "50001571", "50001573", "50002922", "50002777", "50002780", "50002778", "50002781", "50002920", "50002921", "50002776", "50001572", "50001578", "50001577", "50001576", "50001574", "50001575"
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>FIVE</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "5"
			Case Else
				buf = buf & "					<shop:ExpectedDeliveryPeriodType>TEN</shop:ExpectedDeliveryPeriodType>"	'발송 예정일 타입 코드'
				'mayDeliverDay = "10"
		End Select
'		buf = buf & "					<shop:ExpectedDeliveryPeriodDirectInput>"&mayDeliverDay&"</shop:ExpectedDeliveryPeriodDirectInput>"	'발송 예정'일 직접 입력값
		buf = buf & "					<shop:CustomProductAfterOrderYn>Y</shop:CustomProductAfterOrderYn>"	'주문 확인 후 제작 상품 여부 “Y” 또는 “N” ProductType의 CustomMade 대신 사용되는 필드
	Else
		buf = buf & "					<shop:CustomProductAfterOrderYn>N</shop:CustomProductAfterOrderYn>"	'주문 확인 후 제작 상품 여부 “Y” 또는 “N” ProductType의 CustomMade 대신 사용되는 필드
	End If
		buf = buf & "				</shop:Delivery>"
		getDeliveryType = buf
	End Function

	Public Function getSellerDiscount(isaleprice)
		Dim buf, sqlStr, ispecialPrice, istartDate, iendDate, iAmount
		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 mustPrice as specialPrice, startDate, endDate "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_outmall_mustPriceItem] "
		sqlStr = sqlStr & " WHERE mallgubun = '"& CMALLNAME &"' "
		sqlStr = sqlStr & " and itemid = '"& Fitemid &"' "
		'sqlStr = sqlStr & " and getdate() >= startDate and getdate() <= endDate "
		sqlStr = sqlStr & " ORDER BY startDate DESC "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			ispecialPrice	= rsget("specialPrice")
			istartDate		= rsget("startDate")
			iendDate		= rsget("endDate")
		End If
		rsget.Close
		'iAmount = MustPrice - ispecialPrice
		iAmount = isaleprice - ispecialPrice	'2019-07-05 현주님 요청

		buf = ""
		buf = buf & "				<shop:SellerDiscount>"														'판매자 즉시 할인 | 선택이나, 입력할 경우 하단 #은 필수
		buf = buf & "					<shop:Amount>"&iAmount&"</shop:Amount>"									'#PC 즉시 할인액/할인율 | PC할인만 적용하려면 MobileAmount에는 0을 입력..끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
		buf = buf & "					<shop:StartDate>"&LEFT(istartDate, 10) & " " & Num2Str(hour(istartDate),2,"0","R") & ":" & Num2Str(minute(istartDate),2,"0","R")&"</shop:StartDate>"				'PC 즉시 할인 시작일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00, 10, 20, 30, 40, 50분으로만 설정 가능
		buf = buf & "					<shop:EndDate>"&LEFT(iendDate, 10)& " " & Num2Str(hour(iendDate),2,"0","R") & ":" & Num2Str(minute(iendDate),2,"0","R")&"</shop:EndDate>"					'PC 즉시 할인 종료일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 23시 59분을 붙여서 저장됨..매시각 09, 19, 29, 39, 49, 59분으로만 설정 가능
		buf = buf & "					<shop:MobileAmount>"&iAmount&"</shop:MobileAmount>"						'#모바일 즉시 할인액/할인율 | 모바일 할인만 적용하려면 Amount에 0을 입력..끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
		buf = buf & "					<shop:MobileStartDate>"&LEFT(istartDate, 10) & " " & Num2Str(hour(istartDate),2,"0","R") & ":" & Num2Str(minute(istartDate),2,"0","R")&"</shop:MobileStartDate>"	'모바일 즉시 할인 시작일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00, 10, 20, 30, 40, 50분으로만 설정 가능
		buf = buf & "					<shop:MobileEndDate>"&LEFT(iendDate, 10)& " " & Num2Str(hour(iendDate),2,"0","R") & ":" & Num2Str(minute(iendDate),2,"0","R")&"</shop:MobileEndDate>"		'모바일 즉시 할인 종료일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 23시 59분을 붙여서 저장됨..매시각 09, 19, 29, 39, 49, 59분으로만 설정 가능
		buf = buf & "				</shop:SellerDiscount>"
		getSellerDiscount = buf
	End Function

	Public Function getNvstoregiftItemContParamToReg()
		Dim strRst, strSQL, strtextVal
		strRst = ("<div align=""center"">")
		'로고 이미지
		strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_logo_nvstoregift.jpg""></center></p><br>")
		'기본 이미지
		strRst = strRst & ("<p><center><img src="""& FbasicImage &"""></center></p><br>")

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " where mallid in ('','"&CMALLNAME&"') and linkgbn = 'topContents' and itemid = '"&Fitemid&"' " & VBCRLF
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		if Not(rsget.EOF or rsget.BOF) then
			strRst = strRst & rsget("textVal") & "<br>"
		End If
		rsget.Close

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

		strRst = strRst & ("<p><center><a href=""http://storefarm.naver.com/tenbytengift"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_nvstoregift.jpg""></a></center></p><br>")
		'#배송 주의사항
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_nvstoregift.jpg"">")
		strRst = strRst & ("</div>")
		getNvstoregiftItemContParamToReg = strRst

		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = ""
			strRst = strRst & ("<div align=""center"">")
			'로고 이미지
			strRst = strRst & ("<p><center><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_logo_nvstoregift.jpg""></center></p><br>")
			'기본 이미지
			strRst = strRst & ("<p><center><img src="""& FbasicImage &"""></center></p><br>")

			If ForderComment <> "" Then
				strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
			End If
			strRst = strRst & strtextVal & "<br>"
			strRst = strRst & ("<p><center><a href=""http://storefarm.naver.com/tenbytengift"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_nvstoregift.jpg""></a></center></p><br>")
			strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_nvstoregift.jpg"">")
			strRst = strRst & ("</div>")
			getNvstoregiftItemContParamToReg = strRst
		End If
		rsget.Close
	End Function

	Public Function getSellerComment
		Dim buf, icomment
		icomment = Fordercomment
		icomment = replace(icomment,"\","")
		icomment = replace(icomment,"*","")
		icomment = replace(icomment,"?","")
		icomment = replace(icomment,"""","")
		icomment = replace(icomment,"<","")
		icomment = replace(icomment,">","")
		buf = ""

		If len(icomment) > 1300 Then
			icomment = DDotFormat(icomment,1290)
		End If

		If len(icomment) = 2 AND instr(icomment, chr(13)) Then
			icomment = ""
		End If

		If IsNULL(icomment) OR Trim(icomment) = "" Then
			buf = buf & "				<shop:SellerCommentUsable>N</shop:SellerCommentUsable>"			'판매자 특이사항 사용 여부 | Y or N..Y입력시 SellerCommentContent 필수, N 입력시 특이 사항 없음으로 저장되며 SellerCommentContent 필드 무시..상품 수정시 SellerCommentUsable 요소를 삭제하고 전송하면 기존에 저장된 값이 변경되지 않는다.
'			buf = buf & "				<shop:SellerCommentContent></shop:SellerCommentContent>"		'판매자 특이사항 직접 입력 값 | SellerCommentUsable이 Y일 때 저장
		Else
			buf = buf & "				<shop:SellerCommentUsable>Y</shop:SellerCommentUsable>"			'판매자 특이사항 사용 여부 | Y or N..Y입력시 SellerCommentContent 필수, N 입력시 특이 사항 없음으로 저장되며 SellerCommentContent 필드 무시..상품 수정시 SellerCommentUsable 요소를 삭제하고 전송하면 기존에 저장된 값이 변경되지 않는다.
			buf = buf & "				<shop:SellerCommentContent><![CDATA["&icomment&"]]></shop:SellerCommentContent>"		'판매자 특이사항 직접 입력 값 | SellerCommentUsable이 Y일 때 저장
		End If
'		buf = buf & "				<shop:SellerCustomCode1></shop:SellerCustomCode1>"				'판매자가 내부에서 사용하는 코드
'		buf = buf & "				<shop:SellerCustomCode2></shop:SellerCustomCode2>"				'판매자가 내부에서 사용하는 코드
		getSellerComment = buf
	End Function

	Public Function getNvstoregiftItemInfoCdToReg
		Dim buf, strSQL, mallinfoCd, infoContent, mallinfodiv, mallinfoName
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00001') THEN '상세페이지 참조' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'N') THEN 'N' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'Y') THEN 'Y' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) < 2 THEN i.itemname "
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) >= 2 THEN F.infocontent "
		strSQL = strSQL & " 	 WHEN c.partnerShipInfoType='P' THEN '텐바이텐 1644-6035' " & vbcrlf
		strSQL = strSQL & " 	 WHEN LEN(isnull(F.infocontent, '')) < 2 THEN '상세페이지 참조' " & vbcrlf
		strSQL = strSQL & " ELSE isnull(F.infocontent, '') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " WHERE M.mallid = 'nvstorefarm' and IC.itemid='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " ORDER BY infocd ASC " & vbcrlf
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			mallinfodiv = rsget("mallinfodiv")
			Select Case mallinfodiv
				Case "01"		mallinfoName = "Wear"
				Case "02"		mallinfoName = "Shoes"
				Case "03"		mallinfoName = "Bag"
				Case "04"		mallinfoName = "FashionItems"
				Case "05"		mallinfoName = "SleepingGear"
				Case "06"		mallinfoName = "Furniture"
				Case "07"		mallinfoName = "ImageAppliances"
				Case "08"		mallinfoName = "HomeAppliances"
				Case "09"		mallinfoName = "SeasonAppliances"
				Case "10"		mallinfoName = "OfficeAppliances"
				Case "11"		mallinfoName = "OpticsAppliances"
				Case "12"		mallinfoName = "MicroElectronics"
				Case "13"		mallinfoName = "Cellphone"
				Case "14"		mallinfoName = "Navigation"
				Case "15"		mallinfoName = "CarArticles"
				Case "16"		mallinfoName = "MedicalAppliances"
				Case "17"		mallinfoName = "KitchenUtensils"
				Case "18"		mallinfoName = "Cosmetic"
				Case "19"		mallinfoName = "Jewellery"
				Case "20"		mallinfoName = "Food"
				Case "21"		mallinfoName = "GeneralFood"
				Case "22"		mallinfoName = "DietFood"
				Case "23"		mallinfoName = "Kids"
				Case "24"		mallinfoName = "MusicalInstrument"
				Case "25"		mallinfoName = "SportsEquipment"
				Case "26"		mallinfoName = "Books"
				Case "27"		mallinfoName = "LodgmentReservation"
				Case "28"		mallinfoName = "TravelPackage"
				Case "30"		mallinfoName = "RentCar"
				Case "31"		mallinfoName = "RentalHa"
				Case "32"		mallinfoName = "RentalEtc"
				Case "33"		mallinfoName = "DigitalContents"
				Case "35"		mallinfoName = "Etc"
				Case "47"		mallinfoName = "Biochemistry"
				Case "48"		mallinfoName = "Biocidal"
			End Select

			buf = ""
			buf = buf & "				<shop:"&mallinfoName&">"
			buf = buf & "					<shop:NoRefundReason><![CDATA[상세페이지 참조]]></shop:NoRefundReason>"
			buf = buf & "					<shop:ReturnCostReason><![CDATA[상세페이지 참조]]></shop:ReturnCostReason>"
			buf = buf & "					<shop:QualityAssuranceStandard><![CDATA[상세페이지 참조]]></shop:QualityAssuranceStandard>"
			buf = buf & "					<shop:CompensationProcedure><![CDATA[상세페이지 참조]]></shop:CompensationProcedure>"
			buf = buf & "					<shop:TroubleShootingContents><![CDATA[상세페이지 참조]]></shop:TroubleShootingContents>"
			Do until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
'			    If mallinfoCd = "Size" Then
				If infoContent <> "" Then
			    	infoContent = replace(infoContent, "*", "x")
			    End If
'			    End If
				buf = buf & "					<shop:"&mallinfoCd&"><![CDATA["&infoContent&"]]></shop:"&mallinfoCd&">"
				rsget.MoveNext
			Loop
			buf = buf & "				</shop:"&mallinfoName&">"
		End If
		rsget.Close
		getNvstoregiftItemInfoCdToReg = buf
	End Function

	'// 업로드 이미지 XML 생성
	Public Function getNvstoregiftImageRegXML(oServ, oOper)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature, strSQL, i, shoppingWindowImgCnt, arrRows
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1o1934_01"
		End If
		Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, oOper)

		strRst = ""
		strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
		strRst = strRst & "	<soapenv:Header/>"
		strRst = strRst & "	<soapenv:Body>"
		strRst = strRst & "		<shop:UploadImageRequest>"
		strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
		strRst = strRst & "			<shop:AccessCredentials>"
		strRst = strRst & "				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
		strRst = strRst & "				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
		strRst = strRst & "				<shop:Signature>"&osignature&"</shop:Signature>"
		strRst = strRst & "			</shop:AccessCredentials>"
		strRst = strRst & "			<shop:Version>2.0</shop:Version>"
		strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
		strRst = strRst & "			<ImageURLList>"
		If (application("Svr_Info") = "Dev") Then
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/basic/146/B001469141.jpg</shop:URL>"
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/add1/146/A001469141_01.jpg</shop:URL>"
			strRst = strRst & "				<shop:URL>http://webimage.10x10.co.kr/image/add2/146/A001469141_02.jpg</shop:URL>"
		Else

			strSQL = ""
			strSQL = strSQL & " SELECT COUNT(*) as cnt "
			strSQL = strSQL & " FROM db_etcmall.dbo.tbl_nvstorefarm_uploadimage "
			strSQL = strSQL & " WHERE ITEMID = '"& Fitemid &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
				shoppingWindowImgCnt = rsget("cnt")
			rsget.Close

			If shoppingWindowImgCnt = 0 Then
				strRst = strRst & "				<shop:URL>"&FbasicImage&"</shop:URL>"
				strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
				rsget.CursorLocation = adUseClient
				rsget.CursorType=adOpenStatic
				rsget.Locktype=adLockReadOnly
				rsget.Open strSQL, dbget
				If Not(rsget.EOF or rsget.BOF) Then
					For i=1 to rsget.RecordCount
						If rsget("imgType") = "0" Then
							strRst = strRst & "				<shop:URL>"&"http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")&"</shop:URL>"
						End If
						rsget.MoveNext
						If i >= 4 Then Exit For
					Next
				End If
				rsget.Close
			Else
				strSQL = ""
				strSQL = strSQL & " SELECT IMAGENAME, GUBUN "
				strSQL = strSQL & " FROM db_etcmall.dbo.tbl_nvstorefarm_uploadimage "
				strSQL = strSQL & " WHERE ITEMID = '"& Fitemid &"' "
				strSQL = strSQL & " ORDER BY GUBUN ASC "
				rsget.CursorLocation = adUseClient
				rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.Eof Then
					arrRows = rsget.getRows()
				End If
				rsget.Close

				If isArray(arrRows) then
					For i = 0 To UBound(arrRows, 2)
						strRst = strRst & "				<shop:URL>"& webImgUrl & "/image/nvadd" & CStr(arrRows(1,i)) & "/" & GetImageSubFolderByItemid(Fitemid) + "/"  & arrRows(0,i) &"</shop:URL>"
					Next
				End If
			End If
		End If
		strRst = strRst & "			</ImageURLList>"
		strRst = strRst & "		</shop:UploadImageRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		getNvstoregiftImageRegXML = strRst
	End Function

	'// 상품등록 XML 생성
	Public Function getNvstoregiftItemRegXML(oServ, oOper, isEdit)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature, isDiscount, saleprice
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1o1934_01"
		End If
		Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, oOper)

		isDiscount = fnIsSpecialDate
		saleprice = getSalePrice

		strRst = ""
		strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:shop=""http://shopn.platform.nhncorp.com/"">"
		strRst = strRst & "	<soapenv:Header/>"
   		strRst = strRst & "	<soapenv:Body>"
		strRst = strRst & "		<shop:ManageProductRequest>"
		strRst = strRst & "			<shop:RequestID>"&reqID&"</shop:RequestID>"
		strRst = strRst & "			<shop:AccessCredentials>"
		strRst = strRst & "				<shop:AccessLicense>"&oaccessLicense&"</shop:AccessLicense>"
		strRst = strRst & "				<shop:Timestamp>"&oTimestamp&"</shop:Timestamp>"
		strRst = strRst & "				<shop:Signature>"&osignature&"</shop:Signature>"
		strRst = strRst & "			</shop:AccessCredentials>"
		strRst = strRst & "			<shop:Version>2.0</shop:Version>"
		strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"							''??
		strRst = strRst & "			<Product>"
		If isEdit = "Y" Then
			strRst = strRst & "			<shop:ProductId>"&FNvstoregiftGoodNo&"</shop:ProductId>"		'상품ID | 없으면 등록, 있으면 수정
		End If
		strRst = strRst & "				<shop:StatusType>SALE</shop:StatusType>"			'# 상품판매상태 | 등록은 SALE(판매중)만 입력, 수정시 SALE, SUSP(판매 중지)만 입력, StockQuantity가 0 이면 OSTK(품절)로 저장됨
		strRst = strRst & "				<shop:SaleType>NEW</shop:SaleType>"					'상품 판매 유형..미입력시 NEW로 저장
		strRst = strRst & getzCostomMadeInd													'#주문 제작 상품 여부
		strRst = strRst & "				<shop:CategoryId>"&FCateKey&"</shop:CategoryId>"	'#Leaf 카테고리 | ID 상품등록시 필수 | ModelType의 모델명ID가 입력된 경우 해당 모델명 ID에 매핑된  Leaf 카테고리 ID로 저장하며 요청으로 전달된 CategoryId는 무시된다
'		strRst = strRst & "				<shop:LayoutType></shop:LayoutType>"				'상품 상세 레이아웃 타입 코드 | 관련 코드 상품 상세 레이아웃 타입 : 코드 미입력 시 베이직형 (BASIC)으로 저장된다
		strRst = strRst & "				<shop:Name><![CDATA["&getItemNameFormat&"]]></shop:Name>"			'#상품명
'		strRst = strRst & "				<shop:PublicityPhraseContent></shop:PublicityPhraseContent>"		'홍보 문구
'		strRst = strRst & "				<shop:PublicityPhraseStartDate></shop:PublicityPhraseStartDate>"	'홍보 문구 전시 시작일
'		strRst = strRst & "				<shop:PublicityPhraseEndDate></shop:PublicityPhraseEndDate>"		'홍보 문구 전시 종료일
		strRst = strRst & "				<shop:SellerManagementCode>"&FItemid&"</shop:SellerManagementCode>"	'판매자 상품 코드
'		strRst = strRst & "				<shop:SellerBarCode></shop:SellerBarCode>"							'판매자 바코드
		strRst = strRst & "				<shop:Model>"	'모델 정보| 모델 ID 정보가 없는 경우 브랜드명, 제조사명만 수정 가능..인증유형이 "방송통신기자재 적합인증/적합등록/잠정인증 어린이제품 안전인증/안전확인/공급자적합성확인 인 경우 필수, 제조사명(ManufacturerName), 브랜드명(BrandName),모델명(ModelName)이 필수로 입력
'		strRst = strRst & "					<shop:Id></shop:Id>"									'모델 ID
		strRst = strRst & "					<shop:ManufacturerName><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></shop:ManufacturerName>"		'제조사명
		strRst = strRst & "					<shop:BrandName><![CDATA["&chkIIF(trim(FSocname_kor)="" or isNull(FSocname_kor),"상품설명 참조",FSocname_kor)&"]]></shop:BrandName>"				'브랜드명
		If FNeedCert = "Y" Then
			strRst = strRst & getModelName
		End If
		strRst = strRst & "				</shop:Model>"
'		strRst = strRst & "				<shop:AttributeValueList></shop:AttributeValueList>"		' ,로 분리된 속성의 목록 | 현재는 사용하지 않으며 향후 사용 예정
		If FNeedCert = "Y" Then
			strRst = strRst & getNvstoregiftCertInfo()
		End If
		If FNeedCert = "Y" Then
			If getNvstoregiftCertInfo = "" Then
				strRst = strRst & "			<shop:KCCertifiedProductExclusion>Y</shop:KCCertifiedProductExclusion>"	'KC인증대상제외타입 | Y : KC인증대상아님, N: KC인증대상, KC_EXEMPTION : KC면제대상
			Else
				If FCateKey = "50004234" Then
					strRst = strRst & "			<shop:KCCertifiedProductExclusion>Y</shop:KCCertifiedProductExclusion>"	'KC인증대상제외타입 | Y : KC인증대상아님, N: KC인증대상, KC_EXEMPTION : KC면제대상
				Else
					strRst = strRst & "			<shop:KCCertifiedProductExclusion>"&Chkiif(Fsafetyyn="Y", "N", "Y")&"</shop:KCCertifiedProductExclusion>"	'KC인증대상제외타입 | Y : KC인증대상아님, N: KC인증대상, KC_EXEMPTION : KC면제대상
				End If
			End If
		End If
		strRst = strRst & getOriginAreaType															'#원산지 정보
'		strRst = strRst & "				<shop:ManufactureDate></shop:ManufactureDate>"				'제조 일자 | YYYY-MM-DD 형식
'		strRst = strRst & "				<shop:ValidDate></shop:ValidDate>"							'유효 일자 | YYYY-MM-DD 형식
		strRst = strRst & "				<shop:TaxType>"&CHKIIF(FVatInclude="N","DUTYFREE","TAX")&"</shop:TaxType>"	'#부가세 | 과세 : TAX, 면세 : DUTYFREE, 영세 : SMALL
		strRst = strRst & "				<shop:MinorPurchasable>"&Chkiif(IsAdultItem() = "Y", "N", "Y")&"</shop:MinorPurchasable>"	'#미성년자 구매 가능 여부 Y or N
		strRst = strRst & getImageType																'#이미지 정보
		strRst = strRst & "				<shop:DetailContent><![CDATA["&getNvstoregiftItemContParamToReg&"]]></shop:DetailContent>"		'#상품 상세 정보
'		strRst = strRst & "				<shop:SellerNoticeId></shop:SellerNoticeId>"										'공지사항 번호
		strRst = strRst & "				<shop:AfterServiceTelephoneNumber><![CDATA[1644-6035]]></shop:AfterServiceTelephoneNumber>"		'#A/S 전화번호
		strRst = strRst & "				<shop:AfterServiceGuideContent><![CDATA[A/S 관련은 텐바이텐 고객행복센터를 통해 문의해 주시기 바랍니다.]]></shop:AfterServiceGuideContent>"	'#A/S 안내
'		strRst = strRst & "				<shop:PurchaseReviewExposure></shop:PurchaseReviewExposure>"						'구매평 노출 여부 | Y or N, 구매평 노출 설정 가능 카테고리일 경우에만 유효하며 그 외에는 Y로 설정된다. 미입력 시 Y로 저장됨
		If (FItemID = "1488156") Then
			strRst = strRst & "				<shop:RegularCustomerExclusiveProduct>Y</shop:RegularCustomerExclusiveProduct>"		'단골 회원 전용 상품 여부 | Y or N 미입력시 N으로 저장됨
		Else
			strRst = strRst & "				<shop:RegularCustomerExclusiveProduct>N</shop:RegularCustomerExclusiveProduct>"		'단골 회원 전용 상품 여부 | Y or N 미입력시 N으로 저장됨
		End If

		If (FItemID = "2362615") OR (FItemID = "2357727") Then
			strRst = strRst & "				<shop:KnowledgeShoppingProductRegistration>N</shop:KnowledgeShoppingProductRegistration>"	'네이버 쇼핑 등록 | Y or N 네이버 광고주가 아닌 경우 N으로 저장됨
		Else
			strRst = strRst & "				<shop:KnowledgeShoppingProductRegistration>Y</shop:KnowledgeShoppingProductRegistration>"	'네이버 쇼핑 등록 | Y or N 네이버 광고주가 아닌 경우 N으로 저장됨
		End If
'		strRst = strRst & "				<shop:GalleryId></shop:GalleryId>"							'갤러리 번호
'		strRst = strRst & "				<shop:SaleStartDate></shop:SaleStartDate>"					'판매 시작일 | YYYY-MM-DD 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00분으로만 설정 가능
'		strRst = strRst & "				<shop:SaleEndDate></shop:SaleEndDate>"						'판매 종료일 | YYYY-MM-DD HH:mm형식..날짜까지만 입력하는 경우 자동으로 23시 59분을 붙여서 저장됨.매시각 59분으로만 설정 가능
'		strRst = strRst & "				<shop:SalePrice>"&Clng(GetRaiseValue(MustPrice/10)*10)&"</shop:SalePrice>"		'#판매가
		strRst = strRst & "				<shop:SalePrice>"& saleprice &"</shop:SalePrice>"		'#판매가 / 2019-07-05 17:17 변경
		If (isEdit = "Y")  Then
			If (Foptioncnt = 0) Then
				strRst = strRst & "				<shop:StockQuantity>"&getLimitNvstoregiftEa&"</shop:StockQuantity>"		'#재고 수량 | 상품등록시 필수, 상품 수정시 재고 수량을 입력하지 않으면 스토어팜 DB에 저장된 현재 재고값이 변하지 않는다. 수정시 재고 수량 0으로 입력되면 StatusType으로 전달된 항목은 무시되며 상품 상태는 OSTK(품절)로 저장됨
			End If
		Else
			strRst = strRst & "				<shop:StockQuantity>"&getLimitNvstoregiftEa&"</shop:StockQuantity>"		'#재고 수량 | 상품등록시 필수, 상품 수정시 재고 수량을 입력하지 않으면 스토어팜 DB에 저장된 현재 재고값이 변하지 않는다. 수정시 재고 수량 0으로 입력되면 StatusType으로 전달된 항목은 무시되며 상품 상태는 OSTK(품절)로 저장됨
		End If
'		strRst = strRst & "				<shop:MinPurchaseQuantity></shop:MinPurchaseQuantity>"					'최소 구매 수량
		If FItemid = "1488156" Then
			strRst = strRst & "				<shop:MaxPurchaseQuantityPerId>1</shop:MaxPurchaseQuantityPerId>"	'1인 최대 구매 수량
		End If
		strRst = strRst & "				<shop:MaxPurchaseQuantityPerOrder>"&getOrderMaxNum&"</shop:MaxPurchaseQuantityPerOrder>"	'1회 최대 구매 수량
		strRst = strRst & getDeliveryType														'배송 정보 | 미입력시 배송 없는 상품으로 등록됨
		If isDiscount <> "NN" Then
			strRst = strRst & getSellerDiscount(saleprice)
		End If
'		strRst = strRst & "				<shop:MultiPurchaseDiscount>"							'복수 구매 할인 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#복수 구매 할인액/할인율 | 끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
'		strRst = strRst & "					<shop:OrderAmount></shop:OrderAmount>"				'#복수 구매 할인 조건 금액/개수 | 끝문자(개, 숫자)에 따라 단위 구분..ex)값이 10개이면 개수, 1000이면 금액을 나타낸다
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'복수 구매 할인 시작일 | YYYY-MM-DD 형식
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'복수 구매 할인 종료일 | YYYY-MM-DD 형식..시작일을 입력한 경우 필수
'		strRst = strRst & "				</shop:MultiPurchaseDiscount>"
'		strRst = strRst & "				<shop:Mileage>"											'상품 구매시 적립되는 네이버페이 포인트 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#네이버페이 포인트 적립액/적립율 | 끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'네이버페이 포인트 유효 기간 시작일..YYYY-MM-DD 형식
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'네이버페이 포인트 유효 기간 종료일..YYYY-MM-DD 형식, 시작일을 입력한 경우 필수
'		strRst = strRst & "				</shop:Mileage>"
'		strRst = strRst & "				<shop:ReviewPoint>"												'구매평 작성 시 적립되는 네이버페이 포인트 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:PurchaseReviewPoint></shop:PurchaseReviewPoint>"		'구매평 작성 시 적립되는 네이버페이 포인트 | 구매평, 프리미엄 구매평 둘 중 하나만 필수 입력
'		strRst = strRst & "					<shop:PremiumReviewPoint></shop:PremiumReviewPoint>"		'프리미엄 구매평 작성 시 적립되는 네이버페이 포인트 | 구매평, 프리미엄 구매평 둘 중 하나만 필수 입력
'		strRst = strRst & "					<shop:RegularCustomerPoint></shop:RegularCustomerPoint>"	'단골 회원이 구매평이나 프리미엄 구매평 작성 시 추가 적립되는 네이버페이 포인트 | 구매평이나 프리미엄 구매평이 있는 경우에만 입력
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"							'네이버페이 포인트 유효 기간 시작일 | YYYY-MM-DD 형식
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"								'네이버페이 포인트 유효 기간 종료일 | YYYY-MM-DD 형식, 시작일을 입력한 경우 필수
'		strRst = strRst & "				</shop:ReviewPoint>"
'		strRst = strRst & "				<shop:FreeInterest>"								'무이자 할부 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:Month></shop:Month>"						'#무이자 할부 개월 수
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"				'무이자 할부 시작일 | YYYY-MM-DD 형식
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"					'무이자 할부 종료일 | YYYY-MM-DD 형식, 시작일을 입력한 경우 필수
'		strRst = strRst & "				</shop:FreeInterest>"
'		strRst = strRst & "				<shop:Gift>"										'사은품 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:Name></shop:Name>"						'#사은품
'		strRst = strRst & "				</shop:Gift>"
'		strRst = strRst & "				<shop:ECoupon>"										'ECOUPON | 이쿠폰 카테고리 상품인 경우 필수
'		strRst = strRst & "					<shop:PeriodType></shop:PeriodType>"			'#e쿠폰 유효기간 구분
'		strRst = strRst & "					<shop:ValidStartDate></shop:ValidStartDate>"	'e쿠폰 유효기간 시작일..YYYY-MM-DD형식, e쿠폰 유효기간 구분 타입(PeriodType)이 특정기간인 경우 필수
'		strRst = strRst & "					<shop:ValidEndDate></shop:ValidEndDate>"		'e쿠폰 유효기간 종료일..YYYY-MM-DD형식, e쿠폰 유효기간 구분 타입(PeriodType)이 특정기간인 경우 필수
'		strRst = strRst & "					<shop:PeriodDays></shop:PeriodDays>"			'e쿠폰 유효기간 내용..e쿠폰 유효기간 구분 타입(PeriodType)이 자동 기간인 경우 필수
'		strRst = strRst & "					<shop:PublicInformationContents></shop:PublicInformationContents>"		'e쿠폰 발행처
'		strRst = strRst & "					<shop:ContactInformationContents></shop:ContactInformationContents>"	'e쿠폰 연락처
'		strRst = strRst & "					<shop:UsePlaceType></shop:UsePlaceType>"			'e쿠폰 사용 장소 타입
'		strRst = strRst & "					<shop:UsePlaceContents></shop:UsePlaceContents>"	'e쿠폰 사용 장소
'		strRst = strRst & "					<shop:RestrictCart></shop:RestrictCart>"			'e쿠폰 장바구니 제한 | Y or N
'		strRst = strRst & "				</shop:ECoupon>"
'		strRst = strRst & "				<shop:PurchaseApplicationUrl></shop:PurchaseApplicationUrl>"	'휴대폰 구매신청서 URL | 휴대폰 카테고리 상품인 경우 필수
'		strRst = strRst & "				<shop:CellPhonePrice></shop:CellPhonePrice>"					'고객부담 휴대폰 단말기 대금 | 휴대폰 카테고리 상품인 경우 필수
'		strRst = strRst & "				<shop:WifiOnly></shop:WifiOnly>"		'Wifi 전용 상품 여부 | Y or N..태블릿 카테고리 상품인 경우 필수..Y 입력시 PurchaseApplicationUrl, CellPhonePrice 입력불가..N 입력시 PurchaseApplicationUrl, CellPhonePrice 입력 필수
		strRst = strRst & "				<shop:ProductSummary>"					'상품 요약 정보 | 상품 등록시 필수, 상품 수정 시에는 기존에 상품 요약 정보가 입력된 경우에만 생략할 수 있다. 이 경우 기존에 저장된 상품 요약 정보 값이 유지된다.
		strRst = strRst & getNvstoregiftItemInfoCdToReg
		strRst = strRst & "				</shop:ProductSummary>"
		strRst = strRst & getSellerComment
		If (Fitemdiv="06" or Fitemdiv="16") Then
			strRst = strRst & "				<shop:CustomProductYn>Y</shop:CustomProductYn>"	'맞춤 제작 상품 여부 “Y” 또는 “N” ProductType의 UseReturnCancelNotification 대신 사용되는 필드
		Else
			strRst = strRst & "				<shop:CustomProductYn>N</shop:CustomProductYn>"	'맞춤 제작 상품 여부 “Y” 또는 “N” ProductType의 UseReturnCancelNotification 대신 사용되는 필드
		End If
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</shop:ManageProductRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		getNvstoregiftItemRegXML = strRst
	End Function
End Class

Class CNvstoregift
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
	Public FRectGubun

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

	Public Sub getNvstoregiftNotRegOneItem
		Dim strSql, addSql, i

		If FRectItemID <> "" Then
			addSql = addSql & " and i.itemid in (" & FRectItemID & ")"
			'옵션 전체 품절인 경우 등록 불가.
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
            addSql = addSql & " WHERE optCnt-optNotSellCnt < 1 "
            addSql = addSql & " )"
            addSql = addSql & " and isNULL(c.infodiv,'') not in ('','20','22')"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.nvstoregiftStatCD,-9) as nvstoregiftStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName, isnull(ca.needCert, '') as needCert "
		strSql = strSql & "	, isnull(bm.CateKey, '') as CateKey, isnull(R.APIaddImg, '') as APIaddImg, p.purchasetype "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_nvstorefarm_cate_mapping] as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_nvstorefarm_category as ca on ca.catekey = bm.catekey "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_nvstoregift_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtusing = 'Y'"
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "							'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
		strSql = strSql & " and i.cate_large <> '' "
'		strSql = strSql & " and i.cate_large <> '999' "		'선물하기 관리카테고리는 999
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & " and not exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoregift') "
		strSql = strSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoregift') "
		strSql = strSql & " and not exists(select top 1 y.idx FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] y with (nolock) "
		strSql = strSql & " 				WHERE convert(varchar(6), y.cdl + y.cdm) = convert(varchar(6), (i.cate_large + i.cate_mid)) and y.mallgubun = 'nvstoregift') "
		If FRectGubun <> "IMG" Then
			strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.[dbo].[tbl_nvstoregift_regItem] where nvstoregiftStatCD > 3) "
		End If
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvstoregiftItem
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
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FNvstoregiftStatCD	= rsget("nvstoregiftStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FCateKey			= rsget("CateKey")
                FOneItem.FSocname_kor		= rsget("socname_kor")
                FOneItem.FAPIaddImg			= rsget("APIaddImg")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
                FOneItem.FRegImageName 		= rsget("regImageName")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FNeedCert 			= rsget("needCert")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FPurchasetype		= rsget("purchasetype")
		End If
		rsget.Close
	End Sub

	Public Sub getNvstoregiftEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(m.nvstoregiftGoodNo, '') as nvstoregiftGoodNo, m.nvstoregiftprice, m.nvstoregiftSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, isnull(bm.CateKey, '') as CateKey, isnull(m.APIaddImg, '') as APIaddImg "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, isnull(ca.needCert, '') as needCert, p.purchasetype "
    	strSql = strSql & "	,(CASE WHEN i.isusing = 'N' "
		strSql = strSql & "		or i.isExtUsing='N'"
'		strSql = strSql & "		or uc.isExtUsing='N'"		''2018-12-03 김진영 수정 // 제휴판매안함이라도 스토어팜 판매 가능이면 판매
'		strSql = strSql & "		or ((i.deliveryType = 9) and (i.sellcash < 1000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn <> 'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
'		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large=''" '선물하기 관리카테고리는 999이므로 제외처리
'		strSql = strSql & "		or ((i.sailyn = 'N') and ( ((i.sellcash-i.buycash)/i.sellcash)*100 < "&CMAXMARGIN&" )) "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		''strSql = strSql & "		or i.makerid  in (Select makerid From db_temp.dbo.tbl_EpShop_not_in_makerid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'N') "
		''strSql = strSql & "		or i.itemid  in (Select itemid From db_temp.dbo.tbl_EpShop_not_in_itemid Where mallgubun='"&CMALLGUBUN&"' and isusing = 'Y') "
		strSql = strSql & " 	or exists(SELECT top 1 n.makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid n with (nolock) WHERE n.makerid=i.makerid and n.mallgubun = 'nvstoregift') "
		strSql = strSql & " 	or exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstoregift') "
		strSql = strSql & " 	or exists(select top 1 y.idx FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] y with (nolock) "
		strSql = strSql & " 				WHERE convert(varchar(6), y.cdl + y.cdm) = convert(varchar(6), (i.cate_large + i.cate_mid)) and y.mallgubun = 'nvstoregift') "
		strSql = strSql & "		or ((i.LimitYn = 'Y') and (i.LimitNo - i.LimitSold <= "&CMAXLIMITSELL&")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_partner.dbo.tbl_partner as p with (nolock) on i.makerid = p.id"
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_nvstoregift_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_nvstorefarm_cate_mapping as bm on bm.tenCateLarge=i.cate_large and bm.tenCateMid=i.cate_mid and bm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_nvstorefarm_category as ca on ca.catekey = bm.catekey "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIaddImg = 'Y' "
		strSql = strSql & " and m.nvstoregiftStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.nvstoregiftGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvstoregiftItem
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
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.FNvstoregiftGoodNo		= rsget("nvstoregiftGoodNo")
				FOneItem.FNvstoregiftprice		= rsget("nvstoregiftprice")
				FOneItem.FNvstoregiftSellYn		= rsget("nvstoregiftSellYn")

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv       = rsget("infoDiv")
	            FOneItem.Fsafetyyn      = rsget("safetyyn")
	            FOneItem.FsafetyDiv     = rsget("safetyDiv")
	            FOneItem.FsafetyNum     = rsget("safetyNum")
	            FOneItem.FmaySoldOut    = rsget("maySoldOut")
	            FOneItem.Fregitemname    = rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
                FOneItem.FCateKey			= rsget("CateKey")
                FOneItem.FNeedCert			= rsget("needCert")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FOrderMaxNum 		= rsget("orderMaxNum")
				FOneItem.FPurchasetype		= rsget("purchasetype")
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

Function getNvstoregiftGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 nvstoregiftGoodNo FROM db_etcmall.[dbo].[tbl_nvstoregift_regItem] WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getNvstoregiftGoodNo = rsget("nvstoregiftGoodNo")
	rsget.Close
End Function

%>
