<%
CONST CMAXMARGIN = 15
CONST CMALLNAME = "lfmall"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 5									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST APIURL	= "https://b2b.lfmall.co.kr"
CONST AuthId	= "tenten"
CONST AuthKey	= "Ten1010*!!"
CONST CDEFALUT_STOCK = 1000

Class CLfmallItem
	Public FItemid
	Public FItemname
	Public FSmallImage
	Public FMakerid
	Public FRegdate
	Public FLastUpdate
	Public FOrgPrice
	Public FSellCash
	Public FBuyCash
	Public FSellYn
	Public FSaleYn
	Public FLimitYn
	Public FLimitNo
	Public FLimitSold
	Public FLfmallRegdate
	Public FLfmallLastUpdate
	Public FLfmallGoodNo
	Public FLfmallPrice
	Public FLfmallSellYn
	Public FRegUserid
	Public FLfmallStatCd
	Public Flfbrandcode
	Public FItemKindCode
	Public FSeasonCode
	Public FColor1Code
	Public FdisplayProductName
	Public FStandardCategoryId
	Public FprodSpecCd
	Public ForderMakingRd
	Public FAdultType
	Public FOrderMaxNum
	Public FCateMapCnt
	Public FDeliverytype
	Public FDefaultdeliverytype
	Public FDefaultfreeBeasongLimit
	Public FOptionCnt
	Public FRegedOptCnt
	Public FRctSellCNT
	Public FAccFailCNT
	Public FLastErrStr
	Public FInfoDiv
	Public FOptAddPrcCnt
	Public FOptAddPrcRegType
	Public FItemDiv
	Public FOrgSuplyCash
	Public FIsusing
	Public FKeywords
	Public FVatinclude
	Public FOrderComment
	Public FBasicImage
	Public FbasicImage600
	Public FbasicImage1000
	Public FbasicImage600str
	Public FbasicImage1000str
	Public FBasicimageNm
	Public FMainImage
	Public FMainImage2
	Public FSourcearea
	Public FMakername
	Public Fitemsize
	Public FItemsource
	Public FUsingHTML
	Public FItemcontent

	Public FtenCateLarge
	Public FtenCateMid
	Public FtenCateSmall
	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FItemkindname
	Public FNitypecd
	Public FCateKey
	Public FDepth1Name
	Public FDepth2Name
	Public FDepth3Name
	Public FDepth4Name

	Public FrequireMakeDay
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public Fregitemname
	Public FregImageName
	Public FSpecialPrice
	Public FStartDate
	Public FEndDate
	Public FNotSchIdx
	Public FLastUpdateUserId
	Public FIdx
	Public FNewItemName
	Public FLimitCount
	Public FItemoption
	Public FOptionname
	Public FOptlimitno
	Public FOptlimitsold

	Public Function getRegedOptionCnt
		Dim sqlStr
		sqlStr = ""
		sqlStr = sqlStr & " SELECT COUNT(*) as Cnt  "
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_lfmall_new_regedoption] "
		sqlStr = sqlStr & " WHERE itemoption <> '0000' "
		sqlStr = sqlStr & " and itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			getRegedOptionCnt = rsget("Cnt")
		rsget.Close
	End Function

	Public Function getOrderMaxNum()
		getOrderMaxNum = FOrderMaxNum
		If FOrderMaxNum > "999" Then
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

	Function getLimitEa2()
		dim ret : ret = (FLimitno-FLimitSold)
		If FLimityn = "Y" Then
			ret = FLimitno - FLimitSold - 5
		Else
			ret = 999
		End If

		if (ret < 1) Then ret = 0
		getLimitEa2 = ret
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

			If FLfmallPrice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
				End If
			End If
		End If
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FregImageName
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	Public Function IsAdultItem()
		Select Case FAdultType
			Case "1", "2"
				IsAdultItem = "Y"
			Case Else
				IsAdultItem = "N"
		End Select
	End Function

	'// lfmall 판매여부 반환
	Public Function getLfmallSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getLfmallSellYn = "Y"
			Else
				getLfmallSellYn = "N"
			End If
		Else
			getLfmallSellYn = "N"
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

	'최대 구매 수량
	Public Function getLimitLfmallEa()
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
		getLimitLfmallEa = ret
	End Function

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

	Public Function getLfmallContParamToReg()
		Dim strRst, strSQL
		strRst = ""
		strRst = strRst & "<style type='text/css'>BODY { font-size: 12px; font-family: '돋움','돋움' }</style>"
		strRst = strRst & "<p align='center'><img src='http://fiximage.10x10.co.kr/web2008/etc/top_notice_lfmall.jpg'></p><br>"
		If Fitemsize <> "" Then
		 	strRst = strRst & "- 사이즈 : " & Fitemsize & "<br>"
		End if
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
		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_lfmall.jpg"">")
		getLfmallContParamToReg = strRst
	End Function

	Public Function getLfmallItemCertInfo
		Dim buf, strSql
		Dim KcGb, KcCertifyNo
		strSql = ""
		strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_API_LFMALL_CertInfo_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				KcGb		= rsget("KcGb")
				KcCertifyNo = rsget("KcCertifyNo")
				rsget.MoveNext
			Next
		End If
		rsget.Close
		buf = ""
		buf = buf & "			<KcCert>"&VBCRLF
		buf = buf & "				<KcGb>"&KcGb&"</KcGb>"&VBCRLF											'#인증정보 | 10.   통과 (삭제) - 전송시 오류메시지 전달, 20.   받지 않음 (삭제) - 전송시 오류메시지 전달, 30    대상 아님, 40    [어린이제품] 안전인증 - 인증번호 필수, 41    [어린이제품] 안전확인 - 인증번호 필수, 42    [어린이제품] 공급자적합성확인, 50    [전기용품] 안전인증 - 인증번호 필수, 51    [전기용품] 안전확인 - 인증번호 필수, 52    [전기용품] 공급자적합성확인, 60    [생활용품] 안전인증 - 인증번호 필수, 61    [생활용품] 안전확인 - 인증번호 필수, 62    [생활용품] 공급자 적합성확인, 63    [생활용품] 안전기준 준수대상, 64    [생활용품] 어린이보호포장, 70.   [방송통신기자재] 적합인증, 71.   [방송통신기자개] 적합등록 , 72.   [방송통신기자재] 잠적인증, 80.   상세 설명에 별도표기
		buf = buf & "				<KcCertifyNo>"&KcCertifyNo&"</KcCertifyNo>"&VBCRLF						'인증번호 | 인증번호가 필수인 경우는 반드시 입력
		buf = buf & "			</KcCert>"&VBCRLF
		getLfmallItemCertInfo = buf
	End Function

	Public Function getLfmallLeadTime()
		Dim CateLargeMid, leadTime
		If isNull(FtenCateLarge) AND isNull(FtenCateMid) Then
			FtenCateLarge = "999"
			FtenCateMid = "999"
		End If

		CateLargeMid = CStr(FtenCateLarge) & CStr(FtenCateMid)
		Select Case CateLargeMid
			Case "030331", "040010", "040011", "040020", "040030", "040040", "040050", "040070", "040080", "040090", "040100", "040121", "055070", "055080"
				leadTime = 15
			Case "050777", "055090", "055100", "055110", "055120", "060070"
				leadTime = 10
			Case "050045", "080007", "080010", "080020", "080030", "080031", "080040", "080050", "080051", "080060", "080070", "080071", "080080", "080090", "090005", "090010", "090011", "090020", "090040"
				leadTime = 7
			Case "010130", "010140", "010150", "010160", "020001", "020010", "020020", "020030", "020060", "020070", "020090", "020100", "020110", "020111", "020130", "020222", "020333", "020334", "025014", "025015", "025020", "025022", "025050", "025060", "025080", "025100", "025102", "025103", "025104", "025105", "025106", "025108", "025109", "025110", "025111", "025112", "025113", "025114", "025116", "025120", "025456", "030300", "030320", "030330", "030340", "030345", "030350", "030360", "030370", "030380", "030420", "030421", "030450", "035009", "035010", "035011", "035012", "035013", "035014", "035015", "035016", "035017", "035018", "035019", "035020", "035021", "035022", "045001", "045002", "045003", "045004", "045005", "045006", "045007", "045008", "045009", "045010", "045011", "050010", "050020", "050030", "050040", "050050", "050070", "050110", "050120", "050666", "055222", "060010", "060020", "060040", "060050", "060060", "060080", "060090", "060120", "060130", "060140", "060150", "060160", "070010", "070020", "070030", "070040", "070050", "070070", "070110", "070120", "070140", "070150", "070160", "070200", "070201", "070202", "070203", "090060", "090061", "090070", "090071", "090080"
				leadTime = 5
			Case Else
				leadTime = 3
		End Select

		If FprodSpecCd = "20" and leadTime < 10 Then		''주문제작
			leadTime = 10
		End If

		getLfmallLeadTime = leadTime
	End Function


	' Public Function getLfmallOptionParam
	' 	Dim buf, isOptSoldout
	' 	Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, optDc1, optDc2
	' 	Dim optNm, optDc, optLimit, itemoption, MultiYN
	' 	chkMultiOpt = false
	' 	MultiTypeCnt = 0

	' 	buf = ""
	' 	If FOptionCnt = 0 Then			'단품
	' 		buf = buf & "			<Option>"&VBCRLF
	' 		buf = buf & "				<OptionCode><![CDATA[0000]]></OptionCode>"&VBCRLF	'옵션코드 | 사이즈 단위로 키값이 필요할 경우 사용
	' 		buf = buf & "				<OptionType>1</OptionType>"&VBCRLF				'옵션유형 | [1]조합형 [2]콤보형
	' 		buf = buf & "				<OptionNm1>선택</OptionNm1>"&VBCRLF				'옵션명1 | 첫번째 옵션의 명
	' 		buf = buf & "				<OptionValue1>단일상품</OptionValue1>"&VBCRLF	'옵션값1 | 첫번째 옵션 값
	' 		buf = buf & "				<OptionNm2></OptionNm2>"&VBCRLF			'옵션명2 | 두번째 옵션의 명
	' 		buf = buf & "				<OptionValue2></OptionValue2>"&VBCRLF	'옵션값2 | 두번째 옵션 값
	' 		buf = buf & "				<CurrentStockQty>"&getLimitLfmallEa&"</CurrentStockQty>"&VBCRLF	'재고수량 | 오직 숫자로만 입력
	' 		buf = buf & "				<ExtraCharge>0</ExtraCharge>"&VBCRLF	'추가금액 | 오직 숫자로만 입력
	' 		buf = buf & "				<SoldoutYn>N</SoldoutYn>"&VBCRLF		'품절여부 | 상품등록 시에는 'N'으로 세팅. 품절이 있을 경우 'Y' 로 전송
	' 		buf = buf & "			</Option>"&VBCRLF
	' 	Else
	' 		strSql = ""
	' 		strSql = "exec [db_item].[dbo].sp_Ten_ItemOptionMultipleTypeList " & FItemid
	'         rsget.CursorLocation = adUseClient
	' 		rsget.CursorType = adOpenStatic
	' 		rsget.LockType = adLockOptimistic
	'         rsget.Open strSql, dbget
	' 		If Not(rsget.EOF or rsget.BOF) Then
	' 			chkMultiOpt = true
	' 			MultiTypeCnt = rsget.recordcount
	' 			Do until rsget.EOF
	' 				arrMultiTypeNm = arrMultiTypeNm & replaceRst(db2Html(rsget("optionTypeName")))&","
	' 				rsget.MoveNext
	' 			Loop
	' 		End If
	' 		rsget.Close

	' 		If chkMultiOpt = false Then		'일반 옵션 일 경우
	' 			strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
	' 			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
	' 			strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
	' 			rsget.CursorLocation = adUseClient
	' 			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	' 			If Not(rsget.EOF or rsget.BOF) then
	' 				If db2Html(rsget("optionTypeName")) <> "" Then
	' 					optNm = db2Html(rsget("optionTypeName"))
	' 				Else
	' 					optNm = "옵션"
	' 				End If
	' 				Do until rsget.EOF
	' 					optLimit = rsget("optLimit")
	' 					optLimit = optLimit-5
	' 					If (optLimit < 1) Then optLimit = 0
	' 					If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
	' 					itemoption	= rsget("itemoption")
	' 					optDc		= db2Html(rsget("optionname"))
	' 					optIsusing	= rsget("isusing")
	' 					optSellYn	= rsget("optsellyn")
	' 					optaddprice	= rsget("optaddprice")

	' 					If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
	' 						isOptSoldout = "Y"
	' 					Else
	' 						isOptSoldout = "N"
	' 					End If
	' 					buf = buf & "			<Option>"&VBCRLF
	' 					buf = buf & "				<OptionCode><![CDATA["&itemoption&"]]></OptionCode>"&VBCRLF
	' 					buf = buf & "				<OptionType>1</OptionType>"							'옵션유형 | [1]조합형 [2]콤보형
	' 					buf = buf & "				<OptionNm1><![CDATA["&optNm&"]]></OptionNm1>"&VBCRLF
	' 					buf = buf & "				<OptionValue1><![CDATA["&optDc&"]]></OptionValue1>"&VBCRLF
	' 					buf = buf & "				<OptionNm2 />"&VBCRLF
	' 					buf = buf & "				<OptionValue2 />"&VBCRLF
	' 					buf = buf & "				<CurrentStockQty>"&optLimit&"</CurrentStockQty>"&VBCRLF
	' 					buf = buf & "				<ExtraCharge>"&optaddprice&"</ExtraCharge>"&VBCRLF
	' 					buf = buf & "				<SoldoutYn>"&isOptSoldout&"</SoldoutYn>"&VBCRLF
	' 					buf = buf & "			</Option>"&VBCRLF
	' 					rsget.MoveNext
	' 				Loop
	' 			End If
	' 			rsget.Close
	' 		Else
	' 			If Right(arrMultiTypeNm,1) = "," Then
	' 				arrMultiTypeNm = Left(arrMultiTypeNm, Len(arrMultiTypeNm) - 1)
	' 			End If

	' 			If MultiTypeCnt = 2 Then	'2중 옵션일 경우
	' 				type1 				= Split(arrMultiTypeNm, ",")(0)
	' 				type2 				= Split(arrMultiTypeNm, ",")(1)
	' 			End If

	' 			strSql = ""
	' 			strSql = strSql & "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
	' 			strSql = strSql & ",(case when CHARINDEX(',',optionname)=0 then 'N' else 'Y' end) as MultiYN "	'상품코드 : 1116421 옵션이 일반,복합 섞임; 2015-09-11 진영//발견 후 추가
	' 			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
	' 			strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
	' 			rsget.CursorLocation = adUseClient
	' 			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	' 			If Not(rsget.EOF or rsget.BOF) then
	' 				Do until rsget.EOF
	' 					optLimit = rsget("optLimit")
	' 					optLimit = optLimit-5
	' 					If (optLimit < 1) Then optLimit = 0
	' 					If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
	' 					itemoption	= rsget("itemoption")
	' 					optDc		= db2Html(rsget("optionname"))
	' 					optDc		= replaceRst(optDc)
	' 					optIsusing	= rsget("isusing")
	' 					optSellYn	= rsget("optsellyn")
	' 					optaddprice	= rsget("optaddprice")
	' 					MultiYN		= rsget("MultiYN")

	' 					If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
	' 						isOptSoldout = "Y"
	' 					Else
	' 						isOptSoldout = "N"
	' 					End If

	' 					If MultiTypeCnt = 2 Then
	' 						If MultiYN = "Y" Then
	' 							optDc1 = split(optDc,",")(0)
	' 							optDc2 = split(optDc,",")(1)

	' 							buf = buf & "			<Option>"&VBCRLF
	' 							buf = buf & "				<OptionCode><![CDATA["&itemoption&"]]></OptionCode>"&VBCRLF
	' 							buf = buf & "				<OptionType>1</OptionType>"&VBCRLF
	' 							buf = buf & "				<OptionNm1><![CDATA["&type1&"]]></OptionNm1>"&VBCRLF
	' 							buf = buf & "				<OptionValue1><![CDATA["&optDc1&"]]></OptionValue1>"&VBCRLF
	' 							buf = buf & "				<OptionNm2><![CDATA["&type2&"]]></OptionNm2>"&VBCRLF
	' 							buf = buf & "				<OptionValue2><![CDATA["&optDc2&"]]></OptionValue2>"&VBCRLF
	' 							buf = buf & "				<CurrentStockQty>"&optLimit&"</CurrentStockQty>"&VBCRLF
	' 							buf = buf & "				<ExtraCharge>"&optaddprice&"</ExtraCharge>"&VBCRLF
	' 							buf = buf & "				<SoldoutYn>"&isOptSoldout&"</SoldoutYn>"&VBCRLF
	' 							buf = buf & "			</Option>"&VBCRLF
	' 						End If
	' 					End If
	' 					rsget.MoveNext
	' 				Loop
	' 			end if
	' 			rsget.Close
	' 		End If
	' 	End If
	' 	getLfmallOptionParam = buf
	' End Function

	Public Function getLfmallOptionParam
		Dim buf, isOptSoldout, lp
		Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, optDc1, optDc2
		Dim optNm, optDc, optLimit, itemoption, MultiYN
		chkMultiOpt = false
		MultiTypeCnt = 0
		lp = 0

		buf = ""
		If FOptionCnt = 0 Then			'단품
			buf = buf & "			<Option>"&VBCRLF
			buf = buf & "				<OptionCode><![CDATA[0000]]></OptionCode>"&VBCRLF	'옵션코드 | 사이즈 단위로 키값이 필요할 경우 사용
			buf = buf & "				<OptionType>1</OptionType>"&VBCRLF				'옵션유형 | [1]조합형 [2]콤보형
			buf = buf & "				<OptionNm1>1</OptionNm1>"&VBCRLF				'옵션명1 | 첫번째 옵션의 명
			buf = buf & "				<OptionValue1>단일상품</OptionValue1>"&VBCRLF	'옵션값1 | 첫번째 옵션 값
			buf = buf & "				<OptionNm2></OptionNm2>"&VBCRLF			'옵션명2 | 두번째 옵션의 명
			buf = buf & "				<OptionValue2></OptionValue2>"&VBCRLF	'옵션값2 | 두번째 옵션 값
			buf = buf & "				<CurrentStockQty>"&getLimitLfmallEa&"</CurrentStockQty>"&VBCRLF	'재고수량 | 오직 숫자로만 입력
			buf = buf & "				<ExtraCharge>0</ExtraCharge>"&VBCRLF	'추가금액 | 오직 숫자로만 입력
			buf = buf & "				<SoldoutYn>N</SoldoutYn>"&VBCRLF		'품절여부 | 상품등록 시에는 'N'으로 세팅. 품절이 있을 경우 'Y' 로 전송
			buf = buf & "			</Option>"&VBCRLF
		Else
			strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					lp = lp + 1
					optLimit = rsget("optLimit")
					optLimit = optLimit-5
					If (optLimit < 1) Then optLimit = 0
					If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					itemoption	= rsget("itemoption")
					optDc		= db2Html(rsget("optionname"))
					optIsusing	= rsget("isusing")
					optSellYn	= rsget("optsellyn")
					optaddprice	= rsget("optaddprice")

					If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
						isOptSoldout = "Y"
					Else
						isOptSoldout = "N"
					End If
					buf = buf & "			<Option>"&VBCRLF
					buf = buf & "				<OptionCode><![CDATA["&itemoption&"]]></OptionCode>"&VBCRLF
					buf = buf & "				<OptionType>1</OptionType>"							'옵션유형 | [1]조합형 [2]콤보형
					buf = buf & "				<OptionNm1>옵션1</OptionNm1>"&VBCRLF
					buf = buf & "				<OptionValue1>"& lp &"</OptionValue1>"&VBCRLF
					buf = buf & "				<OptionNm2>옵션2</OptionNm2>"&VBCRLF
					buf = buf & "				<OptionValue2><![CDATA["&optDc&"]]></OptionValue2>"&VBCRLF
					buf = buf & "				<CurrentStockQty>"&optLimit&"</CurrentStockQty>"&VBCRLF
					buf = buf & "				<ExtraCharge>"&optaddprice&"</ExtraCharge>"&VBCRLF
					buf = buf & "				<SoldoutYn>"&isOptSoldout&"</SoldoutYn>"&VBCRLF
					buf = buf & "			</Option>"&VBCRLF
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If
		getLfmallOptionParam = buf
	End Function

	Public Function getLfmallOptionQtyParam
		Dim buf, isOptSoldout, lp
		Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, optDc1, optDc2
		Dim optNm, optDc, optLimit, itemoption, MultiYN
		chkMultiOpt = false
		MultiTypeCnt = 0
		lp = 0

		buf = ""
		If FOptionCnt = 0 Then			'단품
			buf = buf & "			<Option>"&VBCRLF
			buf = buf & "				<OptionNm1>1</OptionNm1>"&VBCRLF				'옵션명1 | 첫번째 옵션의 명
			buf = buf & "				<OptionValue1>단일상품</OptionValue1>"&VBCRLF	'옵션값1 | 첫번째 옵션 값
			buf = buf & "				<OptionNm2></OptionNm2>"&VBCRLF			'옵션명2 | 두번째 옵션의 명
			buf = buf & "				<OptionValue2></OptionValue2>"&VBCRLF	'옵션값2 | 두번째 옵션 값
			buf = buf & "				<CurrentStockQty>"&getLimitLfmallEa&"</CurrentStockQty>"&VBCRLF	'재고수량 | 오직 숫자로만 입력
			buf = buf & "			</Option>"&VBCRLF
		Else
			strSql = "Select itemoption, isusing, optsellyn, optaddprice, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit "
			strSql = strSql & " From [db_item].[dbo].tbl_item_option "
			strSql = strSql & " where isUsing='Y' and optsellyn='Y' and itemid=" & FItemid
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				Do until rsget.EOF
					lp = lp + 1
					optLimit = rsget("optLimit")
					optLimit = optLimit-5
					If (optLimit < 1) Then optLimit = 0
					If (FLimitYN <> "Y") Then optLimit = CDEFALUT_STOCK
					itemoption	= rsget("itemoption")
					optDc		= db2Html(rsget("optionname"))
					optIsusing	= rsget("isusing")
					optSellYn	= rsget("optsellyn")
					optaddprice	= rsget("optaddprice")

					If (optIsusing <> "Y") OR (optSellYn <> "Y") OR (optLimit = 0) Then
						isOptSoldout = "Y"
					Else
						isOptSoldout = "N"
					End If
					buf = buf & "			<Option>"&VBCRLF
					buf = buf & "				<OptionNm1>옵션1</OptionNm1>"&VBCRLF
					buf = buf & "				<OptionValue1>"& lp &"</OptionValue1>"&VBCRLF
					buf = buf & "				<OptionNm2>옵션2</OptionNm2>"&VBCRLF
					buf = buf & "				<OptionValue2><![CDATA["&optDc&"]]></OptionValue2>"&VBCRLF
					buf = buf & "				<CurrentStockQty>"&optLimit&"</CurrentStockQty>"&VBCRLF
					buf = buf & "			</Option>"&VBCRLF
					rsget.MoveNext
				Loop
			End If
			rsget.Close
		End If
		getLfmallOptionQtyParam = buf
	End Function

	Public Function getLfmallOptionZeroQtyParam
		Dim buf, isOptSoldout, lp
		Dim strRst, strSql, chkMultiOpt, optIsusing, optSellYn, optaddprice, MultiTypeCnt, arrMultiTypeNm, type1, type2, optDc1, optDc2
		Dim optNm, optDc, optLimit, itemoption, MultiYN

		buf = ""
		strSql = "Select itemid, itemoption, optionNm1, optionValue1, optionNm2, optionValue2, outmallSellyn, outmalllimitno "
		strSql = strSql & " From db_etcmall.dbo.tbl_lfmall_new_regedoption "
		strSql = strSql & " where itemid=" & FItemid
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				buf = buf & "			<Option>"&VBCRLF
				buf = buf & "				<OptionNm1>"& rsget("OptionNm1") &"</OptionNm1>"&VBCRLF
				buf = buf & "				<OptionValue1>"& rsget("optionValue1") &"</OptionValue1>"&VBCRLF
				buf = buf & "				<OptionNm2>"& rsget("optionNm2") &"</OptionNm2>"&VBCRLF
				buf = buf & "				<OptionValue2><![CDATA["&rsget("optionValue2")&"]]></OptionValue2>"&VBCRLF
				buf = buf & "				<CurrentStockQty>0</CurrentStockQty>"&VBCRLF
				buf = buf & "			</Option>"&VBCRLF
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getLfmallOptionZeroQtyParam = buf
	End Function

	'상품품목정보
    public function getLfmallItemInfoCd()
		Dim buf, NiCd, NiValueGb, NiValue, i, strSql
		buf = ""
		strSql = ""
		strSql = strSql & " EXEC db_etcmall.[dbo].[usp_API_LFMALL_InfoCodeMap_Get] " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If i = 1 Then
					buf = buf & "			<NiTypeCd>"& rsget("nitypecd") &"</NiTypeCd>"&VBCRLF
				End If
				buf = buf & "			<ProdNoti>"&VBCRLF
				buf = buf & "				<NiCd>"& rsget("NiCd") &"</NiCd>"&VBCRLF						'#고시항목코드
				buf = buf & "				<NiValueGb>"& rsget("NiValueGb") &"</NiValueGb>"&VBCRLF			'#고시항목값구분 | T:TEXT ( 'T' 만 사용 )
				buf = buf & "				<NiValue><![CDATA["& rsget("NiValue") &"]]></NiValue>"&VBCRLF	'#고시항목값 | 015(KC안전인증 대상 유무), 020(품질보증서 제공 유무), 097(수입여부) 항목일 때는 Y/N 으로 전송
				buf = buf & "			</ProdNoti>"&VBCRLF
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getLfmallItemInfoCd = buf
	End Function

	'이미지 정보
    public function getLfmallAddImageParam()
		Dim buf, strSQL, i, addImgUrl, fImage

		buf = ""
		buf = buf & "			<Image>"&VBCRLF
		If NOT(isnull(FbasicImage1000)) and NOT(FbasicImage1000 = "") Then
			fImage = FbasicImage1000str
		ElseIf NOT(isnull(FbasicImage600)) and NOT(FbasicImage600 = "") Then
			fImage = FbasicImage600str
		Else
			fImage = FbasicImage
		End If
		fImage = fImage & "/10x10/thumbnail/1500x1500/quality/85/"
		buf = buf & "				<Image1><![CDATA["&fImage&"]]></Image1>"&VBCRLF
		buf = buf & "			</Image>"&VBCRLF

		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				addImgUrl = ""
				If (NOT(IsNULL(rsget("addimage_1000")))) AND (rsget("addimage_1000") <> "") AND (Len(rsget("addimage_1000"))) > 0 Then
					addImgUrl = "add" & rsget("gubun") & "_1000/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_1000")
				ElseIf (NOT(IsNULL(rsget("addimage_600")))) AND (rsget("addimage_600") <> "") AND (Len(rsget("addimage_600"))) > 0 Then
					addImgUrl = "add" & rsget("gubun") & "_600/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_600")
				Else
					addImgUrl = "add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400")
				End If

				If rsget("imgType") = "0" Then
					buf = buf & "			<Image>"&VBCRLF
					buf = buf & "				<Image1><![CDATA[http://webimage.10x10.co.kr/image/"& addImgUrl &"/10x10/thumbnail/1500x1500/quality/85/]]></Image1>"&VBCRLF
					buf = buf & "			</Image>"&VBCRLF
				End If
				rsget.MoveNext
				If i>=4 Then Exit For
			Next
		End If
		rsget.Close
		getLfmallAddImageParam = buf
	End Function

	'상품정보 멀티옵션 XML
	Public Function getlfmallItemRegParameter(iType)
		Dim strRst, ActionType, ImageChangeYn, leadtime, prodSpecCd
		leadtime = getLfmallLeadTime()

		If leadtime > 3 Then
			prodSpecCd = 20
		Else
			prodSpecCd = 10
		End If

		Select Case iType
			Case "REG"
				ActionType = "I"
				ImageChangeYn = "N"
			Case "EDIT"
				ActionType = "U"
				If isImageChanged Then
					ImageChangeYn = "Y"
				Else
					ImageChangeYn = "N"
				End If

				If FlfmallStatCD <> "7" Then
					ImageChangeYn = "Y"
				End If
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8"" ?>"&VBCRLF
		strRst = strRst & "<ProductInfo>"&VBCRLF
		strRst = strRst & "	<Header>"&VBCRLF
		strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
		strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
		strRst = strRst & "		<Format>XML</Format>"&VBCRLF
		strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
		strRst = strRst & "	</Header>"&VBCRLF
		strRst = strRst & "	<Body>"&VBCRLF
		strRst = strRst & "		<Product>"&VBCRLF
		strRst = strRst & "			<ActionType>"&ActionType&"</ActionType>"&VBCRLF				'#처리유형 | I : 신규등록, U : 정보갱신, D : 상품삭제시 2. 상품삭제 시 SupplyProductCode만 필요하다."
		strRst = strRst & "			<SupplyProductCode>"&FItemid&"</SupplyProductCode>"&VBCRLF	'#입점업체상품코드 | 입점하는 업체의 상품코드(유니크해야한다)
		strRst = strRst & "			<ProductCode>"& Chkiif(ActionType="I", "", FLfmallGoodNo) &"</ProductCode>"&VBCRLF	'#LG패션에서 부여한 상품코드. 등록 시 반환, 수정 시 필수
		strRst = strRst & "			<BrandCode>EIHA</BrandCode>"&VBCRLF							'#브랜드코드 | LF Mall에서 지정한 값. [3.브랜드목록] 참조
		strRst = strRst & "			<ItemKindCode>"&FItemKindCode&"</ItemKindCode>"&VBCRLF		'#품목코드 | LF Mall에서 지정한 값. [4.품목목록] 참조
		strRst = strRst & "			<EmodelName></EmodelName>"&VBCRLF							'모델명
		strRst = strRst & "			<SeasonCode>G</SeasonCode>"&VBCRLF							'#시즌코드 | A - 봄, B - 여름,  C - 가을, D - 겨울, E - S/S, F - F/W, G - 사계절
		strRst = strRst & "			<Color1Code>B0</Color1Code>"&VBCRLF							'#색상코드 | 상품단위가 스타일단위일 경우 'XX', 아닐 경우는 [5.색상목록]  참조
		strRst = strRst & "			<ProductName><![CDATA["&FdisplayProductName&"]]></ProductName>"&VBCRLF	'#상품명 | LF Mall에서 사용할 상품명
		strRst = strRst & "			<SearchKeyword><![CDATA["&RightCommaDel(Trim(FKeywords))&"]]></SearchKeyword>"&VBCRLF	'검색키워드 | 검색으로 추출될 수 있는 단어의 조합 (없을 경우 상품명)
		strRst = strRst & "			<ItemYear></ItemYear>"&VBCRLF								'출시년도 | 미필수
		strRst = strRst & "			<ListPrice>"&Clng(FOrgprice)&"</ListPrice>"&VBCRLF				'#최초소비자가 | 정상가. 오직 숫자로만 입력 특정업체만 반영가능. 담당MD와 상의요망
		strRst = strRst & "			<ProductPrice>"&Clng(MustPrice())&"</ProductPrice>"&VBCRLF		'#현재판매가 | 오직 숫자로만 입력
		strRst = strRst & "			<MakeName><![CDATA["&FMakerName&"]]></MakeName>"&VBCRLF		'#제조사
'		strRst = strRst & "			<NativeCode><![CDATA[KR]]></NativeCode>"&VBCRLF				'원산지코드 | 코드조회에 원산지코드 참고, 원산지가 여러 개인 경우 원산지코드=XX로 입력 후 원산지명에 텍스트 입력
		strRst = strRst & "			<NativeName><![CDATA["&Chkiif(fnStrLength(Fsourcearea) >= 50, chrbyte(Fsourcearea,20,""), Fsourcearea)&"]]></NativeName>"&VBCRLF			'#원산지명
		strRst = strRst & "			<ProductDesc><![CDATA["& getLfmallContParamToReg() &"]]></ProductDesc>"&VBCRLF	'#상품설명
		strRst = strRst & "			<MaterialDesc><![CDATA["& FItemsource &"]]></MaterialDesc>"&VBCRLF	'#소재설명
		strRst = strRst & "			<ShippingYn>N</ShippingYn>"&VBCRLF							'해외배송상품여부 | Y, N
		strRst = strRst & "			<DeliveryFeeFreeYn>N</DeliveryFeeFreeYn>"&VBCRLF			'배송비무료여부 | Y, N(현재사용안함)
		strRst = strRst & "			<DeliveryFee>3000</DeliveryFee>"&VBCRLF						'기본배송비 | 미입력시 해당브랜드의 기본배송비 입력.
		strRst = strRst & "			<MinOrdAmt>50000</MinOrdAmt>"&VBCRLF						'무료배송비최소구매금액 | 미입력시 해당브랜드의 무료배송비최소구매금액 입력.
		strRst = strRst & "			<jejuDeliveryFee>3000</jejuDeliveryFee>"&VBCRLF				'제주추가배송비 | 미입력시 해당브랜드의 기본 제주배송비 입력.
		strRst = strRst & "			<islandDeliveryFee>3000</islandDeliveryFee>"&VBCRLF			'도서산간비 | 미입력시 해당브랜드의 도서산간비 입력.
		strRst = strRst & "			<ReturnAbleYn>Y</ReturnAbleYn>"&VBCRLF						'#반품가능여부
		strRst = strRst & "			<ChangeAbleYn>Y</ChangeAbleYn>"&VBCRLF						'#교환가능여부
		strRst = strRst & "			<ReturnFeeFreeYn>N</ReturnFeeFreeYn>"&VBCRLF				'#반품배송료무료여부
		strRst = strRst & "			<ChangeFeeFreeYn>N</ChangeFeeFreeYn>"&VBCRLF				'#교환배송료무료여부
		strRst = strRst & "			<EtcMemo1></EtcMemo1>"&VBCRLF								'기타메모1 | 별도의 정보가 필요할 경우 사용한다, 예) 사은품정보
		strRst = strRst & "			<EtcMemo2></EtcMemo2>"&VBCRLF								'기타메모2 | 요약정보가 있을경우 입력
		strRst = strRst & "			<ProdType>10</ProdType>"&VBCRLF								'#상품타입 | 10 일반상품, 20 LG패션상품권, 30 디지털상품권
		strRst = strRst & "			<OptionSetYn>N</OptionSetYn>"&VBCRLF						'#옵션설정여부 | 상품단위가 스타일단위가 아닌 경우 'Y'로 입력한다.
		strRst = strRst & "			<OptionSoldOutDisplayYn>N</OptionSoldOutDisplayYn>"&VBCRLF	'품절시미노출여부 | Y : 품절 시 노출, N : 품절 시 미노출
		strRst = strRst & "			<StandardCategoryId>"&FStandardCategoryId&"</StandardCategoryId>"&VBCRLF	'#카테고리ID | 카테고리ID(신규카테고리 sheet 참조)
'		strRst = strRst & "			<testKcGb>30</testKcGb>"&VBCRLF								'매뉴얼에 누락
'		strRst = strRst & "			<ImportationGb>10</ImportationGb>"&VBCRLF					'매뉴얼에 누락
		strRst = strRst & "			<DeliveryMethod>10</DeliveryMethod>"&VBCRLF					'배송방법 | 10 택배/소포/등기, 20 직접 배송(화물 배달), 30 병행수입, 40 구매대행 , 50 해외배송
		strRst = strRst & "			<OverseasAgency></OverseasAgency>"&VBCRLF					'구매대행업자 | 배송방법 40 구매대행 선택시 필수
		strRst = strRst & "			<OverseasSeller></OverseasSeller>"&VBCRLF					'해외판매자 |배송방법 40 구매대행 선택시 필수
		strRst = strRst & "			<OverseasBizno></OverseasBizno>"&VBCRLF						'구매대행사업자번호 | 배송방법 40 구매대행 선택시 필수
		strRst = strRst & "			<OverseasAddr></OverseasAddr>"&VBCRLF						'해외판매자주소 | 배송방법 40 구매대행 선택시 필수
		strRst = strRst & "			<StatusGb>10</StatusGb>"&VBCRLF								'상품구분 | 10 신상품, 20 중고상품
		strRst = strRst & getLfmallItemCertInfo()
		strRst = strRst & "			<ImportationGb>10</ImportationGb>"&VBCRLF					'#제조수입구분 | 10 대상아님, 20 구매대행상품, 30 병행수입상품
		strRst = strRst & "			<codYn>N</codYn>"&VBCRLF									'#착불여부 | Y, N
		strRst = strRst & "			<prodSpecCd>"&prodSpecCd&"</prodSpecCd>"&VBCRLF				'#상품특성 | 10 일반상품, 20 주문제작, 30 설치상품, 40 신선/냉동/냉장식품, 50 일반식품
'		strRst = strRst & "			<orderMakingRd>"&ForderMakingRd&"</orderMakingRd>"&VBCRLF	'제작소요일 | nn일(2자리) -- sp에 포함됨
		strRst = strRst & "			<orderMakingRd>"&leadtime&"</orderMakingRd>"&VBCRLF			'제작소요일 | nn일(2자리) -- sp에 포함됨
		strRst = strRst & "			<ImageChangeYn>"&ImageChangeYn&"</ImageChangeYn>"&VBCRLF	'이미지변경여부 | Y 이미지갱신, N 이미지 미갱신 Image의 반영 여부를 정함
		strRst = strRst & "			<SeparateSettingCost></SeparateSettingCost>"&VBCRLF			'별도설치비 | 별도설치비가 필요한 상품일 경우 입력
		strRst = strRst & "			<ModelCd></ModelCd>"&VBCRLF									'모델번호 | 제조업체의 상품코드
		strRst = strRst & "			<ModelNm></ModelNm>"&VBCRLF									'모델명
		strRst = strRst & "			<ProdSexGb>X</ProdSexGb>"&VBCRLF							'성별 | 검색필터로 사용되는 항목으로 필수, F - 여성, M - 남성, U - 남여공용, B - 남아, G - 여아, Y - 아동공용, X - 없음
		strRst = strRst & "			<MinorBuyYn>"&CHKIIF(IsAdultItem= "Y", "N", "Y")&"</MinorBuyYn>"&VBCRLF		'미성년자 구매가능여부 | Y, N 기본값은 Y
		strRst = strRst & "			<VatCd>"&Chkiif(Fvatinclude="Y", "1", "2")&"</VatCd>"&VBCRLF		'부가세여부 | 부가세 구분을 숫자로 입력합니다.1 : 과세(기본값), 2 : 면세, 3 : 해당없음, 4 : 비과세, 5 : 영세
		strRst = strRst & "			<MinOrdQty>1</MinOrdQty>"&VBCRLF							'최소주문수량 | 고객의 한번의 주문으로 구매 가능한 상품의 최소 수량입력합니다. 최소주문수량은 숫자로 입력합니다.
		strRst = strRst & "			<MaxOrdQty>"&getOrderMaxNum&"</MaxOrdQty>"&VBCRLF			'최대주문수량 | 고객의 한번의 주문으로 구매 가능한 상품의 최대 수량입력합니다. 최대주문수량은 숫자로 입력합니다.
		strRst = strRst & "			<DeliveryDt></DeliveryDt>"&VBCRLF							'평균배송일자 | 평균배송일자는 숫자로 입력합니다.
		strRst = strRst & "			<GiftWrapYn>N</GiftWrapYn>"&VBCRLF							'#선물포장 가능여부
		strRst = strRst & "			<GiDd>"&leadtime&"</GiDd>"&VBCRLF							'#출고기한 | 상품의 출고기한일을 입력합니다. 상품특성[10,40,50]이면서, 배송방법[10]인 경우 최대 3일, 상품특성[30]인 경우 최대 30일, 배송방법[20,30,40,50]인 경우 최대 30일, 상품특성[20,60]인 경우 최대99일
'		strRst = strRst & "			<GiHr></GiHr>"&VBCRLF										'출고마감시간 | 당일발송상품의 경우, 측 출고기한일이 0일인 경우 1~24 사이의 시간을 입력합니다.
		strRst = strRst & "			<GiOptnCd>10</GiOptnCd>"&VBCRLF								'#출고조건옵션 | 출고조건 옵션을 입력합니다. 10 - 주말/공휴일 제외, 20 - 주말/공휴일 및 전일 제외, 30 - 일요일/공휴일 제외, 40 - 주말/공휴일 포함
		strRst = strRst & getLfmallOptionParam()
		strRst = strRst & getLfmallItemInfoCd()
		strRst = strRst & getLfmallAddImageParam()
		strRst = strRst & "		</Product>"&VBCRLF
		strRst = strRst & "	</Body>"&VBCRLF
		strRst = strRst & "</ProductInfo>"
		getlfmallItemRegParameter = strRst
	End Function

	'상품 판매상태 변경 XML
	Public Function getLFmallSellynParameter(ichgSellYn)
		Dim strRst, iProductStatusCode
		'10 : 정보오류 / 20 : 정보부족 / 40 : 승인대기 / 60 : 자동품절 / 70 : 일시중단 / 90 : 정상상품 / 99 : 영구중단
		Select Case ichgSellYn
			Case "Y"	iProductStatusCode = "90"
			Case "N"	iProductStatusCode = "70"
		End Select

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"&VBCRLF
		strRst = strRst & "<ProductStatusInfo>"&VBCRLF
		strRst = strRst & "	<Header>"&VBCRLF
		strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
		strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
		strRst = strRst & "		<Format>XML</Format>"&VBCRLF
		strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
		strRst = strRst & "	</Header>"&VBCRLF
		strRst = strRst & "	<Body>"&VBCRLF
		strRst = strRst & "		<ProductStatus>"&VBCRLF
		strRst = strRst & "			<ProductCode>"&FLfmallGoodNo&"</ProductCode>"&VBCRLF
		strRst = strRst & "			<ProductStatusCode>"&iProductStatusCode&"</ProductStatusCode>"&VBCRLF
		strRst = strRst & "		</ProductStatus>"&VBCRLF
		strRst = strRst & "	</Body>"&VBCRLF
		strRst = strRst & "</ProductStatusInfo>"&VBCRLF
		getLFmallSellynParameter = strRst
	End Function

	'상품 조회 XML
	Public Function getLfmallItemViewParameter()
		Dim strRst

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"&VBCRLF
		strRst = strRst & "<ProductInfo>"&VBCRLF
		strRst = strRst & "	<Header>"&VBCRLF
		strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
		strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
		strRst = strRst & "		<Format>XML</Format>"&VBCRLF
		strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
		strRst = strRst & "	</Header>"&VBCRLF
		strRst = strRst & "	<Body>"&VBCRLF
		strRst = strRst & "		<Product>"&VBCRLF
		strRst = strRst & "			<ProductCode>"&FLfmallGoodNo&"</ProductCode>"&VBCRLF
		strRst = strRst & "		</Product>"&VBCRLF
		strRst = strRst & "	</Body>"&VBCRLF
		strRst = strRst & "</ProductInfo>"&VBCRLF
		getLfmallItemViewParameter = strRst
	End Function

	'재고 수정 XML
	Public Function getlfmallQuantityParameter(izero)
		Dim strRst, strSQL, optCnt, optValue1Cnt, OptionSetYn
		OptionSetYn = "Y"

		strSQL = ""
		strSQL = strSQL & " SELECT COUNT(*) cnt "
		strSQL = strSQL & " FROM db_etcmall.[dbo].[tbl_lfmall_new_regedoption] "
		strSQL = strSQL & " WHERE itemid = '"& FItemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			optCnt = rsget("cnt")
		End If
		rsget.Close

		strSQL = ""
		strSQL = strSQL & " select top 10 itemid "
		strSQL = strSQL & " ,sum(case when len(itemoption) <> 4 then 1 else 0 end) as chgOptcnt "
		strSQL = strSQL & " ,count(*) as cnt "
		strSQL = strSQL & " from db_etcmall.dbo.tbl_lfmall_new_regedoption as r "
		strSQL = strSQL & " where itemid = '"& FItemid &"' "
		strSQL = strSQL & " and outmallsellyn = 'Y' "
		strSQL = strSQL & " group by itemid "
		strSQL = strSQL & " having count(*) = 2 "
		strSQL = strSQL & " and sum(case when len(itemoption) <> 4 then 1 else 0 end) > 0 "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			OptionSetYn = "N"
		End If
		rsget.Close

		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""UTF-8""?>"&VBCRLF
		strRst = strRst & "<ProductInfo>"&VBCRLF
		strRst = strRst & "	<Header>"&VBCRLF
		strRst = strRst & "		<AuthId><![CDATA["&AuthId&"]]></AuthId>"&VBCRLF
		strRst = strRst & "		<AuthKey><![CDATA["&AuthKey&"]]></AuthKey>"&VBCRLF
		strRst = strRst & "		<Format>XML</Format>"&VBCRLF
		strRst = strRst & "		<Charset>UTF-8</Charset>"&VBCRLF
		strRst = strRst & "	</Header>"&VBCRLF
		strRst = strRst & "	<Body>"&VBCRLF
		strRst = strRst & "		<Product>"&VBCRLF
		strRst = strRst & "			<ProductCode>"&FLfmallGoodNo&"</ProductCode>"&VBCRLF
		If optCnt <= 1 Then
			strRst = strRst & "			<OptionSetYn>N</OptionSetYn>"&VBCRLF						'#옵션설정여부 | 옵션이 있는 상품인 경우 Y로 입력하고 옵션명, 옵션값, 옵션별 재고수량을 입력한다. 옵션이 없는 상품인 경우 N을 입력하고 옵션별 재고수량 입력
			strRst = strRst & "			<ProductStockQty>"&getLimitEa2&"</ProductStockQty>"&VBCRLF	'#옵션설정 N인 경우 재고수량
		Else
			strRst = strRst & "			<OptionSetYn>"&OptionSetYn&"</OptionSetYn>"&VBCRLF						'#옵션설정여부 | 옵션이 있는 상품인 경우 Y로 입력하고 옵션명, 옵션값, 옵션별 재고수량을 입력한다. 옵션이 없는 상품인 경우 N을 입력하고 옵션별 재고수량 입력
			If izero = "Z" Then
				strRst = strRst & getLfmallOptionZeroQtyParam()
			Else
				strRst = strRst & getLfmallOptionQtyParam()
			End If
		End If
		strRst = strRst & "		</Product>"&VBCRLF
		strRst = strRst & "	</Body>"&VBCRLF
		strRst = strRst & "</ProductInfo>"&VBCRLF
		getlfmallQuantityParameter = strRst
	End Function

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class

Class CLfmall
	Public FItemList()
	Public FResultCount
	Public FOneItem
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
	Public FRectStartMargin
	Public FRectEndMargin
	Public FRectMakerid
	Public FRectLfmallGoodNo
	Public FRectMatchCate
	Public FRectoptExists
	Public FRectoptnotExists
	Public FRectEzwelNotReg
	Public FRectMinusMigin
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectLfmallYes10x10No
	Public FRectLfmallNo10x10Yes
	Public FRectLfmallKeepSell
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectoptAddprcExists
	Public FRectoptAddprcExistsExcept
	Public FRectoptAddPrcRegTypeNone
	Public FRectregedOptNull
	Public FRectFailCntExists
	Public FRectezwelDelOptErr
	Public FRectisMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectNotinmakerid
	Public FRectNotinitemid
	Public FRectScheduleNotInItemid
	Public FRectExcTrans
	Public FRectPriceOption
	Public FRectExtNotReg
	Public FRectReqEdit
	Public FRectDeliverytype
	Public FRectMwdiv
	Public FRectIsextusing
	Public FRectCisextusing

	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FsearchName

	Public FRectOrdType
	Public FRectIsSpecialPrice

	Public FRectIsGetDate
	Public FRectIdx

	Public Sub getLfmallNotRegOneItem
		Dim strSql, addSql, i
		strSql = ""
		strSql = strSql & "  EXEC [db_etcmall].[dbo].[usp_API_LFMALL_Reg_Get] '"& FRectItemID &"' "
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If Not(rsget.EOF or rsget.BOF) Then
			Set FOneItem = new CLfmallItem
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
				FOneItem.FbasicImage600		= rsget("basicimage600")
				FOneItem.FbasicImage1000	= rsget("basicimage1000")
				FOneItem.FbasicImage600str	= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage600")
				FOneItem.FbasicImage1000str	= "http://webimage.10x10.co.kr/image/basic1000/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage1000")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.Fitemsize			= rsget("itemsize")
				FOneItem.FItemsource		= rsget("itemsource")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FlfmallStatCD		= rsget("lfmallStatcd")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.Flfbrandcode 		= rsget("lfbrandcode")
				FOneItem.FItemKindCode 		= rsget("ItemKindCode")
				FOneItem.FSeasonCode		= rsget("SeasonCode")
				FOneItem.FColor1Code		= rsget("Color1Code")
				FOneItem.FdisplayProductName = rsget("displayProductName")
				FOneItem.FStandardCategoryId = rsget("StandardCategoryId")
				FOneItem.FprodSpecCd		= rsget("prodSpecCd")
				FOneItem.ForderMakingRd		= rsget("orderMakingRd")
				FOneItem.FAdultType			= rsget("adultType")
				FOneItem.FOrderMaxNum		= rsget("orderMaxNum")
		End If
		rsget.Close
	End Sub

	Public Sub getLfmallEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			'선택상품이 있다면
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & " ,m.lfmallGoodNo, m.lfmallSellyn, m.regImageName, isNull(m.lfmallprice, 0) as lfmallprice "
		strSql = strSql & " , c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, c.safetyyn, c.safetyNum, c.safetyDiv "
		strSql = strSql & " , IsNULL(c.itemsize,'') as itemsize, IsNULL(c.itemsource,'') as itemsource "
		strSql = strSql & " , ni.itemkindcode as ItemKindCode, 'G' as SeasonCode, 'XX' as Color1Code, am.CateKey as StandardCategoryId, m.lfmallStatcd "
		strSql = strSql & " , '[텐바이텐] ' + i.itemname as displayProductName "
		strSql = strSql & " , CASE WHEN i.itemdiv in ('06', '16') THEN '20' ELSE '10' END as prodSpecCd "
		strSql = strSql & " , CASE WHEN i.itemdiv in ('06', '16') THEN '15' ELSE '03' END as orderMakingRd "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30') "
		strSql = strSql & "		or rtrim(ltrim(isNull(i.deliverfixday, ''))) <> '' "
		strSql = strSql & " 	or i.itemdiv not in ('01', '16', '07') "		'01 : 일반, 16 : 주문제작, 07 : 구매제한, 06 : 주문제작문구 상품 품절처리(LF미지원)
		strSql = strSql & " 	or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & "		or i.makerid in (SELECT makerid FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or i.itemid in (SELECT itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid WHERE mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "		or ((i.LimitYn='Y') and (i.LimitNo-i.LimitSold<="&CMAXLIMITSELL&")) "
		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < isNull(f.outmallstandardMargin, "& CMAXMARGIN &")) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & "	,CASE WHEN i.cate_large IN ('010') THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('020')THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('025')THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('030')THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('035')THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('055')THEN 'K730' "
		strSql = strSql & "	 WHEN i.cate_large IN ('060')THEN 'K730' "
		strSql = strSql & "	 WHEN i.cate_large IN ('050')THEN 'K731' "
		strSql = strSql & "	 WHEN i.cate_large IN ('045')THEN 'K720' "
		strSql = strSql & "	 WHEN i.cate_large IN ('040')THEN 'K720' "
		strSql = strSql & "	 ELSE '10'  END  lfbrandcode "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_lfmall_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lfmall_cate_mapping as am on am.tenCateLarge = i.cate_large and am.tenCateMid = i.cate_mid and am.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_etcmall.dbo.tbl_lfmall_noti_mapping as ni on ni.tenCateLarge = i.cate_large and ni.tenCateMid = i.cate_mid and ni.tenCateSmall = i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " LEFT JOIN db_partner.dbo.tbl_partner_addInfo as f on f.partnerid = '"& CMALLNAME &"' "
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & addSql
		strSql = strSql & " and m.lfmallGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CLfmallItem
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
				FoneItem.FoptionCnt			= rsget("optioncnt")
				FOneItem.FmaySoldOut    	= rsget("maySoldOut")
				FOneItem.FLfmallGoodNo		= rsget("lfmallGoodNo")
				FOneItem.FLfmallSellYn		= rsget("lfmallSellYn")
				FOneItem.FbasicImage		= "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicImage")
				FOneItem.FmainImage			= "http://webimage.10x10.co.kr/image/main/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage")
				FOneItem.FmainImage2		= "http://webimage.10x10.co.kr/image/main2/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("mainimage2")
				FOneItem.FbasicImage600		= rsget("basicimage600")
				FOneItem.FbasicImage1000	= rsget("basicimage1000")
				FOneItem.FbasicImage600str	= "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage600")
				FOneItem.FbasicImage1000str	= "http://webimage.10x10.co.kr/image/basic1000/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("basicimage1000")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName		= rsget("regImageName")
				FOneItem.FLfmallprice		= rsget("lfmallprice")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FOrderMaxNum		= rsget("orderMaxNum")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fitemsize			= rsget("itemsize")
				FOneItem.FItemsource		= rsget("itemsource")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
				FOneItem.Flfbrandcode 		= rsget("lfbrandcode")
				FOneItem.FItemKindCode 		= rsget("ItemKindCode")
				FOneItem.FSeasonCode		= rsget("SeasonCode")
				FOneItem.FColor1Code		= rsget("Color1Code")
				FOneItem.FdisplayProductName = rsget("displayProductName")
				FOneItem.FStandardCategoryId = rsget("StandardCategoryId")
				FOneItem.FprodSpecCd		= rsget("prodSpecCd")
				FOneItem.ForderMakingRd		= rsget("orderMakingRd")
				FOneItem.FAdultType			= rsget("adultType")
				FOneItem.FlfmallStatCD		= rsget("lfmallStatcd")
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
