<%
CONST CMAXMARGIN = 15			'' MaxMagin임..
CONST CMALLNAME = "gsshop"
CONST CMAXLIMITSELL = 5			'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CGSSHOPMARGIN = 12
CONST CUPJODLVVALID = True		''업체 조건배송 등록 가능여부
CONST COurCompanyCode = 1003890	'' 협력사코드
CONST COurRedId = "TBT"

Class CGSShopItem
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
	Public ForderComment
	Public FoptionCnt
	Public FbasicImage
	Public FmainImage
	Public FmainImage2
	Public Fsourcearea
	Public Fmakername
	Public FUsingHTML
	Public Fitemcontent
	Public FGsshopGoodNo
	Public FGsshopprice
	Public FGsshopSellYn
	Public FbasicimageNm
	Public FregImageName

	Public FregedOptCnt
	Public FaccFailCNT
	Public FlastErrStr
	Public Fdeliverytype
	Public FrequireMakeDay

	Public FinfoDiv
	Public Fsafetyyn
	Public FsafetyDiv
	Public FsafetyNum
	Public FmaySoldOut
	Public FIsNulltoTimeout

	Public FUserid
	Public FSocname
	Public FSocname_kor
	Public FDeliver_name
	Public FReturn_zipcode
	Public FReturn_address
	Public FReturn_address2
	Public FMaeipdiv
	Public FDeliveryCd
	Public FDeliveryAddrCd
	Public FBrandcd
	Public FDivname

	Public FIcnt
	Public FDivcode
	Public Fcdd_Name
	Public Fcdl_Name
	Public Fcdm_Name
	Public Fcds_Name

	Public FSafecode
	Public FSafecode_NAME
	Public FIsvat
	Public FIsvat_NAME
	Public FInfodiv1
	Public FInfodiv2
	Public FInfodiv3
	Public FInfodiv4
	Public FInfodiv5
	Public FInfodiv6


	Public FtenCDLName
	Public FtenCDMName
	Public FtenCDSName
	Public FDispNo
	Public FDispNm
	Public FDispLrgNm
	Public FDispMidNm
	Public FDispSmlNm
	Public Fdisptpcd
	Public FCateIsUsing
	Public FD_NAME

	Public FDispThnNm

	Public FItemoption
	Public Foptsellyn
	Public Foptlimityn
	Public Foptlimitno
	Public Foptlimitsold

	Public Fvatinclude
	Public FGSShopStatCd
	Public FOptNotMatch
	Public FAdultType

	Function getDispGubunNm()
		getDispGubunNm = getDisptpcdName
	End Function

	Public Function getDisptpcdName
		If (Fdisptpcd="B") Then
			getDisptpcdName = "<font color='blue'>전문</font>"
		Elseif (Fdisptpcd = "D") Then
			getDisptpcdName = "일반"
		Else
			getDisptpcdName = Fdisptpcd
		End if
	End Function


	'// 품절여부
	Public function IsSoldOut()
		ISsoldOut = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold<1))
	End Function

	'// 품절여부
	Public function IsSoldOutLimit5Sell()
		IsSoldOutLimit5Sell = (FSellyn<>"Y") or ((FLimitYn="Y") and (FLimitNo-FLimitSold <= CMAXLIMITSELL))
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

	'옵션 판매상태 수정
	Public Function isOptNotMatch()
		Dim strSql, arrRows, isOptionExists, tmpCnt
		Dim bufcnt, i, optLimit, optlimityn, isUsing, optsellyn, optNameDiff, forceExpired
		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_gsshop '"&CMALLNAME&"'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic
		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		tmpCnt = 0
		isOptNotMatch = "N"
		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				optNameDiff			= (ArrRows(12,i)=1)
				forceExpired		= (ArrRows(13,i)=1)
				If ((forceExpired) or (optNameDiff) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5)) Then
					tmpCnt = tmpCnt + 1
				End If
			Next

			If FOptionCnt = 1 AND tmpCnt = 1 AND i = 1 Then
				isOptNotMatch = "Y"
			ElseIf (FOptionCnt >= 1) AND (i = tmpCnt) Then
				isOptNotMatch = "Y"
			End If
		End If
	End Function

	Public Function MustPrice()
		Dim GetTenTenMargin, tmpPrice, sqlStr, specialPrice, vBigPrice, vSmallPrice, ownItemCnt
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
			If FGsshopprice = 0 Then
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					MustPrice = FSellCash
				End If
			Else
				If GetTenTenMargin < CMAXMARGIN Then
					MustPrice = Forgprice
				Else
					If (FSellCash < Round(FGsshopprice * 0.25, 0)) Then
						MustPrice = CStr(GetRaiseValue(Round(FGsshopprice * 0.25, 0)/10)*10)
					Else
						MustPrice = CStr(GetRaiseValue(FSellCash/10)*10)
					End If
				End If
			End If
		End If
	End Function

	public function GetGSLmtQty()
		CONST CLIMIT_SOLDOUT_NO = 5

		If (Flimityn="Y") then
			If (Flimitno - Flimitsold) < CLIMIT_SOLDOUT_NO Then
				GetGSLmtQty = 0
			Else
				GetGSLmtQty = Flimitno - Flimitsold - CLIMIT_SOLDOUT_NO
			End If
		Else
			GetGSLmtQty = 999
		End If
	End Function

	Public Function IsOptionSoldOut()
		CONST CLIMIT_SOLDOUT_NO = 5
		IsOptionSoldOut = false
		If (FItemOption = "0000") Then Exit Function
		IsOptionSoldOut = (Foptsellyn="N") or ((Foptlimityn="Y") and (Foptlimitno - Foptlimitsold <= CLIMIT_SOLDOUT_NO))
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

	'// GSShop 판매여부 반환
	Public Function getGSShopSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FsellYn="Y" and FisUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold > CMAXLIMITSELL) then
				getGSShopSellYn = "Y"
			Else
				getGSShopSellYn = "N"
			End If
		Else
			getGSShopSellYn = "N"
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
'		buf = "[텐바이텐]"&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐] 이라고 붙임
		buf = replace(FItemName,"'","")						'최초 상품명 앞에 [텐바이텐] 삭제

		If Left(FItemName, Len(Trim(FSocname_kor)) + 2) = "[" & FSocname_kor & "]" Then
		ElseIf (Left(FItemName, len(FSocname_kor)) <> FSocname_kor) Then
			buf = FSocname_kor & " " & Replace(FItemName,"'","")		'[텐바이텐] 문구 삭제 / 브랜드한글명 붙임 / 2020-07-30 위로 원복
		End If

		buf = replace(buf,"&#8211;","-")
		buf = replace(buf,"~","-")
		buf = replace(buf,"&","＆")
		buf = replace(buf,"<","[")
		buf = replace(buf,">","]")
		buf = replace(buf,"%","프로")
		buf = replace(buf,"+","%2B")
		buf = replace(buf,":","%3A")
		buf = replace(buf,"[무료배송]","")
		buf = replace(buf,"[무료 배송]","")
		getItemNameFormat = buf
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

	'상품분류별 안전인증
	Public Function getGSShopItemSafeInfoParam()
		Dim buf, strSql, regCertCnt, regSafetydiv
		Dim safeCertGbnCd, safeCertOrgCd, safeCertModelNm, safeCertNo, safeCertDt
		If FDivcode = "" Then			'상품분류를 지정안한 카테고리
			rw "상품분류를 지정해주세요"
			Exit Function
			response.end
		End If

		buf = ""
		If (FSafecode = "3") Then		'SafeCode가 3(비대상)이라면..
			buf = buf & "&safeCertGbnCd=0"		'(*)안전인증구분정보 | 0 : 해당사항없음, 1 : 전기안전인증, 2 : 공산품안전인증, 3 : 공산품자율안전확인번호, 4 : 전기용품자율안전확인
			buf = buf & "&safeCertOrgCd=0"		'(*)인증기관 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
			buf = buf & "&safeCertModelNm="		'인증모델명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
			buf = buf & "&safeCertNo="			'인증번호 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
			buf = buf & "&safeCertDt="			'인증일 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
'			buf = buf & "&safeCertFileNm="		'안전인증파일명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
		Else							'SafeCode가 1(필수,선택)이라면..
			If (Fsafetyyn) = "Y" AND (FSafecode = "1" OR FSafecode = "2") Then			'SafeCode가 1(필수,선택)이고 텐바이텐에 안전인증여부가 Y라면
				strSql = ""
				strSql = strSql & " SELECT COUNT(*) as cnt, safetydiv FROM db_item.dbo.tbl_safetycert_tenReg WHERE itemid = " &Fitemid& " GROUP BY safetydiv "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If not rsget.EOF Then
					regCertCnt = rsget("cnt")
					regSafetydiv = rsget("safetydiv")
				End If
				rsget.Close

				If regCertCnt > 0 AND (regSafetydiv = "30" OR regSafetydiv = "60" OR regSafetydiv = "90") Then
					If regSafetydiv = "30" Then
						safeCertGbnCd = "7"
						safeCertOrgCd = "701"
					ElseIf regSafetydiv = "60" Then
						safeCertGbnCd = "8"
						safeCertOrgCd = "801"
					ElseIf regSafetydiv = "90" Then
						safeCertGbnCd = "C"
						safeCertOrgCd = "C01"
					End If
					buf = buf & "&safeCertGbnCd="&safeCertGbnCd								'(*)안전인증구분정보 | 0 : 해당사항없음, 1 : 전기안전인증, 2 : 공산품안전인증, 3 : 공산품자율안전확인번호, 4 : 전기용품자율안전확인
					buf = buf & "&safeCertOrgCd="&safeCertOrgCd								'(*)인증기관 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertModelNm="											'인증모델명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertNo="												'인증번호 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertDt=" 												'인증일 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
	'				buf = buf & "&safeCertFileNm=Y"											'안전인증파일명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
				ElseIf regCertCnt > 0 AND (regSafetydiv <> "30" AND regSafetydiv <> "60" AND regSafetydiv <> "90") Then
					strSql = ""
					strSql = strSql & " EXEC [db_item].[dbo].[usp_API_GSShop_SafeInfo_Get] " & FItemid
					rsget.CursorLocation = adUseClient
					rsget.CursorType=adOpenStatic
					rsget.Locktype=adLockReadOnly
					rsget.Open strSql, dbget
					If Not(rsget.EOF or rsget.BOF) Then
						Do Until rsget.EOF
							buf = buf & "&safeCertGbnCd=" & rsget("safeCertGbnCd")			'(*)안전인증구분정보 | 0 : 해당사항없음, 1 : 전기안전인증, 2 : 공산품안전인증, 3 : 공산품자율안전확인번호, 4 : 전기용품자율안전확인
							buf = buf & "&safeCertOrgCd=" & rsget("safeCertOrgCd")			'(*)인증기관 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
							buf = buf & "&safeCertModelNm=" & rsget("safeCertModelNm")		'인증모델명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
							buf = buf & "&safeCertNo=" & rsget("safeCertNo")				'인증번호 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
							buf = buf & "&safeCertDt=" & rsget("safeCertDt") 				'인증일 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
			'				buf = buf & "&safeCertFileNm=Y"									'안전인증파일명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
							rsget.MoveNext
						Loop
					End If
					rsget.Close
				Else
					buf = buf & "&safeCertGbnCd=0"		'(*)안전인증구분정보 | 0 : 해당사항없음, 1 : 전기안전인증, 2 : 공산품안전인증, 3 : 공산품자율안전확인번호
					buf = buf & "&safeCertOrgCd=0"		'(*)인증기관 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertModelNm="		'인증모델명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertNo="			'인증번호 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
					buf = buf & "&safeCertDt="			'인증일 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
	'				buf = buf & "&safeCertFileNm="		'안전인증파일명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
				End If
			Else						'그 외의 것은 전부 해당없음 처리
				buf = buf & "&safeCertGbnCd=0"		'(*)안전인증구분정보 | 0 : 해당사항없음, 1 : 전기안전인증, 2 : 공산품안전인증, 3 : 공산품자율안전확인번호
				buf = buf & "&safeCertOrgCd=0"		'(*)인증기관 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
				buf = buf & "&safeCertModelNm="		'인증모델명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
				buf = buf & "&safeCertNo="			'인증번호 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
				buf = buf & "&safeCertDt="			'인증일 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
'				buf = buf & "&safeCertFileNm="		'안전인증파일명 | 안전인증 구분정보코드가 '0'일 경우 0 아닐경우는 입력
			End If
		End If
		getGSShopItemSafeInfoParam = buf
	End Function

	Public Function getGSCateParam()
		Dim strSql, bufcnt, cateKey, buf, cateGbn, isDefaultCate
		buf = ""
		strSql = ""
		strSql = strSql & " SELECT TOP 2 c.CateKey, c.cateGbn "
		strSql = strSql & " FROM db_item.dbo.tbl_gsshop_cate_mapping as m "
		strSql = strSql & " JOIN db_temp.dbo.tbl_gsshop_Category as c on m.CateKey = c.CateKey "
		strSql = strSql & " WHERE tenCateLarge='" & FtenCateLarge & "' "
		strSql = strSql & " and tenCateMid='" & FtenCateMid & "' "
		strSql = strSql & " and tenCateSmall='" & FtenCateSmall & "' "
		strSql = strSql & " ORDER BY c.cateGbn ASC " ''B : 브랜드 / D : 일반
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
'rw strSql
'response.end
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) then
			Do until rsget.EOF
				cateGbn = "S"			'기준매장 : S / 백화점매장 : D / 파트너스매장 : P / BP매장 : B
				isDefaultCate = "N"
				If rsget("cateGbn") = "B" Then
					cateGbn = "P"
					isDefaultCate = "Y"
				End If

			    cateKey  = rsget("CateKey")
				buf = buf & "&prdSectListSectid="&cateKey
				buf = buf & "&prdSectListSectGbn="&cateGbn
				buf = buf & "&prdSectListSectStdYn="&isDefaultCate
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSCateParam = bufcnt&"|_|"&buf
	End Function

	'협력사지급율/액 | 기본값 : 판매가*(1-0.13) // 마진12퍼
    Function getGSShopSuplyPrice()
		'getGSShopSuplyPrice = CLNG(FSellCash * (100-CGSSHOPMARGIN) / 100)
		getGSShopSuplyPrice = CLNG(MustPrice * (100-CGSSHOPMARGIN) / 100)
    End Function

	'상품 분류별 MDID 지정
	Public Function getMdIdMapping(divCode)
		Dim strSql
		strSql = ""
		strSql = strSql & " SELECT TOP 1 mdid "
		strSql = strSql & " FROM db_item.[dbo].[tbl_gsshop_mdid_mapping]  "
		strSql = strSql & " WHERE divcode = '"& divCode &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			getMdIdMapping = rsget("mdid")
		Else 
			getMdIdMapping = "80055"
		End If
		rsget.Close
	End Function

   ''주문제작 여부
    Public Function getzCostomMadeInd()
		Dim ordMnfcYn, ordMnfcTypCd, ordMnfcTermDdcnt, ordMnfcCntnt
		Dim buf
		If (Fitemdiv="06" or Fitemdiv="16" or FtenCateLarge="040") Then
			If Fitemdiv = "06" Then
				ordMnfcTypCd = "10"
				ordMnfcCntnt = "주문제작요청사항"
			ElseIf Fitemdiv="16" OR FtenCateLarge="040" Then
				ordMnfcTypCd = "20"
			End If

			If FtenCateLarge="040" Then
				ordMnfcTermDdcnt = 15
			ElseIf (FrequireMakeDay > 5) Then
				ordMnfcTermDdcnt = FrequireMakeDay
			ElseIf (FrequireMakeDay < 1) Then
				ordMnfcTermDdcnt = 5
			Else
				ordMnfcTermDdcnt = FrequireMakeDay + 1
			End If
			ordMnfcYn = "Y"
		Else
			ordMnfcYn = "N"
		End If

		buf = ""
		buf = buf & "&ordMnfcYn="&ordMnfcYn					'(*)주문제작여부
		buf = buf & "&ordMnfcTypCd="&ordMnfcTypCd			'(*)주문제작유형코드 | 주문제작여부가 'Y'일 경우 필수입력항목입니다.('N'일 때는 NULL) NULL : 해당없음, 10 : 맞춤제작, 20 : 주문후제작, 30 : 주문후수입
		buf = buf & "&ordMnfcCntnt="&ordMnfcCntnt			'(*)주문제작내용 | 주문제작유형이 10인 맞춤제작일 경우 필수입력항목입니다.
		buf = buf & "&ordMnfcTermDdcnt="&ordMnfcTermDdcnt	'(*)주문제작기간일수 | 주문제작여부가 'Y'일 경우 필수입력항목입니다.('N'일 때는 NULL)
		getzCostomMadeInd = buf
    End Function

	'//New 상품등록 파라메터 생성
	Public Function getGSShopItemNewRegParameter(v)
		Dim strRst
		Dim DeliverCd, DeliverAddrCd, brandcd
		'################################ 택배사/반품지 최초 확인 #################################
'2017-04-24 진영 수정..텐배던 업배던 CJ에 출고지 물류로..이유 : 따로따로 했을 때 묶음배송이 안 됨..
'		If (Fdeliverytype = "9") OR (Fdeliverytype = "7") OR (Fdeliverytype = "2") Then	'업체배송이라면
'			DeliverCd		= FDeliveryCd
'			DeliverAddrCd	= FDeliveryAddrCd
'			DeliverCd = "CJ"															'CJ택배
'			DeliverAddrCd = "0001"														'0001로 등록 협의 완료(도봉구 물류)
'		Else																			'텐배라면
'			DeliverCd = "CJ"															'CJ택배
'			DeliverAddrCd = "0001"														'0001로 등록 협의 완료(도봉구 물류)
'		End If

		DeliverCd = "HJ"															'한진택배
		DeliverAddrCd = "0001"														'0001로 등록 협의 완료(도봉구 물류)
		brandcd = "115985"
		'##########################################################################################

		'################################ 이미지 리스트 최초 호출 #################################
		Dim CallImage, CntImage, NmImage
		CallImage = getGSShopAddImageParam()
		CntImage = Split(CallImage, "|_|")(0)
		NmImage = Split(CallImage, "|_|")(1)
		'##########################################################################################

		'################################ 속성(옵션) 항목 최초 호출 ###############################
		Dim CallOpt, COptyn, CntOpt, NmOpt
		CallOpt = getGSShopOptionParam()
		COptyn = Split(CallOpt, "|_|")(0)
		CntOpt = Split(CallOpt, "|_|")(1)
		NmOpt = Split(CallOpt, "|_|")(2)
		'##########################################################################################

		'################################ 매장정보 항목 최초 호출 #################################
		Dim CallCate, CntCate, NmCate
		CallCate = getGSCateParam()
		CntCate = Split(CallCate, "|_|")(0)
		NmCate = Split(CallCate, "|_|")(1)
		'##########################################################################################

		'################################ 정부 고시 항목 최초 호출 ################################
		Dim CallInfoCd, CntInfoCd, NmInfoCd
		CallInfoCd = getGSShopItemInfoCdParam()
		CntInfoCd = Split(CallInfoCd, "|_|")(0)
		NmInfoCd = Split(CallInfoCd, "|_|")(1)
		'##########################################################################################
		'전송 구분 및 반복리스트 건수
		strRst = ""
		strRst = strRst & "regGbn=I"														'(*)등록구분 | I : 신규, U : 수정
		strRst = strRst & "&regId="&COurRedId												'(*)등록자	| 해당 협력사를 식별할수 있는 영문대문자 3자(예 : TBT)로 전송
		strRst = strRst & "&regSubjCd=SUP"													'(*)등록주체코드 | 엠디가 수정한 경우 : MD, 협력사가 수정한 경우 : SUP
		strRst = strRst & "&prdCntntListCnt="&CntImage										'(*)이미지리스트건수 | 상품이미지리스트 (prdCntntList) 반복횟수를 지정합니다.
'		strRst = strRst & "&prdDescdGnrlListCnt=0"											'(*)일반기술서리스트건수 | 내부상담원이 보는 텍스트기술서이며, 무조건 0 혹은 NULL로 셋팅
'		strRst = strRst & "&prdDescdHtmlItmListCnt="										'(*)이미지항목기술서리스트건수 | 도서몰전용필드 : 0 혹은 NULL로 셋팅
		strRst = strRst & "&attrPrdListCnt="&CntOpt											'(*)속성[옵션]리스트건수
		strRst = strRst & "&prdSectListCnt="&CntCate										'매장정보리스트건수
		strRst = strRst & "&prdGovPublsItmListCnt="&CntInfoCd								'(*)정부고시항목리스트건수 | 1건이상입력
		strRst = strRst & "&prdDescdHtmlImgListCnt=0"										'(*)상품상세기술서이미지건수 | 당사 이미지서버로 등록될 상세 기술수 이미지 건수 없는경우 null 또는 0
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		strRst = strRst & "&brandCd="&brandcd												'(*)브랜드코드 | 7152로 엑셀에 있던데?
		strRst = strRst & "&dlvPickMthodCd=3200"											'(*)배송수거방법코드 | 3200 : 직송(택배)-업체수거
		strRst = strRst & "&dlvsCoCd="&DeliverCd											'(*)택배사코드 | 배송택배사코드, 우선CJ로 등록
		strRst = strRst & "&saleStrDtm="&FormatDate(now(), "00000000000000")				'(*)판매시작일시
		strRst = strRst & "&saleEndDtm=29991231235959"										'(*)판매종료일시 | 상품을 중단(판매종료)하려면 중단시점의 판매종료일시를 입력합니다.
		strRst = strRst & "&cardUseLimitYn=N"												'카드사용제한여부
		strRst = strRst & "&baseAccmLimitYn=Y"												'(*)기본적립금제한여부 | 기본값 : Y
		strRst = strRst & "&selAccmApplyYn=Y"												'(*)선택적립금적용여부 | 기본값 : Y
		strRst = strRst & "&selAccRt="														'(*)선택적립율 | 기본값 : NULL
		strRst = strRst & "&immAccmDcLimitYn=Y"												'(*)즉시적립금할인제한여부 | 기본값 : Y
		strRst = strRst & "&immAccmDcRt="													'(*)즉시적립율 | 기본값 : NULL
		strRst = strRst & "&mnfcCoNm="&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)	'(*)제조사명
'		strRst = strRst & "&operMdId=80055"													'(*)운영mdid
		strRst = strRst & "&operMdId="& getMdIdMapping(FDivcode)							'(*)운영mdid
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)상품분류코드
		strRst = strRst & "&orgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지명 | 상품의 원산지명을 입력합니다. 예)미국,한국,중국 등
		strRst = strRst & "&prdNm="&DDotFormat(getItemNameFormat, 10)						'(*)상품명(송장) | 운송장에 입력되는 상품명입니다.
		strRst = strRst & "&regChanlGrpCd=GE"												'(*)등록채널그룹코드 | 판매할 상품을 채널그룹코드입니다. GE : 인터넷상품
		strRst = strRst & "&ordPrdTypCd=02"													'(*)주문상품유형코드 | 속성의 주문가능수량(재고)을 관리하는 구분코드입니다.02 : 상품속성별주문수량관리 01 : 상품별 주문수량관리
		'strRst = strRst & "&chrDlvYn="&CHKIIF(FSellcash>=30000, "N", "Y")					'(*)유료배송여부
		strRst = strRst & "&taxTypCd="&CHKIIF(FVatInclude="N","01","02")					'(*)세금유형코드 | 상품의 세금유형을 입력합니다. 01 : 면세, 02 : 과세, 03 : 영세
		strRst = strRst & "&dlvDtGuideCd=N"													'(*)배송일자안내코드 | 기본값 : N
		strRst = strRst & "&prdTypCd="&CHKIIF(COptyn = "Y","S","P")							'(*)상품유형코드 | 상품의 속성(옵션)이 구분을 입력합니다. P : 일반 (속성구분이 없는 경우) S : 속성 (색상/사이즈/형태/사이즈가 있는 경우) | P로 등록한 후에 S로변경하면 옵션추가가능//S->P로 일반상품 전환은 안 됨
		strRst = strRst & "&oboxCd="														'(*)합포장코드 | 기본값 : NULL
		strRst = strRst & "&chrDlvYn=Y"	'2016-06-21 19:28 김진영 수정..3만원 이상이면 N을 보냈으나 30000원 미만 2500원 코드 : 7237257 를 받음으로 무조건 Y로 전송..
		strRst = strRst & "&chrDlvcAmt=3000"												'유료배송비금액
		strRst = strRst & "&shipLimitAmt=50000"												'유료배송비면제기준금액
		strRst = strRst & "&exchRtpChrYn=Y"													'(*)교환반품유료여부 | 교환,반품시 배송비를 받을지 여부를 입력합니다.
		strRst = strRst & "&rtpAmt=6000"													'반품비 | 반품비를 사용할 금액을 입력 (교환반품유료여부를 Y로 전송해야 반영됨)
		strRst = strRst & "&exchAmt=6000"													'교환비 | 교환비를 사용할 금액을 입력 (교환반품유료여부를 Y로 전송해야 반영됨)
		strRst = strRst & "&chrDlvAddYn=N"													'(*)유료배송추가여부
		strRst = strRst & "&ilndDlvPsblYn=Y"												'도서지방배송가능여부
		strRst = strRst & "&jejuDlvPsblYn=Y"												'제주도배송가능여부
		strRst = strRst & "&dd3InDlvNoadmtRegonYn=N"										'3일내배송불가지역여부
		strRst = strRst & "&ilndChrDlvYn=Y"													'도서지방유료배송여부 | 직송-택배일경우만 추가유료배송
		strRst = strRst & "&ilndChrDlvcAmt=3000"											'도서지방유료배송비	도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndExchRtpChrYn=Y"												'도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndRtpAmt=6000"												'도서지방반품비 | 도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndExchAmt=6000"												'도서지방교환비 | 도서지방 추가배송비 유료일 경우
		strRst = strRst & "&jejuChrDlvYn=Y"													'제주도유료배송여부 | 직송-택배일경우만 추가유료배송 가능
		strRst = strRst & "&jejuChrDlvcAmt=3000"											'제주도유료배송비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuExchRtpChrYn=Y"												'제주도교환반품유료여부	제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuRtpAmt=6000"												'제주도반품비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuExchAmt=6000"												'제주도교환비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&prdGbnCd=00"													'(*)상품구분코드 | 일반상품,사은품,경품을 구분하는 값입니다.00 : 일반상품, 02 : 사은품-업체제공
		strRst = strRst & "&bundlDlvCd=A01"													'(*)묶음배송코드 | 묶음배송 가능/불가능을 지정하는 값입니다. A01 : 가능, A02 : 불가능
		strRst = strRst & "&modelNo="														''''모델번호
		strRst = strRst & "&cpnApplyTypCd=09"												'(*)쿠폰적용유형코드 | 할인쿠폰 적용 또는 제한하는 값입니다. 00 : 쿠폰허용, 03 : 상품쿠폰만 적용, 09 : 쿠폰제한
		If Fitemdiv="06" OR Fitemdiv="16" OR FtenCateLarge="040" Then
			strRst = strRst & "&openAftRtpNoadmtYn=Y"										'(*)개봉후반품불가여부 | 기본값 : Y,N	(주문제작은 Y // 아닌건 N)
		Else
			strRst = strRst & "&openAftRtpNoadmtYn=N"										'(*)개봉후반품불가여부 | 기본값 : Y,N	(주문제작은 Y // 아닌건 N)
		End If
		strRst = strRst & "&istTypCd="														'(*)입고유형코드 | 기본값 : NULL
'		strRst = strRst & "&chrDlvcCd=7237257"												'(*)유료배송비코드
		strRst = strRst & "&prdRelspAddrCd="&DeliverAddrCd									'(*)상품출고지주소코드
		strRst = strRst & "&prdRetpAddrCd="&DeliverAddrCd									'(*)상품반송지주소코드
		strRst = strRst & "&separOrdNoadmtYn=N"												'(*)단독주문불가여부 | 기본값 : N
		strRst = strRst & "&gftTypCd=00"													'(*)사은품유형코드 | 00 : 판매상품, 02 : 사은품-업체제공
		strRst = strRst & "&prchTypCd=03"													'(*)매입유형코드 | 03 : 수수료매입
		strRst = strRst & "&zrwonSaleYn=N"													'(*)0원판매여부
		strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)하위협력사코드 | 협력사코드와 동일하게 입력
		strRst = strRst & getzCostomMadeInd													'(*)주문제작여부 및 항목 함수호출
		strRst = strRst & "&attrTypExposCd=L"												'(*)속성유형노출코드 | L : 리스트
		strRst = strRst & "&adultCertYn="&Chkiif(IsAdultItem() = "Y", "Y", "N")&""			'(*)성인인증여부	(우선은N으로)
		strRst = strRst & "&barcdNo="														'바코드번호
		strRst = strRst & "&apntDlvDlvsCoCd="												'(*)지정배송택배사코드 | 기본값 : NULL
		strRst = strRst & "&apntPickDlvsCoCd="												'(*)지정수거택배사코드 | 기본값 : NULL
		strRst = strRst & "&gnuinYn=N"														'(*)정품여부 | 기본값 : N
		strRst = strRst & "&frmlesPrdTypCd=N"												'(*)무형상품유형코드 | 기본값 : N
		strRst = strRst & "&rsrvSalePrdYn=N"												'예약판매여부
		'이중이상의 옵션이라면 옵션타입명을 선택으로 고정시키고 CJMall과 같이 2~3중 옵션을 나누지 않고 하나의 옵션에 다 넣게..
		strRst = strRst & "&attrTypNm1="&CHKIIF(COptyn = "Y","선택","")						'속성유형명1 | 속성정보의 속성값 타이틀을 변경하고자 할때 쓰이는 컬럼. 빈값으로 보낼경우, 색상 으로 표시된다.
		strRst = strRst & "&attrTypNm2="													'속성유형명2 | 속성정보의 타이틀을 변경하고자 할때 쓰이는 컬럼 빈값으로 보낼경우, 사이즈 로 표시된다.
		strRst = strRst & "&attrTypNm3="													'속성유형명3 | 속성정보의 타이틀을 변경하고자 할때 쓰이는 컬럼 빈값으로 보낼경우, 스타일 으로 표시된다.
		strRst = strRst & "&attrTypNm4="													'속성유형명4 | 속성정보의 타이틀을 변경하고자 할때 쓰이는 컬럼 빈값으로 보낼경우, 사은품 으로 표시된다.
		strRst = strRst & "&attrSaleEndStModYn="											'속성판매종료상태수정설정 | 속성구분(S) 상품판매상태를 변경할 때 사용하는 항목으로, 상품마스터 종료 및 해제 시 속성상품의 상태도 함께 종료 및 해제하려면 Y, 상품마스터와 속성 별도로 상태변경 동작 시엔 N
		'상품확장(prdAddInfo)
		strRst = strRst & "&prdBaseCmposCntnt="&Trim(chrbyte(getItemNameFormat,56,"Y"))		'(*)상품기본구성내용 | 상품명과 동일하게 입력
		strRst = strRst & "&orgprdPkgCnt=1"													'(*)본품포장갯수
		strRst = strRst & "&prdAddCmposCntnt="												'상품추가구성내용
		strRst = strRst & "&addCmposPkgCnt="												'추가구성포장개수
		strRst = strRst & "&addCmposOrgpNm="												'추가구성원산지명
		strRst = strRst & "&addCmposMnfcCoNm="												'추가구성제조사명
		strRst = strRst & "&prdGftCmposCntnt="												'상품사은품구성내용
		strRst = strRst & "&gftPkgCnt="														'사은품포장개수
		strRst = strRst & "&gftCmposOrgpNm="												'사은품구성원산지명
		strRst = strRst & "&gftCmposMnfcCoNm="												'사은품구성제조사명
		strRst = strRst & "&prdUnitValCd40=A01"												'(*)상품무게정보 | A01 : 2.5kg미만, A02 : 2.5kg이상 ~ 5kg미만, A03 : 5kg이상 ~ 20kg미만, A04 : 30kg이상, A05 : 20kg이상 ~ 30kg미만
		strRst = strRst & "&prdUnitValCd20=B01"												'(*)상품길이정보 | B01 : 80cm미만, B02 : 80cm이상 ~ 120cm미만, B03 : 120cm이상 ~ 160cm미만, B04 : 160cm이상
		'상품예정정보(prdSchdInfo)
'		strRst = strRst & "&prdSchdInfoRsrvOrdStrDt="										'예약주문가능시작일시 | 상품기본의 예약판매여부가 'Y'일 경우만 필수입력항목입니다.
'		strRst = strRst & "&prdSchdInfoRsrvOrdEndDt="										'예약주문가능종료일시 | 상품기본의 예약판매여부가 'Y'일 경우만 필수입력항목입니다.
'		strRst = strRst & "&prdSchdInfoRsrvRelsStrDt="										'예약출고시작일시 | 상품기본의 예약판매여부가 'Y'일 경우만 필수입력항목입니다.
'		strRst = strRst & "&prdSchdInfoRsrvRelsEndDt="										'예약출고종료일시 | 상품기본의 예약판매여부가 'Y'일 경우만 필수입력항목입니다.
		'상품가격(prdPrc)
		strRst = strRst & "&prdPrcValidStrDtm="&FormatDate(now(), "00000000000000")			'(*)유효시작일시
		strRst = strRst & "&prdPrcValidEndDtm=29991231235959"								'(*)유효종료일시
		strRst = strRst & "&prdPrcSalePrc="&MustPrice										'(*)판매가격
'		strRst = strRst & "&prdPrcPrchPrc="													'(SYS)매입가격 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
		strRst = strRst & "&prdPrcSupGivRtamtCd=01"											'(*)협력사지급율/액코드 | 01 : 액
		strRst = strRst & "&prdPrcSupGivRtamt="&getGSShopSuplyPrice()						'(*)협력사지급율/액 | 기본값 : 판매가*(1-0.12)
		'노출상품명(prdNmChg)
		strRst = strRst & "&prdNmChgValidStrDtm="&FormatDate(now(), "00000000000000")		'(*)유효시작일시
		strRst = strRst & "&prdNmChgValidEndDtm=29991231235959"								'(*)유효종료일시
		strRst = strRst & "&prdNmChgExposPrdNm=" & Trim(chrbyte(getItemNameFormat,56,"Y"))					'(*)노출상품명 | GSShop노출상품명
		'상품이미지(prdCntntList)
		strRst = strRst & NmImage

		If v = 1 Then
			'상품상세기술서(prdDescdHtml)
			strRst = strRst & getGSShopItemContParam()
		Else
			strRst = strRst & "&prdDescdHtmlDescdExplnCntnt=" & Server.URLEncode("<div align=""center""><p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		End If

		'상품기본-속성
		strRst = strRst & NmOpt
		'상품전시매장(prdSectList)
		strRst = strRst & NmCate															'(*)매장정보아이디
		'안전인증(prdSafeCertInfo)
		strRst = strRst & getGSShopItemSafeInfoParam()
		'정부고시항목(prdGovPublsItmList)
		strRst = strRst & NmInfoCd
		'rw strRst
		'response.end
		getGSShopItemNewRegParameter = strRst
	End Function

	'상품품목정보
	public function getGSShopItemInfoCdParam()
		Dim strSql, bufcnt, buf, certNum
		Dim mallinfoCd,infoContent,infotype, infocd, mallinfodiv
		' strSql = ""
		' strSql = strSql & " SELECT TOP 1 certNum "
		' strSql = strSql & " FROM db_item.dbo.tbl_safetycert_tenReg "
		' strSql = strSql & " WHERE itemid='"&FItemID&"' "
		' rsget.CursorLocation = adUseClient
		' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		' If Not(rsget.EOF or rsget.BOF) then
		' 	certNum = rsget("certNum")
		' End If
		' rsget.Close

		' buf = ""
		' strSql = ""
		' strSql = strSql & " SELECT TOP 100 M.* , " & vbcrlf
		' strSql = strSql & "		CASE " & vbcrlf
        ' strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') and isnull(IC.safetyNum, '') = '' THEN '"&certNum&"' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') and isnull(IC.safetyNum, '') <> '' THEN IC.safetyNum " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00001') THEN '제품소재참고' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00002') THEN '가입조건참고' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00003') THEN '주요사항참고' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00004') THEN '해당없음' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00005') THEN '가공식품' " & vbcrlf
		' strSql = strSql & "			WHEN (M.infoCd='00006') THEN '건강기능식품' " & vbcrlf
		' strSql = strSql & "			WHEN c.infotype='J' and F.chkDiv='N' THEN '해당없음' " & vbcrlf
		' strSql = strSql & "			WHEN c.infotype='P' AND c.infoCd <> '22009' THEN '텐바이텐 고객행복센터 1644-6035' " & vbcrlf
		' strSql = strSql & "			WHEN LEN(F.infocontent) <= 1 THEN F.infocontent + ' 포함' " & vbcrlf
		' strSql = strSql & "		ELSE convert(varchar(500),F.infocontent) " & vbcrlf
		' strSql = strSql & " END AS infocontent " & vbcrlf
		' strSql = strSql & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		' strSql = strSql & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		' strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		' strSql = strSql & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemID&"' " & vbcrlf
		' strSql = strSql & " WHERE M.mallid = '"&CMALLNAME&"' and IC.itemid='"&FItemID&"' "
		' rsget.CursorLocation = adUseClient
		' rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		' bufcnt = rsget.RecordCount
		' If Not(rsget.EOF or rsget.BOF) then
		' 	Do until rsget.EOF
		' 	    mallinfoCd  = rsget("mallinfoCd")
		' 	    infoContent = rsget("infoContent")
		' 		infocd		= rsget("infocd")
		' 		mallinfodiv = rsget("mallinfodiv")
		' 		If isnull(infoContent) Then
		' 			infoContent = ""
		' 		End If

		' 		infoContent = replace(infoContent, "&", "＆")
		' 		infoContent = replace(infoContent, "?", "？")
		' 		infoContent = replace(infoContent, "%", "％")

		' 		buf = buf & "&govPublsItmCd="&mallinfoCd						'(*)정부고시항목값
		' 		buf = buf & "&govPublsItmCntnt="&infoContent					'(*)정부고시항목내용
		' 		rsget.MoveNext
		' 	Loop
		' End If
		' rsget.Close

		buf = ""
		strSql = ""
		strSql = strSql & " EXEC db_item.dbo.usp_API_GSShop_InfoCodeMap_Get " & FItemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			Do Until rsget.EOF
			    mallinfoCd  = rsget("mallinfoCd")
			    infoContent = rsget("infoContent")
				If isnull(infoContent) Then
					infoContent = ""
				End If

				infoContent = replace(infoContent, "&", "＆")
				infoContent = replace(infoContent, "?", "？")
				infoContent = replace(infoContent, "%", "％")

				buf = buf & "&govPublsItmCd="&mallinfoCd						'(*)정부고시항목값
				buf = buf & "&govPublsItmCntnt="&infoContent					'(*)정부고시항목내용
				rsget.MoveNext
			Loop
		End If
		rsget.Close
		getGSShopItemInfoCdParam = bufcnt&"|_|"&buf
	End Function

	'//상품설명 파라메터 생성
	Public Function getGSShopItemContParam()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		strRst = strRst & Server.URLEncode("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & Server.URLEncode("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & Server.URLEncode("<tr>")
		strRst = strRst & Server.URLEncode("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""상품명"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & Server.URLEncode("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '돋움', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & Server.URLEncode("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & getItemNameFormat
		strRst = strRst & Server.URLEncode("</p>")
		strRst = strRst & Server.URLEncode("</td>")
		strRst = strRst & Server.URLEncode("</tr>")
		strRst = strRst & Server.URLEncode("</table>")
		strRst = strRst & Server.URLEncode("</div>")

		If ForderComment <> "" Then
			strRst = strRst & Server.URLEncode("- 주문시 유의사항 :<br>" & Fordercomment & "<br>")
		End If

		Fitemcontent = replace(Fitemcontent,"&nbsp;"," ")
		Fitemcontent = replace(Fitemcontent,"&nbsp"," ")
		Fitemcontent = replace(Fitemcontent,"&"," ")
		Fitemcontent = replace(Fitemcontent,chr(13)," ")
		Fitemcontent = replace(Fitemcontent,chr(10)," ")
		Fitemcontent = replace(Fitemcontent,chr(9)," ")

		Select Case FUsingHTML
			Case "Y"
				'strRst = strRst & Server.URLEncode(Fitemcontent & "<br>")
				strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case "H"
				'strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case Else
				'strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
				strRst = strRst & nl2br(ReplaceBracket(Fitemcontent)) & "<br>"
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
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_gsshop.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		'(*)기술서설명내용 | GSShop에 노출되는 HTML기술서		prdDescdHtmlDescdExplnCntnt
		getGSShopItemContParam = "&prdDescdHtmlDescdExplnCntnt=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = Server.URLEncode("<div align=""center"">"& strtextVal & "</div>")
			getGSShopItemContParam = "&prdDescdHtmlDescdExplnCntnt=" & strRst
		End If
		rsget.Close
	End Function

	'//상품설명 파라메터 생성
	Public Function getGSShopItemContParamEucKR()
		Dim strRst, strSQL, strtextVal
		strRst = Server.URLEncode("<div align=""center"">")
		strRst = strRst & Server.URLEncode("<p><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_gsshop.jpg""></p><br>")
		strRst = strRst & Server.URLEncode("<div style=""width:100%; max-width:700px; margin:0; padding:0; margin-bottom:14px; padding-bottom:6px; background:url(http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_namebg.png) left bottom no-repeat;"">")
		strRst = strRst & Server.URLEncode("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
		strRst = strRst & Server.URLEncode("<tr>")
		strRst = strRst & Server.URLEncode("<th style=""vertical-align:middle; width:73px; height:42px; text-align:center; margin:0; padding:3px 0 0 0;""><img src=""http://fiximage.10x10.co.kr/web2008/etc/gs_pdt_nametit.png"" alt=""상품명"" style=""vertical-align:top; display:inline;""/></th>")
		strRst = strRst & Server.URLEncode("<td style=""width:627px; vertical-align:middle; text-align:left; font-size:14px; line-height:1.2; color:#000; font-weight:bold; font-family:dotum, dotumche, '돋움', sans-serif; margin:0; padding:4px 0 0 0;"">")
		strRst = strRst & Server.URLEncode("<p style=""letter-spacing:-0.03em; margin:0; padding:12px 10px;"">")
		strRst = strRst & Server.URLEncode(getItemNameFormat)
		strRst = strRst & Server.URLEncode("</p>")
		strRst = strRst & Server.URLEncode("</td>")
		strRst = strRst & Server.URLEncode("</tr>")
		strRst = strRst & Server.URLEncode("</table>")
		strRst = strRst & Server.URLEncode("</div>")

		If ForderComment <> "" Then
			strRst = strRst & Server.URLEncode("- 주문시 유의사항 :<br>" & Fordercomment & "<br>")
		End If

		Fitemcontent = replace(Fitemcontent,"&nbsp;"," ")
		Fitemcontent = replace(Fitemcontent,"&nbsp"," ")
		Fitemcontent = replace(Fitemcontent,"&"," ")
		Fitemcontent = replace(Fitemcontent,chr(13)," ")
		Fitemcontent = replace(Fitemcontent,chr(10)," ")
		Fitemcontent = replace(Fitemcontent,chr(9)," ")

		Select Case FUsingHTML
			Case "Y"
				strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				'strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case "H"
				strRst = strRst & Server.URLEncode(nl2br(Fitemcontent) & "<br>")
				'strRst = strRst & nl2br(Fitemcontent) & "<br>"
			Case Else
				strRst = strRst & Server.URLEncode(nl2br(ReplaceBracket(Fitemcontent)) & "<br>")
				'strRst = strRst & nl2br(ReplaceBracket(Fitemcontent)) & "<br>"
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
		strRst = strRst & Server.URLEncode("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_gsshop.jpg"">")
		strRst = strRst & Server.URLEncode("</div>")
		'(*)기술서설명내용 | GSShop에 노출되는 HTML기술서		prdDescdHtmlDescdExplnCntnt
		getGSShopItemContParamEucKR = "&prdDescdHtmlDescdExplnCntnt=" & strRst

		strSQL = ""
		strSQL = strSQL & " SELECT itemid, mallid, linkgbn, textVal " & VBCRLF
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_etcLink " & VBCRLF
		strSQL = strSQL & " WHERE mallid in ('','"&CMALLNAME&"') and linkgbn = 'contents' and itemid = '"&Fitemid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) Then
			strtextVal = rsget("textVal")
			strRst = Server.URLEncode("<div align=""center"">"& strtextVal & "</div>")
			getGSShopItemContParamEucKR = "&prdDescdHtmlDescdExplnCntnt=" & strRst
		End If
		rsget.Close
If (session("ssBctID") = "kjy8517") Then
	' rw FUsingHTML & "!!!!!!!!!!!!!!!!!!!!!!"
	' rw getGSShopItemContParamEucKR
end if
	End Function

	'상품 이미지
	Public Function getGSShopAddImageParam()
		Dim strRstCnt, strRst, strSQL, i
		'최초 빅사이즈 이미지
		'(*)이미지url | 가장 큰 이미지의 URL 입력하면 자동리사이징 처리됨 (GSShop 최대이미지 : 550x550)
		strRst = "&prdCntntListCntntUrlNm="&Server.URLEncode(FbasicImage)
		strRstCnt = 1
		'미니사이즈  이미지
		strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType=adOpenStatic
		rsget.Locktype=adLockReadOnly
		rsget.Open strSQL, dbget
		If Not(rsget.EOF or rsget.BOF) Then
			For i=1 to rsget.RecordCount
				If rsget("imgType") = "0" Then
					strRst = strRst & "&prdCntntListCntntUrlNm=" & Server.URLEncode("http://webimage.10x10.co.kr/image/add" & rsget("gubun") & "/" & GetImageSubFolderByItemid(Fitemid) & "/" & rsget("addimage_400"))
					strRstCnt = strRstCnt + 1
				End If
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getGSShopAddImageParam = strRstCnt&"|_|"&strRst
	End Function

	'옵션 파라메터 생성
	Public Function getGSShopOptionParam()
		Dim strSql, strRst, itemSu, validSellno, optionname, fixday, optaddprice
		Dim ret, bufcnt, optyn, i
		ret = ""
		strSql = ""
		strSql = strSql & " SELECT T.* "
		strSql = strSql & " INTO #T1 "
		strSql = strSql & "	FROM ( "
		strSql = strSql & " 	SELECT i.itemid, i.limitno ,i.limitsold, o.itemoption, convert(varchar(96),o.optionname) as optionname" & VBCRLF
		strSql = strSql & " 	, o.optlimitno, o.optlimitsold " & VBCRLF
		strSql = strSql & " 	,Case When (o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold > 5) Then 'Y' " & VBCRLF
		strSql = strSql & " 		  When (o.optlimityn = 'Y') and (o.optlimitno - o.optlimitsold <= 5) Then 'N' " & VBCRLF
		strSql = strSql & " 		  When (isnull(o.itemid, '') = '') Then 'Y' " & VBCRLF		'-- 옵션 없는 단품이나 Y로 처리
		strSql = strSql & " 	else o.optsellyn end as optsellyn " & VBCRLF
		strSql = strSql & " 	, o.optlimityn, i.deliverfixday, o.optaddprice " & VBCRLF
		strSql = strSql & " 	,DATALENGTH(o.optionname) as optnmLen, isnull(r.outmallOptCode,'') as outmallOptCode" & VBCRLF
		strSql = strSql & " 	FROM db_item.dbo.tbl_item as i " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.[dbo].tbl_item_option as o on i.itemid = o.itemid and o.isusing = 'Y' and o.optaddprice=0 " & VBCRLF
		strSql = strSql & " 	LEFT JOIN db_item.[dbo].tbl_outmall_regedoption as r on i.itemid = r.itemid and o.itemoption = r.itemoption and r.mallid = '"&CMALLNAME&"' " & VBCRLF
		strSql = strSql & " 	WHERE i.itemid = "&Fitemid
		strSql = strSql & " ) AS T " & VBCRLF
		strSql = strSql & " WHERE T.optsellyn = 'Y' "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " SELECT t.itemoption "
		strSql = strSql & " INTO #T2 "
		strSql = strSql & " FROM #T1 as t "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_outmall_regedoption as r on  r.outmallOptName = t.optionname  "
		strSql = strSql & " WHERE r.mallid = '"&CMALLNAME&"' "
		strSql = strSql & " and r.itemid = " & Fitemid
		strSql = strSql & " and t.outmallOptCode = '' "
		strSql = strSql & " GROUP BY t.itemoption "
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " SELECT * FROM #T1 "
		strSql = strSql & " WHERE itemoption not in ( "
		strSql = strSql & " 	SELECT itemoption FROM #T2 "
		strSql = strSql & " ) "
		strSql = strSql & " ORDER BY optaddprice, itemoption "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			For i = 1 to rsget.RecordCount
				If rsget.RecordCount = 1 AND IsNull(rsget("itemoption")) Then  ''단일상품
				    optionname="공통"
				    FItemoption = "0000"
					itemSu = GetGSLmtQty()
					optyn	= "N"
					bufcnt	= 1
				Else
					FItemoption		= rsget("itemoption")
					optionname 		= db2Html(rsget("optionname"))
					Foptsellyn 		= rsget("optsellyn")
					Foptlimityn 	= rsget("optlimityn")
					Foptlimitno 	= rsget("optlimitno")
					Foptlimitsold 	= rsget("optlimitsold")
					itemSu = getOptionLimitNo()
					optyn	= "Y"
					If rsget("optnmLen") > 80 Then
					    optionname=DdotFormat(optionname,40)
					End If
				End If

				'2016-01-22 14:38 김진영 URL 인코딩한 옵션명으로
				optionname = replace(optionname,"&","＆")
				optionname = replace(optionname,"%","프로")
				optionname = replace(optionname,"+","%2B")
				optionname = replace(optionname,","," ")

				ret = ret & "&attrPrdListSupAttrPrdCd="&FItemoption							'Null이라더니 Null로 전송하면 안 됨'(SYS)협력사속성상품코드 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
'				ret = ret & "&attrPrdListAttrPrdCd="&Chkiif(rsget("outmallOptCode") <> "", rsget("outmallOptCode"), "")	'(*)(SYS)GS속성상품코드 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
				ret = ret & "&attrPrdListAttrValCd1=00000"									'(*)속성값코드1 | 기본값 : 00000
				ret = ret & "&attrPrdListAttrValCd2=00000"									'(*)속성값코드2 | 기본값 : 00000
				ret = ret & "&attrPrdListAttrValCd3=00000"									'(*)속성값코드3 | 기본값 : 00000
				ret = ret & "&attrPrdListAttrValCd4=00000"									'(*)속성값코드4 | 기본값 : 00000
				ret = ret & "&attrPrdListSaleStrDtm="&FormatDate(now(), "00000000000000")	'(*)판매시작일시
				ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)판매종료일시
				ret = ret & "&attrPrdListModelNo="											'모델번호
				ret = ret & "&attrPrdListAttrVal1="&optionname								'(*)속성값1 | 상품기본의 상품유형코드가 P일 경우 : '공통' 으로 넣으며 속성갯수 1개, S일 경우 : 색상값 없으면 'None', 있으면 값입력하고 속성갯수는 n개
				ret = ret & "&attrPrdListAttrVal2="&ChkIIF(optyn="Y","None","공통")			'(*)속성값2 | 상품기본의 상품유형코드가 P일 경우 : '공통' 으로 넣으며 속성갯수 1개, S일 경우 : 사이즈값 없으면 'None', 있으면 값입력하고 속성갯수는 n개
				ret = ret & "&attrPrdListAttrVal3="&ChkIIF(optyn="Y","None","공통")			'(*)속성값3 | 상품기본의 상품유형코드가 P일 경우 : '공통' 으로 넣으며 속성갯수 1개, S일 경우 : 스타일값 없으면 'None', 있으면 값입력하고 속성갯수는 n개
				ret = ret & "&attrPrdListAttrVal4="&ChkIIF(optyn="Y","None","공통")			'(*)속성값4 | 상품기본의 상품유형코드가 P일 경우 : '공통' 으로 넣으며 속성갯수 1개, S일 경우 : 사은품값 없으면 'None', 있으면 값입력하고 속성갯수는 n개, (본품과 합포장해서 주는 사은품)
'				ret = ret & "&attrPrdListArsAttrVal1="										'(*)자동주문속성값1 | 기본값 : NULL
'				ret = ret & "&attrPrdListArsAttrVal2="										'(*)자동주문속성값2 | 기본값 : NULL
'				ret = ret & "&attrPrdListArsAttrVal3="										'(*)자동주문속성값3 | 기본값 : NULL
'				ret = ret & "&attrPrdListArsAttrVal4="										'(*)자동주문속성값4 | 기본값 : NULL
'				ret = ret & "&attrPrdListAttrPkgCnt="										'(*)속성포장개수 | 기본값 : NULL
				ret = ret & "&attrPrdListAttrCmposCntnt="									'(*)속성구성정보 | 기본값 : NULL
				ret = ret & "&attrPrdListOrgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지명
				ret = ret & "&attrPrdListMnfcCoNm="&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)	'(*)제조사명
				ret = ret & "&attrPrdListSafeStockQty=5"									'(*)안전재고수량 | 안전재고이하로 수량이 내려가면 담당MD에게 알림을 함
				ret = ret & "&attrPrdListTempoutYn=N"										'(*)일시품절여부 | 기본값 : N
'				ret = ret & "&attrPrdListTempoutDtm="										'일시품절일시
				ret = ret & "&attrPrdListChanlGrpCd=AZ"										'(*)채널그룹코드 | AZ : DM외(DM을 제외한 나머지 채널)
				ret = ret & "&attrPrdListOrdPsblQty="&itemSu								'(*)주문가능수량
				rsget.MoveNext
			Next
		End If
		rsget.Close
		getGSShopOptionParam = optyn&"|_|"&bufcnt&"|_|"&ret
	End Function

	Public Function getGSShopImageEditParameter()
		Dim strRst
		'################################ 이미지 리스트 최초 호출 #################################
		Dim CallImage, CntImage, NmImage
		CallImage = getGSShopAddImageParam()
		CntImage = Split(CallImage, "|_|")(0)
		NmImage = Split(CallImage, "|_|")(1)
		'##########################################################################################
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=I"														'(*)수정구분 I : 이미지 수정
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&prdCntntListCnt="&CntImage										'(*)이미지리스트건수 | 상품이미지리스트 (prdCntntList) 반복횟수를 지정합니다.
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		'상품이미지(prdCntntList)
		strRst = strRst & NmImage
		getGSShopImageEditParameter = strRst
	End Function

	Public Function getGSShopSafeCertEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=C"														'(*)수정구분 C : 안전인증정보
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		'안전인증(prdSafeCertInfo)
		strRst = strRst & getGSShopItemSafeInfoParam()
		getGSShopSafeCertEditParameter = strRst
	End Function

	Public Function getGSShopItemEditParameter()
		Dim strRst
		Dim DeliverCd, DeliverAddrCd
		DeliverCd = "HJ"																	'한진택배
		DeliverAddrCd = "0001"														'0001로 등록 협의 완료(도봉구 물류)

		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=A"														'(*)수정구분 A: 상품정보
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&regSubjCd=SUP"													'(*)등록주체코드 | 엠디가 수정한 경우 : MD, 협력사가 수정한 경우 : SUP
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		strRst = strRst & "&dlvsCoCd="&DeliverCd											'(*)택배사코드 | 배송택배사코드, 우선CJ로 등록
		strRst = strRst & "&orgpNm="&chkIIF(trim(Fsourcearea)="" or isNull(Fsourcearea),"상품설명 참조",Fsourcearea)	'(*)원산지명 | 상품의 원산지명을 입력합니다. 예)미국,한국,중국 등
		strRst = strRst & "&chrDlvYn=Y"														'(*)유료배송여부 | 변경한 경우 보냄
		strRst = strRst & "&chrDlvcAmt=3000"												'유료배송비금액
		strRst = strRst & "&shipLimitAmt=50000"												'유료배송비면제기준금액
		strRst = strRst & "&exchRtpChrYn=Y"													'(*)교환반품유료여부 | 변경한 경우 보냄
		strRst = strRst & "&rtpAmt=6000"													'반품비 | 반품비를 사용할 금액을 입력 (교환반품유료여부를 Y로 전송해야 반영됨)
		strRst = strRst & "&exchAmt=6000"													'교환비 | 교환비를 사용할 금액을 입력 (교환반품유료여부를 Y로 전송해야 반영됨)
		strRst = strRst & "&chrDlvAddYn=N"													'(*)유료배송추가여부 | 변경한 경우 보냄
		strRst = strRst & "&ilndDlvPsblYn=Y"												'도서지방배송가능여부
		strRst = strRst & "&jejuDlvPsblYn=Y"												'제주도배송가능여부
		strRst = strRst & "&dd3InDlvNoadmtRegonYn=N"										'3일내배송불가지역여부
		strRst = strRst & "&ilndChrDlvYn=Y"													'도서지방유료배송여부 | 직송-택배일경우만 추가유료배송
		strRst = strRst & "&ilndChrDlvcAmt=3000"											'도서지방유료배송비	도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndExchRtpChrYn=Y"												'도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndRtpAmt=6000"												'도서지방반품비 | 도서지방 추가배송비 유료일 경우
		strRst = strRst & "&ilndExchAmt=6000"												'도서지방교환비 | 도서지방 추가배송비 유료일 경우
		strRst = strRst & "&jejuChrDlvYn=Y"													'제주도유료배송여부 | 직송-택배일경우만 추가유료배송 가능
		strRst = strRst & "&jejuChrDlvcAmt=3000"											'제주도유료배송비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuExchRtpChrYn=Y"												'제주도교환반품유료여부	제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuRtpAmt=6000"												'제주도반품비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&jejuExchAmt=6000"												'제주도교환비 | 제주도 추가배송비 유료일 경우
		strRst = strRst & "&prdRelspAddrCd="&DeliverAddrCd									'(*)상품출고지주소코드
		strRst = strRst & "&prdRetpAddrCd="&DeliverAddrCd									'(*)상품반송지주소코드
		getGSShopItemEditParameter = strRst
	End Function

	'// 상품 기술서(상품설명) 수정 파라메터 생성
	Public Function getGSShopContentsEditParameter()
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=D"														'(*)수정구분 D : 기술서 수정
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		'상품이미지기술서(prdDescdHtml)
		strRst = strRst & getGSShopItemContParamEucKR()
		getGSShopContentsEditParameter = strRst
	End Function


	'// 상품 정부 고시 항목 수정 파라메터 생성
	Public Function getGSShopInfodivEditParameter()
		'################################ 정부 고시 항목 최초 호출 ################################
		Dim CallInfoCd, CntInfoCd, NmInfoCd
		CallInfoCd = getGSShopItemInfoCdParam()
		CntInfoCd = Split(CallInfoCd, "|_|")(0)
		NmInfoCd = Split(CallInfoCd, "|_|")(1)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=G"														'(*)수정구분 G : 정부 고시 항목 수정
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&prdGovPublsItmListCnt="&CntInfoCd								'(*)정부고시항목리스트건수 | 1건이상입력
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		strRst = strRst & "&prdClsCd="&FDivcode												'(*)상품분류코드
		'정부고시항목(prdGovPublsItmList)
		strRst = strRst & NmInfoCd
		'rw strRst
		'response.end
		getGSShopInfodivEditParameter = strRst
	End Function

	'// 전시매장 수정 파라메터 생성
	Public Function getGSShopCategoryEditParameter()
		'################################ 매장정보 항목 최초 호출 #################################
		Dim CallCate, CntCate, NmCate
		CallCate = getGSCateParam()
		CntCate = Split(CallCate, "|_|")(0)
		NmCate = Split(CallCate, "|_|")(1)
		'##########################################################################################

		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=M"														'(*)수정구분 M : 매장정보
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&prdSectListCnt="&CntCate										'(*)매장정보리스트건수 | 1건이상입력
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		'상품전시매장(prdSectList)
		strRst = strRst & NmCate
		'rw strRst
		'response.end
		getGSShopCategoryEditParameter = strRst
	End Function

	'옵션 판매상태 수정
	Public Function getGSShopOptionEditParam()
		Dim strSql, arrRows, isOptionExists, tmpCnt
		Dim ret, bufcnt, i, itemoption, optLimit, optlimityn, isUsing, optsellyn, optNameDiff, forceExpired, ooptCd
		ret = ""
		strSql = "exec db_item.dbo.sp_Ten_OutMall_optEditParamList_gsshop '"&CMALLNAME&"'," & Fitemid
		rsget.CursorLocation = adUseClient
		rsget.CursorType = adOpenStatic
		rsget.LockType = adLockOptimistic

		rsget.Open strSql, dbget
		bufcnt = rsget.RecordCount
		If Not(rsget.EOF or rsget.BOF) Then
			arrRows = rsget.getRows
		End If
		rsget.close

		isOptionExists = isArray(arrRows)
		tmpCnt = 0
		If (isOptionExists) Then
			For i = 0 To UBound(ArrRows,2)
				itemoption			= ArrRows(1,i)
				optLimit			= ArrRows(4,i)
				optlimityn			= ArrRows(5,i)
				isUsing				= ArrRows(6,i)
				optsellyn			= ArrRows(7,i)
				optNameDiff			= (ArrRows(12,i)=1)
				forceExpired		= (ArrRows(13,i)=1)
				ooptCd				= ArrRows(15,i)

				If LEN(ooptCd) > 2 Then
					ret = ret & "&attrPrdListSupAttrPrdCd="&itemoption							'Null이라더니 Null로 전송하면 안 됨'(SYS)협력사속성상품코드 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
					ret = ret & "&attrPrdListAttrPrdCd="&ooptCd									'(*)(SYS)GS속성상품코드 | (SYS는 저희쪽에서 자동으로 생성해주는 코드 및 값을 말합니다. Null로 보내주시면 됩니다.)
					If ((forceExpired) or (optNameDiff) or (isUsing="N") or (optsellyn="N") or (optlimityn = "Y" AND optLimit <= 5)) Then
						ret = ret & "&attrPrdListSaleEndDtm="&FormatDate(now(), "00000000000000")	'(*)판매종료일시
						tmpCnt = tmpCnt + 1
					Else
						ret = ret & "&attrPrdListSaleEndDtm=29991231235959"							'(*)판매종료일시
					End If
				End If
			Next

			If FOptionCnt = 1 AND tmpCnt = 1 AND i = 1 Then
				FOptNotMatch = "Y"
			ElseIf (FOptionCnt > 1) AND (i = tmpCnt) Then
				FOptNotMatch = "Y"
			End If
		End If
		getGSShopOptionEditParam = bufcnt&"|_|"&ret
	End Function

	'// 상품 옵션 추가 및 수량 수정 파라메터 생성
	Public Function getGSShopOptParameter()
		'################################ 속성(옵션) 항목 최초 호출 ###############################
		Dim CallOpt, COptyn, CntOpt, NmOpt
		CallOpt = getGSShopOptionParam()
		COptyn = Split(CallOpt, "|_|")(0)
		CntOpt = Split(CallOpt, "|_|")(1)
		NmOpt = Split(CallOpt, "|_|")(2)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=SA"														'(*)수정구분 SA : 속성추가 및 주문가능수량수정
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&attrPrdListCnt="&CntOpt											'(*)속성[옵션]리스트건수
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		strRst = strRst & "&prdTypCd="&CHKIIF(COptyn = "Y","S","P")							'(*)상품유형코드 | 상품의 속성(옵션)이 구분을 입력합니다. P : 일반 (속성구분이 없는 경우) S : 속성 (색상/사이즈/형태/사이즈가 있는 경우) | P로 등록한 후에 S로변경하면 옵션추가가능//S->P로 일반상품 전환은 안 됨
		strRst = strRst & "&subSupCd="&COurCompanyCode										'(*)하위협력사코드 | 하위협력사로 관리하지 않는 경우 협력사코드와 동일하게 입력해주셔야 합니다.
		'상품기본-속성
		strRst = strRst & NmOpt
		getGSShopOptParameter = strRst
	End Function

	'// 상품 옵션 상태 수정 파라메터 생성
	Public Function getGSShopOptSellParameter()
		'################################ 속성(옵션) 항목 최초 호출 ###############################
		Dim CallOptSell, COptyn, CntOptSell, NmOptSell
		CallOptSell	= getGSShopOptionEditParam()
		CntOptSell	= Split(CallOptSell, "|_|")(0)
		NmOptSell	= Split(CallOptSell, "|_|")(1)
		'##########################################################################################
		Dim strRst
		strRst = ""
		strRst = strRst & "regGbn=U"														'(*)등록구분 U : 수정
		strRst = strRst & "&modGbn=SS"														'(*)수정구분 SS : 속성판매종료
		strRst = strRst & "&regId="&COurRedId												'(*)등록자
		strRst = strRst & "&attrPrdListCnt="&CntOptSell										'(*)속성[옵션]리스트건수
		'상품기본(prdBaseInfo)
		strRst = strRst & "&supPrdCd="&Fitemid												'(*)협력사상품코드
		strRst = strRst & "&supCd="&COurCompanyCode											'(*)협력사코드
		'상품기본-속성
		strRst = strRst & NmOptSell
		getGSShopOptSellParameter = strRst
	End Function

End Class

Class CGSShop
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
	Public FRectItemName
	Public FRectSellYn
	Public FRectLimitYn
	Public FRectSailYn
	Public FRectonlyValidMargin
	Public FRectGSShopgoodno
	Public FRectMatchCate
	Public FRectPrdDivMatch
	Public FRectIsMadeHand
	Public FRectIsOption
	Public FRectIsReged
	Public FRectExtNotReg
	Public FRectExpensive10x10
	Public FRectdiffPrc
	Public FRectGSShopYes10x10No
	Public FRectGSShopNo10x10Yes
	Public FRectExtSellYn
	Public FRectInfoDiv
	Public FRectFailCntOverExcept
	Public FRectFailCntExists
	Public FRectReqEdit

    ''정렬순서
    Public FRectOrdType

	'브랜드 관리
	Public FRectIsMaeip
	Public FRectIsDeliMapping
	Public FRectIsbrandcd
	Public FRectCatekey

	'상품분류
	Public FInfodiv
	Public FCateName
	Public FsearchName

	'카테고리
	Public FRectIsMapping
	Public FRectSDiv
	Public FRectKeyword
	Public FRectCDL
	Public FRectCDM
	Public FRectCDS
	Public FRectdisptpcd
	Public FRectDspNo

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
	Public Sub getGSShopNotRegOneItem
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

            ''' 2013/05/29 특정품목 등록 불가 (화장품, 식품류)
            'addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18','20','21','22')"
			addSql = addSql & " and isNULL(c.infodiv,'') not in ('','18')"	''2022-06-17 예슬님요청..식품 판매
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent , isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(R.gsshopStatCD,-9) as gsshopStatCD, IsNull(R.GSShopPrice, 0) as GSShopPrice "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.dtlCd, '') as divcode, isnull(pm.safecode, '') as safecode, uc.socname_kor "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " JOIN db_user.dbo.tbl_user_c as uc on i.makerid = uc.userid "
		strSql = strSql & " JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_regItem R on i.itemid=R.itemid"
		strSql = strSql & " WHERE i.isusing = 'Y' "
		strSql = strSql & " and i.isExtUsing = 'Y' "
		strSql = strSql & " and i.deliverytype not in ('7','6')"
		IF (CUPJODLVVALID) then
'		    strSql = strSql & " and ((i.deliveryType <> 9) or ((i.deliveryType = 9) and (i.sellcash >= 10000)))"
		ELSE
		    strSql = strSql & "	and (i.deliveryType <> 9)"
	    END IF
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.itemdiv not in ('21', '23', '30') "
		strSql = strSql & " and i.deliverfixday not in ('C','X','G') "						'플라워/화물배송/해외직구 상품 제외
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.itemdiv < 50 and i.itemdiv <> '08' "
'		strSql = strSql & " and i.itemdiv not in ('06', '16') "
		strSql = strSql & " and i.cate_large <> '' "
		strSql = strSql & " and i.cate_large <> '999' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and ((i.LimitYn = 'N') or ((i.LimitYn = 'Y') and (i.LimitNo-i.LimitSold>="&CMAXLIMITSELL&")) )" ''한정 품절 도 등록 안함.
'		strSql = strSql & " and (i.sellcash <> 0 and ((i.sellcash - i.buycash)/i.sellcash)*100 >= " & CMAXMARGIN & ")"
		strSql = strSql & " and (i.sellcash <> 0) "
		strSql = strSql & " and 'Y' = CASE WHEN i.sailyn = 'Y' "
		strSql = strSql & " 				AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 					OR (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) >= " & CMAXMARGIN & ") "
		strSql = strSql & " 				) THEN 'Y' "
		strSql = strSql & " 				WHEN i.sailyn = 'N' AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) >= " & CMAXMARGIN & ") THEN 'Y' ELSE 'N' END "
		strSql = strSql & "	and i.makerid not in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop') "	'등록제외 브랜드
		strSql = strSql & "	and i.itemid not in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"') "		'등록제외 상품
		strSql = strSql & "	and (convert(varchar(6), (i.cate_large + i.cate_mid)) not in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "	'등록제외 카테고리
		strSql = strSql & "	and i.itemid not in (Select itemid From db_item.dbo.tbl_gsshop_regItem where gsshopStatCD>3) "	''gsshopStatCD>=3 등록완료이상은 등록안됨.
		strSql = strSql & "	and IsNull(R.GSShopGoodNo, '') = '' "									'등록상품 제외
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGSShopItem
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
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.Fsourcearea		= rsget("sourcearea")
				FOneItem.Fmakername			= rsget("makername")
				FOneItem.FUsingHTML			= rsget("usingHTML")
				FOneItem.Fitemcontent		= db2html(rsget("itemcontent"))
                FOneItem.FGSShopStatCD		= rsget("gsshopStatCD")
				FOneItem.FGsshopprice		= rsget("gsshopprice")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                FOneItem.FsafetyDiv			= rsget("safetyDiv")
                FOneItem.FsafetyNum			= rsget("safetyNum")
                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
'                FOneItem.FBrandcd			= rsget("brandcd")
                FOneItem.FDeliveryType		= rsget("deliveryType")
'                FOneItem.FDeliveryCd		= rsget("deliveryCd")
'                FOneItem.FDeliveryAddrCd	= rsget("deliveryAddrCd")
                FOneItem.FrequireMakeDay    = rsget("requireMakeDay")
				FOneItem.FAdultType 		= rsget("adulttype")
				FOneItem.FSocname_kor 		= rsget("socname_kor")
		End If
		rsget.Close
	End Sub

	'// GSShop 상품 목록(수정용)
	Public Sub getGSShopEditOneItem
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
		strSql = strSql & "	, m.gsshopGoodNo, m.gsshopprice, m.gsshopSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt, isNULL(convert(char(1), m.regedOptCnt), 'Y') as isNulltoTimeout  "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, m.regImageName "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum "
		strSql = strSql & "	, isnull(pm.dtlcd, '') as divcode, isnull(pm.safecode, '') as safecode "
        strSql = strSql & "	,(CASE WHEN i.isusing='N' "
		strSql = strSql & "		or i.isExtUsing='N'"
		strSql = strSql & "		or uc.isExtUsing='N'"
'		strSql = strSql & "		or ( (i.sailyn='N') and (i.deliveryType = 9) and (i.sellcash < 10000))"
		strSql = strSql & "		or i.deliveryType in ('7','6') "
		strSql = strSql & "		or i.sellyn<>'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X','G')"
		strSql = strSql & "		or i.itemdiv in ('21', '23', '30')"
		strSql = strSql & "		or i.itemdiv >= 50 or i.itemdiv = '08' or i.cate_large = '999' or i.cate_large=''"
'		strSql = strSql & "		or i.itemdiv = '06' or i.itemdiv = '16' "

		strSql = strSql & "		or ((i.sailyn = 'Y') AND ( (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") AND (Round(((i.orgprice - i.orgsuplycash)/ i.orgprice)*100,0) < " & CMAXMARGIN & "))) "
		strSql = strSql & "		or (i.sailyn = 'N') AND (Round(((i.sellcash - i.buycash)/ i.sellcash)*100,0) < " & CMAXMARGIN & ") "

		strSql = strSql & "		or i.makerid  in (Select makerid From [db_temp].dbo.tbl_jaehyumall_not_in_makerid Where mallgubun='gseshop')"
		strSql = strSql & "		or i.itemid  in (Select itemid From [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun='"&CMALLNAME&"')"
		strSql = strSql & "		or (convert(varchar(6), (i.cate_large + i.cate_mid)) in (SELECT convert(varchar(6), cdl+cdm) FROM db_temp.[dbo].[tbl_jaehyumall_not_in_category] WHERE mallgubun='"&CMALLNAME&"')) "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_item.dbo.tbl_GSShop_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN (Select tenCateLarge, tenCateMid, tenCateSmall, count(*) as mapCnt From db_item.dbo.tbl_gsshop_cate_mapping Group by tenCateLarge, tenCateMid, tenCateSmall ) as cm on cm.tenCateLarge=i.cate_large and cm.tenCateMid=i.cate_mid and cm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_gsshop_MngDiv_mapping as pm on pm.tenCateLarge=i.cate_large and pm.tenCateMid=i.cate_mid and pm.tenCateSmall=i.cate_small "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and (m.GSShopStatCd = 3 OR m.GSShopStatCd = 7)  "
		strSql = strSql & addSql
		strSql = strSql & " and m.gsshopGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CGSShopItem
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
				FOneItem.FGsshopGoodNo		= rsget("gsshopGoodNo")
				FOneItem.FGsshopprice		= rsget("gsshopprice")
				FOneItem.FGsshopSellYn		= rsget("gsshopSellYn")
				FOneItem.FbasicimageNm 		= rsget("basicimage")
				FOneItem.FregImageName		= rsget("regImageName")
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
	            FOneItem.FIsNulltoTimeout    = rsget("isNulltoTimeout")

                FOneItem.FDivcode			= rsget("divcode")
                FOneItem.FSafecode			= rsget("safecode")
				FOneItem.FAdultType 		= rsget("adulttype")
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
