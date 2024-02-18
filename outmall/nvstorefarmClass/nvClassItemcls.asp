<%
CONST CMAXMARGIN = 10
CONST CMALLNAME = "nvstorefarmclass"
CONST CUPJODLVVALID = TRUE								''업체 조건배송 등록 가능여부
CONST CMAXLIMITSELL = 0									'' 이 수량 이상이어야 판매함. // 옵션한정도 마찬가지.
CONST CDEFALUT_STOCK = 9999

Class CNvClassItem
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
	Public FNvClassGoodNo
	Public FNvClassprice
	Public FNvClassSellyn
	Public FregedOptCnt
	Public FAccFailCNT
	Public FMaySoldOut
	Public Fregitemname
	Public FLastErrStr
	Public FRequireMakeDay
	Public FSafetyyn
	Public FSafetyDiv
	Public FSafetyNum
	Public FNvClassStatCD
	Public FinfoDiv
	Public FDeliveryType
	Public FSocname_kor
	Public FAPIaddImg
	Public FbasicimageNm
	Public FRegImageName
	Public FCateKey
	Public FNeedCert

	Private Sub Class_Initialize()
	End Sub

	Private Sub Class_Terminate()
	End Sub

	Public Function MustPrice()
		MustPrice = FSellCash
	End Function

    public function getBasicImage()
        if IsNULL(FbasicImageNm) or (FbasicImageNm="") then Exit function
        getBasicImage = FbasicImageNm
    end function

	'// 스토어팜 판매여부 반환
	Public Function getNvClassSellYn()
		'판매상태 (10:판매진행, 20:품절)
		If FSellYn="Y" and FIsUsing="Y" then
			If FLimitYn = "N" or (FLimitYn = "Y" and FLimitNo - FLimitSold >= CMAXLIMITSELL) then
				getNvClassSellYn = "Y"
			Else
				getNvClassSellYn = "N"
			End If
		Else
			getNvClassSellYn = "N"
		End If
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


	Function GetRaiseValue(value)
		If Fix(value) < value Then
			GetRaiseValue = Fix(value) + 1
		Else
			GetRaiseValue = Fix(value)
		End If
	End Function

	Public Function getLimitNvClassEa()
		Dim ret
		If FLimitYn = "Y" Then
			ret = FLimitNo - FLimitSold
			If ret > 10000 Then
				ret = CDEFALUT_STOCK
			End If
		Else
			ret = CDEFALUT_STOCK
		End If

		If (ret < 1) Then ret = 0
		getLimitNvClassEa = ret
	End Function

	Public Function isImageChanged()
		Dim ibuf : ibuf = getBasicImage
		If InStr(ibuf,"-") < 1 Then
			isImageChanged = FALSE
			Exit Function
		End If
		isImageChanged = ibuf <> FRegImageName
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

				If i = 0 Then		'0원 옵션의 재고가 0개 이하면 품절
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

	Function getItemNameFormat()
		Dim buf
		'buf = "[텐바이텐 클래스] "&replace(FItemName,"'","")		'최초 상품명 앞에 [텐바이텐 클래스] 이라고 붙임
		buf = replace(FItemName,"'","")
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
		buf = "				<shop:CustomMade>N</shop:CustomMade>"		'# 주문 제작 상품 여부 Y or N | Y: EstimatedDeliveryTime입력 필수, "N": EstimatedDeliveryTime 입력 불가
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

	'// 상품등록: 상품추가이미지 파라메터 생성
	Public Function getImageType()
		Dim buf, strSql, arrRows, i, basicimgStr, addimgStr
		addimgStr	= ""
		basicimgStr	= ""
		strSql = ""
		strSql = strSql & " SELECT TOP 10 imgType, storefarmURL FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_Image] WHERE itemid = '"&FItemid&"' "
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

	'// 쿠폰정보
	Public Function getECouponType()
		Dim buf, strSql, arrRows, i, UsePlaceContents, ContactInformationContents
		strSql = ""
		strSql = strSql & " SELECT TOP 1 p.tPAddress, p.tPTel "
		strSql = strSql & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSql = strSql & " WHERE k.itemid = '"& FItemID &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			UsePlaceContents = rsget("tPAddress")
			ContactInformationContents = rsget("tPTel")
		End If
		rsget.Close

		If UsePlaceContents = "" Then
			UsePlaceContents = "서울시 종로구 대학로12길 31 자유빌딩 2층"
		End If

		If ContactInformationContents = "" Then
			ContactInformationContents = "1644-6030"
		End If

		buf = ""
		buf = buf & "				<shop:AfterServiceTelephoneNumber><![CDATA["&ContactInformationContents&"]]></shop:AfterServiceTelephoneNumber>"		'#A/S 전화번호
		buf = buf & "				<shop:AfterServiceGuideContent><![CDATA[A/S 관련은 "&FSocname_kor&" 강사님을 통해 문의해 주시기 바랍니다.]]></shop:AfterServiceGuideContent>"	'#A/S 안내
		buf = buf & "				<shop:ECoupon>"											'ECOUPON | 이쿠폰 카테고리 상품인 경우 필수
		buf = buf & "					<shop:PeriodType>FB</shop:PeriodType>"				'#e쿠폰 유효기간 구분 | FX : 특정기간, FB : 자동기간
'		buf = buf & "					<shop:ValidStartDate></shop:ValidStartDate>"		'e쿠폰 유효기간 시작일..YYYY-MM-DD형식, e쿠폰 유효기간 구분 타입(PeriodType)이 특정기간인 경우 필수
'		buf = buf & "					<shop:ValidEndDate></shop:ValidEndDate>"			'e쿠폰 유효기간 종료일..YYYY-MM-DD형식, e쿠폰 유효기간 구분 타입(PeriodType)이 특정기간인 경우 필수
		buf = buf & "					<shop:PeriodDays>30</shop:PeriodDays>"				'e쿠폰 유효기간 내용..e쿠폰 유효기간 구분 타입(PeriodType)이 자동 기간인 경우 필수
		buf = buf & "					<shop:PublicInformationContents><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></shop:PublicInformationContents>"		'e쿠폰 발행처
		buf = buf & "					<shop:ContactInformationContents><![CDATA["&ContactInformationContents&"]]></shop:ContactInformationContents>"	'e쿠폰 연락처
		buf = buf & "					<shop:UsePlaceType>PLACE</shop:UsePlaceType>"		'e쿠폰 사용 장소 타입 | PLACE : 장소, URL : URL
		buf = buf & "					<shop:UsePlaceContents><![CDATA["& UsePlaceContents &"]]></shop:UsePlaceContents>"	'e쿠폰 사용 장소
		buf = buf & "					<shop:RestrictCart>Y</shop:RestrictCart>"			'e쿠폰 장바구니 제한 | Y or N
		buf = buf & "				</shop:ECoupon>"
		getECouponType = buf
	End Function

	Public Function getNvClassItemContParamToReg()
		Dim strRst, strSQL
		strRst = ("<div align=""center"">")
'		strRst = strRst & ("<p><center><a href=""http://storefarm.naver.com/tenbytenclass"" target=""_blank""><img src=""http://fiximage.10x10.co.kr/web2008/etc/top_notice_nvClass.jpg""></a></center></p><br>")

'		If ForderComment <> "" Then
'			strRst = strRst & "- 주문시 유의사항 :<br>" & Fordercomment & "<br>"
'		End If

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
'		strRst = strRst & ("<br><img src=""http://fiximage.10x10.co.kr/web2008/etc/cs_info_nvClass.jpg"">")
		strRst = strRst & ("</div>")

		Dim ticketPlaceName, tPAddress, tPTel, parkingGuide
		Dim strticketPlaceName, strtPAddress, strtPTel, strparkingGuide
		strSQL = ""
		strSQL = strSQL & " SELECT TOP 1 isNull(p.ticketPlaceName, '') as ticketPlaceName, isNull(p.tPAddress, '') as tPAddress, isNull(p.tPTel, '') as tPTel, isNull(p.parkingGuide, '') as parkingGuide "
		strSQL = strSQL & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSQL = strSQL & " WHERE k.itemid = '"& Fitemid &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			ticketPlaceName	= rsget("ticketPlaceName")
			tPAddress		= rsget("tPAddress")
			tPTel			= rsget("tPTel")
			parkingGuide	= rsget("parkingGuide")
		End If
		rsget.Close

		If ticketPlaceName <> "" Then
			strticketPlaceName = "<strong>[장소명]</strong><br />" & ticketPlaceName & "<br />"
		End If

		If tPAddress <> "" Then
			strtPAddress = "<strong>[주소]</strong><br />" & tPAddress & "<br />"
		End If

		If tPTel <> "" Then
			strtPTel = "<strong>[전화]</strong><br />" & tPTel & "<br />"
		End If

		If parkingGuide <> "" Then
			strparkingGuide = "<strong>[주차 정보]</strong><br />" & nl2br(parkingGuide)
		End If

		If (ticketPlaceName <> "") OR (tPAddress <> "") OR (tPTel <> "") OR (parkingGuide <> "") Then
			strRst = strRst & "<div class=""alignCt"" style=""background-color:#f8f8f8; margin-top:100px; padding:57px 0px; width:100%"">"
			strRst = strRst & "<p style=""margin-bottom:0px; margin-left:0px; margin-right:0px; margin-top:0px; padding:0px 8%; text-align:center""><span style=""font-family:malgun gothic,&quot;맑은 고딕&quot;,sans-serif""><span style=""color:#000000""><span style=""font-size:22px; font-weight:600; line-height:1.2"">위치 정보</span></span></span></p>"
			strRst = strRst & "	<p style=""margin-bottom:0px; margin-left:0px; margin-right:0px; margin-top:0px; padding:0px 8%; text-align:center"">&nbsp;</p>"
			strRst = strRst & "	<div style=""padding:11px 8% 0px; text-align:left"">"
			strRst = strRst & "		<span style=""font-family:malgun gothic,&quot;맑은 고딕&quot;,sans-serif"">"
			strRst = strRst & "			<span style=""font-size:16px"">"
			strRst = strRst & "				<span style=""color:#000000"">"
			strRst = strRst & "					<span style=""line-height:1.6"">"
			strRst = strRst & "						"& strticketPlaceName &" "
			strRst = strRst & "						"& strtPAddress &" "
			strRst = strRst & "						"& strtPTel &" "
			strRst = strRst & "						"& strparkingGuide &" "
			strRst = strRst & "					</span>"
			strRst = strRst & "				</span>"
			strRst = strRst & "			</span>"
			strRst = strRst & "		</span>"
			strRst = strRst & "		</p>"
			strRst = strRst & "	</div>"
			strRst = strRst & "</div>"
		End If
		getNvClassItemContParamToReg = strRst
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
		icomment = replace(icomment,"&#160;"," ")	'2018-12-27 이상한 아스키값 치환..maybw 엑셀에서 심어나온듯
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

	Public Function getNvClassItemInfoCdToReg
		Dim buf, strSQL, mallinfoCd, infoContent, mallinfodiv, mallinfoName
		strSQL = ""
		strSQL = strSQL & " SELECT top 100 M.* , " & vbcrlf
		strSQL = strSQL & " CASE WHEN (M.infoCd='00000') AND (IC.safetyyn= 'Y') THEN isNull(TR.certNum, IC.safetyNum) " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (IC.safetyyn<> 'Y') THEN '상세페이지 참조' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00000') AND (isNULL(IC.safetyyn,'N')= 'N') THEN '해당없음' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='00001') THEN '상세페이지 참조' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='10000') THEN '관련법 및 소비자분쟁해결기준에 따름' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'N') THEN 'N' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd in ('17008', '21007', '21009', '22010', '22012')) AND (F.chkdiv = 'Y') THEN 'Y' " & vbcrlf
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) < 2 THEN i.itemname "
		strSQL = strSQL & " 	 WHEN (M.infoCd='21011') AND LEN(isnull(F.infocontent, '')) >= 2 THEN F.infocontent "
		strSQL = strSQL & " 	 WHEN c.infotype='P' THEN K.tPTel " & vbcrlf
		strSQL = strSQL & " 	 WHEN LEN(isnull(F.infocontent, '')) < 2 THEN '상세페이지 참조' " & vbcrlf
		strSQL = strSQL & " ELSE isnull(F.infocontent, '') END AS infocontent " & vbcrlf
		strSQL = strSQL & " FROM db_item.dbo.tbl_OutMall_infoCodeMap M " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item_contents IC ON IC.infoDiv=M.mallinfoDiv " & vbcrlf
		strSQL = strSQL & " INNER JOIN db_item.dbo.tbl_item I ON IC.itemid=I.itemid " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCode c ON M.infocd=c.infocd " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.dbo.tbl_item_infoCont F ON M.infocd=F.infocd and F.itemid='"&FItemid&"' " & vbcrlf
		strSQL = strSQL & " LEFT JOIN db_item.[dbo].[tbl_safetycert_tenReg] as TR ON i.itemid = TR.itemid "
		strSQL = strSQL & " LEFT JOIN ( "
		strSQL = strSQL & " 	SELECT TOP 1 k.itemid, isNull(p.tPTel, '텐바이텐 1644-6030') as tPTel "
		strSQL = strSQL & " 	FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSQL = strSQL & " 	LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSQL = strSQL & " 	WHERE k.itemid = '"& FItemID &"' "
		strSQL = strSQL & " ) as K on K.itemid = IC.itemid "
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
		getNvClassItemInfoCdToReg = buf
	End Function

	Public Function getNvClassItemInfoCdToRegOnlyMobile
		Dim buf, strSql, arrRows, i, UsePlaceContents, ContactInformationContents
		strSql = ""
		strSql = strSql & " SELECT TOP 1 p.tPAddress, p.tPTel "
		strSql = strSql & " FROM db_item.dbo.tbl_ticket_itemInfo k "
		strSql = strSql & " LEFT JOIN db_item.dbo.tbl_ticket_PlaceInfo p on k.ticketPlaceIdx = p.ticketPlaceIdx "
		strSql = strSql & " WHERE k.itemid = '"& FItemID &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If not rsget.Eof Then
			UsePlaceContents = rsget("tPAddress")
			ContactInformationContents = rsget("tPTel")
		End If
		rsget.Close

		If UsePlaceContents = "" Then
			UsePlaceContents = "서울시 종로구 대학로12길 31 자유빌딩 2층"
		End If

		If ContactInformationContents = "" Then
			ContactInformationContents = "1644-6030"
		End If

		buf = ""
		buf = buf & "				<shop:MobileCoupon>"
		buf = buf & "					<NoRefundReason><![CDATA[상세페이지 참조]]></NoRefundReason>"							'제품하자가 아닌 소비자의 단순변심, 착오구매 에 따른 청약철회 등이 불가능한 경우 그 구 체적 사유와 근거
		buf = buf & "					<ReturnCostReason><![CDATA[상세페이지 참조]]></ReturnCostReason>"						'제품하자?오배송 등에 따른 청약철회 등의 경 우 청약철회 등을 할 수 있는 기간 및 통신판 매업자가 부담하는 반품비용 등에 관한 정보
		buf = buf & "					<QualityAssuranceStandard><![CDATA[상세페이지 참조]]></QualityAssuranceStandard>"		'재화 등의 교환·반품·보증 조건 및 품질 보증 기준
		buf = buf & "					<CompensationProcedure><![CDATA[상세페이지 참조]]></CompensationProcedure>"				'대금을 환불받기 위한 방법과 환불이 지연될 경우 지연에 따른 배상금을 지급받을 수 있다 는 사실 및 배상금 지급의 구체적 조건 및 절 차
		buf = buf & "					<TroubleShootingContents><![CDATA[상세페이지 참조]]></TroubleShootingContents>"			'소비자 피해 보상의 처리, 재화 등에 대한 불 만 처리 및 소비자와 사업자 사이의 분쟁처리 에 관한 사항
		buf = buf & "					<Issuer><![CDATA["&chkIIF(trim(Fmakername)="" or isNull(Fmakername),"상품설명 참조",Fmakername)&"]]></Issuer>"		'발행자
		buf = buf & "					<UsableCondition><![CDATA[구매일로부터 30일]]></UsableCondition>"						'유효기간, 이용 조건
		buf = buf & "					<UsableStore><![CDATA["& UsePlaceContents &"]]></UsableStore>"							'이용 가능 매장
		buf = buf & "					<CancelationPolicy><![CDATA[상세페이지 참조]]></CancelationPolicy>"						'환불 조건 및 방법
		buf = buf & "					<CustomerServicePhoneNumber><![CDATA["&ContactInformationContents&"]]></CustomerServicePhoneNumber>"	'소비자 상담 관련 전화번호
		buf = buf & "				</shop:MobileCoupon>"
		getNvClassItemInfoCdToRegOnlyMobile = buf
	End Function

	'// 업로드 이미지 XML 생성
	Public Function getNvClassImageRegXML(oServ, oOper)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature, strSQL, i
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1np6kl_01"
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
		End If
		strRst = strRst & "			</ImageURLList>"
		strRst = strRst & "		</shop:UploadImageRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		getNvClassImageRegXML = strRst
	End Function

	'// 상품등록 XML 생성
	Public Function getNvClassItemRegXML(oServ, oOper, isEdit)
		Dim strRst, reqID, oaccessLicense, oTimestamp, osignature
		If (application("Svr_Info") = "Dev") Then
			reqID = "qa2tc329"
		Else
			reqID = "ncp_1np6kl_01"
		End If
		Call getsecretKey(oaccessLicense, oTimestamp, osignature, oServ, oOper)

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
		strRst = strRst & "			<SellerId>"&reqID&"</SellerId>"
		strRst = strRst & "			<Product>"
		If isEdit = "Y" Then
			strRst = strRst & "			<shop:ProductId>"&FNvClassGoodNo&"</shop:ProductId>"		'상품ID | 없으면 등록, 있으면 수정
		End If
		strRst = strRst & "				<shop:StatusType>SALE</shop:StatusType>"			'# 상품판매상태 | 등록은 SALE(판매중)만 입력, 수정시 SALE, SUSP(판매 중지)만 입력, StockQuantity가 0 이면 OSTK(품절)로 저장됨
		strRst = strRst & "				<shop:SaleType>NEW</shop:SaleType>"					'상품 판매 유형..미입력시 NEW로 저장
		strRst = strRst & getzCostomMadeInd													'#주문 제작 상품 여부

		''test입니다 #######################################################################
		If FItemid = "2525634" Then
			strRst = strRst & "				<shop:CategoryId>50007215</shop:CategoryId>"		'#Leaf 카테고리 | ID 상품등록시 필수 | ModelType의 모델명ID가 입력된 경우 해당 모델명 ID에 매핑된  Leaf 카테고리 ID로 저장하며 요청으로 전달된 CategoryId는 무시된다
		Else
			strRst = strRst & "				<shop:CategoryId>50007332</shop:CategoryId>"		'#Leaf 카테고리 | ID 상품등록시 필수 | ModelType의 모델명ID가 입력된 경우 해당 모델명 ID에 매핑된  Leaf 카테고리 ID로 저장하며 요청으로 전달된 CategoryId는 무시된다
		End If

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
		strRst = strRst & "				</shop:Model>"
'		strRst = strRst & "				<shop:AttributeValueList></shop:AttributeValueList>"		' ,로 분리된 속성의 목록 | 현재는 사용하지 않으며 향후 사용 예정
		strRst = strRst & getOriginAreaType															'#원산지 정보
'		strRst = strRst & "				<shop:ManufactureDate></shop:ManufactureDate>"				'제조 일자 | YYYY-MM-DD 형식
'		strRst = strRst & "				<shop:ValidDate></shop:ValidDate>"							'유효 일자 | YYYY-MM-DD 형식
		strRst = strRst & "				<shop:TaxType>"&CHKIIF(FVatInclude="N","DUTYFREE","TAX")&"</shop:TaxType>"	'#부가세 | 과세 : TAX, 면세 : DUTYFREE, 영세 : SMALL
		strRst = strRst & "				<shop:MinorPurchasable>Y</shop:MinorPurchasable>"			'#미성년자 구매 가능 여부 Y or N
		strRst = strRst & getImageType																'#이미지 정보
		strRst = strRst & "				<shop:DetailContent><![CDATA["&getNvClassItemContParamToReg&"]]></shop:DetailContent>"		'#상품 상세 정보
'		strRst = strRst & "				<shop:SellerNoticeId></shop:SellerNoticeId>"										'공지사항 번호
'		strRst = strRst & "				<shop:PurchaseReviewExposure></shop:PurchaseReviewExposure>"						'구매평 노출 여부 | Y or N, 구매평 노출 설정 가능 카테고리일 경우에만 유효하며 그 외에는 Y로 설정된다. 미입력 시 Y로 저장됨
'		strRst = strRst & "				<shop:RegularCustomerExclusiveProduct></shop:RegularCustomerExclusiveProduct>"		'단골 회원 전용 상품 여부 | Y or N 미입력시 N으로 저장됨
'		strRst = strRst & "				<shop:KnowledgeShoppingProductRegistration></shop:KnowledgeShoppingProductRegistration>"	'네이버 쇼핑 등록 | Y or N 네이버 광고주가 아닌 경우 N으로 저장됨
'		strRst = strRst & "				<shop:GalleryId></shop:GalleryId>"							'갤러리 번호
'		strRst = strRst & "				<shop:SaleStartDate></shop:SaleStartDate>"					'판매 시작일 | YYYY-MM-DD 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00분으로만 설정 가능
'		strRst = strRst & "				<shop:SaleEndDate></shop:SaleEndDate>"						'판매 종료일 | YYYY-MM-DD HH:mm형식..날짜까지만 입력하는 경우 자동으로 23시 59분을 붙여서 저장됨.매시각 59분으로만 설정 가능
		strRst = strRst & "				<shop:SalePrice>"&Clng(GetRaiseValue(MustPrice/10)*10)&"</shop:SalePrice>"		'#판매가
		If (isEdit = "Y")  Then
			If (Foptioncnt = 0) Then
				strRst = strRst & "				<shop:StockQuantity>"&getLimitNvClassEa&"</shop:StockQuantity>"		'#재고 수량 | 상품등록시 필수, 상품 수정시 재고 수량을 입력하지 않으면 스토어팜 DB에 저장된 현재 재고값이 변하지 않는다. 수정시 재고 수량 0으로 입력되면 StatusType으로 전달된 항목은 무시되며 상품 상태는 OSTK(품절)로 저장됨
			End If
		Else
			strRst = strRst & "				<shop:StockQuantity>"&getLimitNvClassEa&"</shop:StockQuantity>"		'#재고 수량 | 상품등록시 필수, 상품 수정시 재고 수량을 입력하지 않으면 스토어팜 DB에 저장된 현재 재고값이 변하지 않는다. 수정시 재고 수량 0으로 입력되면 StatusType으로 전달된 항목은 무시되며 상품 상태는 OSTK(품절)로 저장됨
		End If
'		strRst = strRst & "				<shop:MinPurchaseQuantity></shop:MinPurchaseQuantity>"					'최소 구매 수량
'		strRst = strRst & "				<shop:MaxPurchaseQuantityPerId></shop:MaxPurchaseQuantityPerId>"		'1인 최대 구매 수량
'		strRst = strRst & "				<shop:MaxPurchaseQuantityPerOrder></shop:MaxPurchaseQuantityPerOrder>"	'1회 최대 구매 수량
'		strRst = strRst & "				<shop:SellerDiscount>"									'판매자 즉시 할인 | 선택이나, 입력할 경우 하단 #은 필수
'		strRst = strRst & "					<shop:Amount></shop:Amount>"						'#PC 즉시 할인액/할인율 | PC할인만 적용하려면 MobileAmount에는 0을 입력..끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
'		strRst = strRst & "					<shop:StartDate></shop:StartDate>"					'PC 즉시 할인 시작일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00, 10, 20, 30, 40, 50분으로만 설정 가능
'		strRst = strRst & "					<shop:EndDate></shop:EndDate>"						'PC 즉시 할인 종료일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 23시 59분을 붙여서 저장됨..매시각 09, 19, 29, 39, 49, 59분으로만 설정 가능
'		strRst = strRst & "					<shop:MobileAmount></shop:MobileAmount>"			'#모바일 즉시 할인액/할인율 | 모바일 할인만 적용하려면 Amount에 0을 입력..끝문자(%, 숫자)에 따라 단위가 구분됨..ex)값이 10%이면 할인율, 1000이면 할인액을 나타낸다
'		strRst = strRst & "					<shop:MobileStartDate></shop:MobileStartDate>"		'모바일 즉시 할인 시작일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 자동으로 0시0분을 붙여서 저장됨.매시각 00, 10, 20, 30, 40, 50분으로만 설정 가능
'		strRst = strRst & "					<shop:MobileEndDate></shop:MobileEndDate>"			'모바일 즉시 할인 종료일 | YYYY-MM-DD HH:mm 형식..날짜까지만 입력하는 경우 23시 59분을 붙여서 저장됨..매시각 09, 19, 29, 39, 49, 59분으로만 설정 가능
'		strRst = strRst & "				</shop:SellerDiscount>"
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

		''test입니다 #######################################################################
		If Fitemid = "2525634" Then
		Else
			strRst = strRst & getECouponType
		End If
'		strRst = strRst & "				<shop:PurchaseApplicationUrl></shop:PurchaseApplicationUrl>"	'휴대폰 구매신청서 URL | 휴대폰 카테고리 상품인 경우 필수
'		strRst = strRst & "				<shop:CellPhonePrice></shop:CellPhonePrice>"					'고객부담 휴대폰 단말기 대금 | 휴대폰 카테고리 상품인 경우 필수
'		strRst = strRst & "				<shop:WifiOnly></shop:WifiOnly>"		'Wifi 전용 상품 여부 | Y or N..태블릿 카테고리 상품인 경우 필수..Y 입력시 PurchaseApplicationUrl, CellPhonePrice 입력불가..N 입력시 PurchaseApplicationUrl, CellPhonePrice 입력 필수
		strRst = strRst & "				<shop:ProductSummary>"					'상품 요약 정보 | 상품 등록시 필수, 상품 수정 시에는 기존에 상품 요약 정보가 입력된 경우에만 생략할 수 있다. 이 경우 기존에 저장된 상품 요약 정보 값이 유지된다.

		''test입니다 #######################################################################
		If Fitemid = "2525634" Then
			strRst = strRst & getNvClassItemInfoCdToRegOnlyMobile
		Else
			strRst = strRst & getNvClassItemInfoCdToReg
		End If
		strRst = strRst & "				</shop:ProductSummary>"
		strRst = strRst & getSellerComment
		strRst = strRst & "			</Product>"
		strRst = strRst & "		</shop:ManageProductRequest>"
		strRst = strRst & "	</soapenv:Body>"
		strRst = strRst & "</soapenv:Envelope>"
		''test입니다 #######################################################################
		If Fitemid = "2525634" Then
			response.write strRst
		End If
'response.end
		getNvClassItemRegXML = strRst
	End Function
End Class

Class CNvClass
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

	Public Sub getNvClassNotRegOneItem
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
		End If

		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent "
		strSql = strSql & "	, isNULL(R.nvClassStatCD,-9) as nvClassStatCD"
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
		'strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName, isnull(ca.needCert, '') as needCert "
		strSql = strSql & " ,isNULL(R.regImageName,'') as regImageName "
		strSql = strSql & "	, isnull(R.APIaddImg, '') as APIaddImg "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid=c.itemid "
		strSql = strSql & " LEFT JOIN db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] R on i.itemid = R.itemid"
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1 "
		strSql = strSql & " and i.isusing='Y' "
		strSql = strSql & " and i.deliverfixday not in ('C','X') "
		strSql = strSql & " and i.basicimage is not null "
		strSql = strSql & " and i.cate_large<>'' "
		strSql = strSql & " and i.itemdiv in ('08') "	'티켓/클래스 상품
		strSql = strSql & " and (i.cate_large = '035' and i.cate_mid = '022' and i.cate_small = '010') " '여행/취미 > 취미/강좌 > 클래스 만 등록되야 함
		strSql = strSql & " and i.sellyn = 'Y' "
		strSql = strSql & " and i.sellcash > 0 "
		strSql = strSql & " and not exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstorefarmclass') "
		If FRectGubun <> "IMG" Then
			strSql = strSql & "	and i.itemid not in (Select itemid From db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] where nvClassStatCD > 3) "
		End If
		strSql = strSql & addSql
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvClassItem
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
                FOneItem.FNvClassStatCD		= rsget("nvClassStatCD")
                FOneItem.FinfoDiv			= rsget("infoDiv")
                FOneItem.FDeliveryType		= rsget("deliveryType")
                FOneItem.FSocname_kor		= rsget("socname_kor")
                FOneItem.FAPIaddImg			= rsget("APIaddImg")
                FOneItem.FbasicimageNm 		= rsget("basicimage")
                FOneItem.FRegImageName 		= rsget("regImageName")
                FOneItem.Fsafetyyn			= rsget("safetyyn")
                'FOneItem.FNeedCert 			= rsget("needCert")
		End If
		rsget.Close
	End Sub

	Public Sub getNvClassEditOneItem
		Dim strSql, addSql, i
		If FRectItemID <> "" Then
			addSql = " and i.itemid in (" & FRectItemID & ")"
		End If
		strSql = ""
		strSql = strSql & " SELECT TOP " & FPageSize & " i.* "
		strSql = strSql & "	, c.keywords, c.ordercomment, c.sourcearea, c.makername, c.usingHTML, c.itemcontent, isNULL(c.requireMakeDay,0) as requireMakeDay "
		strSql = strSql & "	, isNULL(m.nvClassGoodNo, '') as nvClassGoodNo, m.nvClassprice, m.nvClassSellYn, isNULL(m.regedOptCnt, 0) as regedOptCnt "
		strSql = strSql & "	, m.accFailCNT, m.lastErrStr, isNULL(m.regitemname,'') as regitemname, m.regImageName "
		strSql = strSql & "	, isnull(m.APIaddImg, '') as APIaddImg "
		strSql = strSql & "	, C.infoDiv, isNULL(C.safetyyn,'N') as safetyyn, isNULL(C.safetyDiv,0) as safetyDiv, C.safetyNum, uc.socname_kor "
    	strSql = strSql & "	,(CASE WHEN i.isusing = 'N' "
		strSql = strSql & "		or i.sellyn <> 'Y'"
		strSql = strSql & "		or i.deliverfixday in ('C','X')"
		strSql = strSql & "		or i.itemdiv <> '08' or i.cate_large = '999' or i.cate_large=''"
		strSql = strSql & " 	or exists(SELECT top 1 n.itemid FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid n with (nolock) WHERE n.itemid=i.itemid and n.mallgubun = 'nvstorefarmclass') "
		strSql = strSql & " 	or i.cate_large + i.cate_mid + i.cate_small <> '035022010' "
		strSql = strSql & "	THEN 'Y' ELSE 'N' END) as maySoldOut "
		strSql = strSql & " FROM db_item.dbo.tbl_item as i "
		strSql = strSql & " JOIN db_item.dbo.tbl_item_contents as c on i.itemid = c.itemid "
		strSql = strSql & " JOIN db_etcmall.dbo.tbl_nvstorefarmclass_regItem as m on i.itemid = m.itemid "
		strSql = strSql & " LEFT JOIN db_user.dbo.tbl_user_c uc on i.makerid = uc.userid"
		strSql = strSql & " WHERE 1 = 1"
		strSql = strSql & " and m.APIaddImg = 'Y' "
		strSql = strSql & " and m.nvClassStatCD = 7 "
		strSql = strSql & addSql
		strSql = strSql & " and m.nvClassGoodNo is Not Null "									'#등록 상품만
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		FResultCount = rsget.RecordCount
		FTotalCount = FResultCount
		If not rsget.EOF Then
			Set FOneItem = new CNvClassItem
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
				FOneItem.FNvClassGoodNo		= rsget("nvClassGoodNo")
				FOneItem.FNvClassprice		= rsget("nvClassprice")
				FOneItem.FNvClassSellyn		= rsget("nvClassSellYn")

	            FOneItem.FoptionCnt         = rsget("optionCnt")
	            FOneItem.FregedOptCnt       = rsget("regedOptCnt")
	            FOneItem.FaccFailCNT        = rsget("accFailCNT")
	            FOneItem.FlastErrStr        = rsget("lastErrStr")
	            FOneItem.Fdeliverytype      = rsget("deliverytype")
				FOneItem.FSocname_kor		= rsget("socname_kor")
	            FOneItem.FrequireMakeDay    = rsget("requireMakeDay")

	            FOneItem.FinfoDiv			= rsget("infoDiv")
	            FOneItem.Fsafetyyn			= rsget("safetyyn")
	            FOneItem.FsafetyDiv			= rsget("safetyDiv")
	            FOneItem.FsafetyNum			= rsget("safetyNum")
	            FOneItem.FmaySoldOut		= rsget("maySoldOut")
	            FOneItem.Fregitemname		= rsget("regitemname")
                FOneItem.FregImageName		= rsget("regImageName")
                FOneItem.FbasicImageNm		= rsget("basicimage")
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

Function getNvClassGoodNo(iitemid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT TOP 1 nvClassGoodNo FROM db_etcmall.[dbo].[tbl_nvstorefarmclass_regItem] WHERE itemid = '"&iitemid&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		getNvClassGoodNo = rsget("nvClassGoodNo")
	rsget.Close
End Function

%>
