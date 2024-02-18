<%
'상품리스트 출력 함수
Const CDEFALUT_STOCK = 99999
Const CaccessKey = "KPLXP4DS46H77W8LSYD1"
Const CsecretKey = "HFAXZZ5DMYKJ7PQBR9O6"

'/키인증	'//2015.05.07 한용민 생성
Function getkey_confirm(byref ref_accessKey, byref ref_secretKey, byref ref_Key_str)
	dim tmpkey_str, tmpkey_confirm
	tmpkey_str=""
	tmpkey_confirm=FALSE

	on Error Resume Next
	HeadAuthorization = Request.ServerVariables("HTTP_Authorization")

	if HeadAuthorization="" then
		tmpkey_str = "키값이 없습니다."
		tmpkey_confirm=FALSE
	end if
	If Ubound(Split(HeadAuthorization,":")) > 1 Then
		tmpkey_str = "키값이 정확하지 않거나, 구분자가(:) 중복됩니다."
		tmpkey_confirm=FALSE
	Else
		ref_accessKey = Trim(Replace(Split(HeadAuthorization,":")(0), "TenByTen", ""))
		ref_secretKey = Split(HeadAuthorization,":")(1)

		if CaccessKey <> ref_accessKey or CsecretKey <> ref_secretKey then
			tmpkey_str = "키값이 일치하지 않습니다."
			tmpkey_confirm=FALSE
		else
			tmpkey_str = ""
			tmpkey_confirm = TRUE
		End If
	End If
	On Error Goto 0

	ref_Key_str = tmpkey_str
	getkey_confirm = tmpkey_confirm
End Function

'/통신후 에러코드 설명서		'//2015.05.07 한용민 생성
function getresult_Status_str(ref_Status)
	dim tmpstr
	if ref_Status="" then
		getresult_Status_str="결과코드가 없습니다."
		exit function
	end if

	if ref_Status="200" then
		tmpstr=""		'/정상통신
	elseif ref_Status="400" then
		tmpstr="잘못된 형태의 토큰값이거나, 유효기간이 만료된 토큰값 입니다."
	elseif ref_Status="401" then
		tmpstr="잘못된 형태의 헤더 인증키 값입니다."
	elseif ref_Status="405" then
		tmpstr="잘못된 형태의 메소드값 입니다."
	elseif ref_Status="500" then
		tmpstr="서버측 버그가 발생 되었습니다."
	else
		tmpstr="알수없는 오류가 발생 되었습니다."
	end if

	getresult_Status_str = tmpstr
end function

'/비트윈 아이디 가져오기	'//2015.05.01 김진영 생성	'/2015.05.07 한용민 수정(상태값 파라메타 반환)
Function getBetweenID(iToken, byref ref_Status, byref ref_result_str)
	Dim betweenAPIURL, objXML, iRbody, jsResult, strParam, i
	IF application("Svr_Info")="Dev" THEN
		betweenAPIURL = "https://commerce.mintnote.com/10x10/users"
	else
		betweenAPIURL = "https://commerce.mintnote.com/10x10/users"
	end if
	on Error Resume Next
	'iToken ="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJwYXlsb2FkIjoie1wiaWRcIjogMTQ1LCBcInVzZXJfaWRcIjogXCI3M21jRXNmMlwifSIsImV4cCI6MTQzMjAxNzM0NzEzM30.NzBOMSNKdOj3gnVOgkbDwCkVSvilmAji9VDqoFf5tOU"
	strParam = "token="&iToken
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objXML.Open "POST", betweenAPIURL , False
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.SetRequestHeader "Authorization","TenByTen "& CaccessKey &":"& CsecretKey &""
		objXML.Send(strParam)
		iRbody = BinaryToText(objXML.ResponseBody,"euc-kr")
		ref_Status = objXML.Status
		ref_result_str = getresult_Status_str(objXML.Status)
		If objXML.Status = "200" Then
			SET jsResult = JSON.parse(iRbody)
				If jsResult.value <> "" Then
					getBetweenID = jsResult.value
				ElseIf jsResult.error <> "" Then
					getBetweenID = jsResult.error
				End If
			SET jsResult = nothing
		Else
			getBetweenID = ""
		End If
	Set objXML = Nothing
	On Error Goto 0
End Function

'/비트윈 아이디로 usersn 가져오기	'//2015.05.01 김진영 생성
Function getTenUserSn(betID, byref usersn)
	Dim strSql, notTableData
	strSql = ""
	strSql = strSql & " SELECT TOP 1 userSn FROM db_etcmall.dbo.tbl_between_userinfo WHERE btwUserCd = '"&betID&"' "
	'response.write strSql & "<br>"
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	'만약 db_etcmall.dbo.tbl_between_userinfo 여기의 btwUserCd와 betID가 같을 경우 셀렉트 후 userSn 받기..
	If rsget.RecordCount > 0 Then
		usersn	= rsget("userSn")
		notTableData = False
	Else
		notTableData = True
	End If
	rsget.Close
	'만약 db_etcmall.dbo.tbl_between_userinfo 여기의 btwUserCd와 betID가 다를 경우 인서트 후 userSn 받기..
	If notTableData Then
		strSql = ""
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_between_userinfo " & VBCRLF
		strSql = strSql & " (btwUserCd, regdate) VALUES " & VBCRLF
		strSql = strSql & " ('" & betID & "', getdate()) " & VBCRLF

		'response.write strSql & "<br>"
		dbget.Execute(strSql)

		strSql = ""
		strSql = strSql & " SELECT IDENT_CURRENT('db_etcmall.dbo.tbl_between_userinfo') as userSn "

		'response.write strSql & "<br>"
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not Rsget.Eof Then
			usersn = rsget("userSn")
		End If
		rsget.close
	End If
End Function

'/장바구니 삭제		'//2015.05.01 김진영 생성
Function DeleteBaguniData(usn)
	Dim strSql
	strSql = ""
	strSql = strSql & " DELETE FROM [db_my10x10].[dbo].tbl_my_baguni " & VBCRLF
	strSql = strSql & " WHERE userKey = '"&usn&"' " & VBCRLF
	dbget.Execute strSql, 1
End Function

'/장바구니 업데이트		'//2015.05.01 김진영 생성
Function UpdateChkOrderYBaguniData(usn)
	Dim strSql
	strSql = ""
	strSql = strSql & " UPDATE [db_my10x10].[dbo].tbl_my_baguni SET " & VBCRLF
	strSql = strSql & " chkOrder = 'Y' " & VBCRLF
	strSql = strSql & " WHERE userKey = '"&usn&"' " & VBCRLF
	dbget.Execute strSql, 1
End Function

'/상품리스트 가져오기	'/2015.05.01 김진영 생성
Function fnBetweenItemlistJsonFlush(ioffset, ilimit)
	Dim strSql, ArrRows, oJson, i, lp, ArrTag
	Dim TItemname, TBrandname, TItemid, TBasicimage, TBasicimage600, TMaskimage, TOrgprice, TSellcash, TItemsource, TItemsize, TItemcontent
	Dim TUsinghtml, TMainimage, TMainimage2, TMainimage3, TOrdercomment, TDefaultfreebeasonglimit, TDefaultdeliverpay, TDeliverytype, TSocname_kor
	Dim TDeliveroverseas, TItemweight, TItemdiv, TMakerid, TItemscore, TKeywords, TSourcearea, TSafetyYN, TSafetyNum, Tcatecode, Tcatename, Taddcatearr
	dim Tcontentsimage400, TExplain

	strSql = "exec db_outmall.dbo.sp_Between_API_ItemList '"&ioffset&"', '"&ilimit&"' "
	rsCTget.CursorLocation = adUseClient
	rsCTget.CursorType = adOpenStatic
	rsCTget.LockType = adLockOptimistic

	'response.write strSql & "<br>"
	rsCTget.Open strSql, dbCTget
	If Not(rsCTget.Eof or rsCTget.BOF) Then
		ArrRows = rsCTget.getRows
	End If
	rsCTget.close

	if isarray(ArrRows) then
		Set oJson = jsArray()
		For i = 0 To UBound(ArrRows, 2)
			if i > UBound(ArrRows, 2) then exit For

			Set oJson(null) = jsObject()
				TItemname					= ArrRows(1,i)
				TBrandname					= ArrRows(2,i)
				TItemid						= ArrRows(3,i)
				TBasicimage					= ArrRows(4,i)
				TBasicimage600				= ArrRows(5,i)
				TMaskimage					= ArrRows(6,i)
				TOrgprice					= ArrRows(7,i)
				TSellcash					= ArrRows(8,i)
				TItemsource					= ArrRows(9,i)
				TItemsize					= ArrRows(10,i)
				TItemcontent				= ArrRows(11,i)
				TUsinghtml					= ArrRows(12,i)
				TMainimage					= ArrRows(13,i)
				TMainimage2					= ArrRows(14,i)
				TMainimage3					= ArrRows(15,i)
				TOrdercomment				= ArrRows(16,i)
				TDefaultfreebeasonglimit	= ArrRows(17,i)
				TDefaultdeliverpay			= ArrRows(18,i)
				TDeliverytype				= ArrRows(19,i)
				TSocname_kor				= ArrRows(20,i)
				TDeliveroverseas			= ArrRows(21,i)
				TItemweight					= ArrRows(22,i)
				TItemdiv					= ArrRows(23,i)
				TMakerid					= ArrRows(24,i)
				TItemscore					= ArrRows(25,i)
				TKeywords					= ArrRows(26,i)
				TSourcearea					= ArrRows(27,i)
				TSafetyYN					= ArrRows(28,i)
				TSafetyNum					= ArrRows(29,i)
				Tcatecode					= ArrRows(30,i)
				Tcatename					= ArrRows(31,i)
				Taddcatearr					= ArrRows(32,i)
				Tcontentsimage400			= ArrRows(33,i)
				TExplain					= ArrRows(34,i)

				If Not(TKeywords="" or isNull(TKeywords)) Then
					ArrTag 					= Split(TKeywords,",")
				End If
				oJson(null)("name")		= ""&TItemname&""
				oJson(null)("brand")	= ""&TBrandname&""
				oJson(null)("code")		= TItemid

				'/카테고리 출력
				Set oJson(null)("catetory") = jsArray()
				Call getbetweencategory(Tcatecode, Tcatename, Taddcatearr, oJson)	'/내부에서 쿼리안함	'//2015.08.26 한용민 추가

				'/이미지출력
				Set oJson(null)("images") = jsArray()
				Call getItemImages(TItemid, TBasicimage, TBasicimage600, TMaskimage, oJson)		'/내부에서 쿼리안함

				oJson(null)("price")	= TOrgprice
				oJson(null)("discount")	= TSellcash

				'/상품정보고시 출력
				Set oJson(null)("info") = jsArray()
				Call getItemInfoCode(TItemid, oJson, TItemsource, TItemsize, TSafetyYN, TSafetyNum, TExplain)	'/내부에서 쿼리안함

				'/상품상세 출력
				oJson(null)("explain")			= getItemContents(TItemid, TItemcontent, TUsinghtml, TMainimage, TMainimage2, TMainimage3, Tcontentsimage400)	'/내부에서 쿼리안함
				oJson(null)("order_attention")	= getDeliverNoticsStr(TBrandname, TSocname_kor, TDefaultfreebeasonglimit, TDefaultdeliverpay, TDeliverytype)	'/내부에서 쿼리안함

				'/상품 주의사항 출력
				Set oJson(null)("caution") = jsObject()
				Set oJson(null)("caution")("delivery") = jsArray()
				oJson(null)("caution")("delivery")(null) = "배송기간은 주문일(무통장입금은 결제완료일)로부터 1일(24시간) ~ 5일정도 걸립니다."
				oJson(null)("caution")("delivery")(null) = "업체배송 상품은 무료배송 되며, 업체조건배송 상품은 특정 브랜드 배송 기준으로 배송비가 부여되며 업체착불배송은 특정 브랜드 배송기준으로 고객님의 배송지에 따라 배송비가 착불로 부과됩니다."
				oJson(null)("caution")("delivery")(null) = "제작기간이 별도로 소요되는 상품의 경우에는 상품설명에 있는 제작기간과 배송시기를 숙지해 주시기 바랍니다."
				oJson(null)("caution")("delivery")(null) = "가구 등의 상품의 경우에는 지역에 따라 추가 배송비용이 발생할 수 있음을 알려드립니다."
				Set oJson(null)("caution")("A/S") = jsArray()
				oJson(null)("caution")("A/S")(null) = "상품 수령일로부터 7일 이내 반품/환불 가능합니다."
				oJson(null)("caution")("A/S")(null) = "변심 반품의 경우 왕복배송비를 차감한 금액이 환불되며, 제품 및 포장 상태가 재판매 가능하여야 합니다."
				oJson(null)("caution")("A/S")(null) = "상품 불량인 경우는 배송비를 포함한 전액이 환불됩니다."
				oJson(null)("caution")("A/S")(null) = "출고 이후 환불요청 시 상품 회수 후 처리됩니다."
				oJson(null)("caution")("A/S")(null) = "주문제작(쥬얼리 포함)/카메라/밀봉포장상품/플라워 등은 변심으로 반품/환불 불가합니다."
				oJson(null)("caution")("A/S")(null) = "완제품으로 수입된 상품의 경우 A/S가 불가합니다."
				oJson(null)("caution")("A/S")(null) = "특정브랜드의 교환/환불/AS에 대한 개별기준이 상품페이지에 있는 경우 브랜드의 개별기준이 우선 적용 됩니다."
				Set oJson(null)("caution")("etc") = jsArray()
				oJson(null)("caution")("etc")(null) = "구매자가 미성년자인 경우에는 상품 구입시 법정대리인이 동의하지 아니하면 미성년자 본인 또는 법정대리인이 구매취소 할 수 있습니다."

				'/배송 주의 사항 시작
				Set oJson(null)("delivery")	= jsArray()
				Set oJson(null)("delivery") = jsObject()
				If IsAboardBeasong(TDeliveroverseas, TItemweight, TDeliverytype) Then
					If IsFreeBeasong(TSellcash, TDefaultfreebeasonglimit, TDeliverytype) Then
						oJson(null)("delivery")("type") = "텐바이텐 무료배송+해외배송"
					Else
						oJson(null)("delivery")("type") = "텐바이텐배송+해외배송"
					End If
						oJson(null)("delivery")("info") = "텐바이텐에서는 구매한 상품을 해외 친구나 친지들이 받아 보실 수 있도록 해외배송 서비스(항공편 이용)를 신설하여 운영을 시작합니다.<br /><br />해외배송을 대행해줄 곳은 국가기관인 우정사업본부이며, 개인적으로 우체국을 통하여 해외배송 서비스를 받을때 보다 편리하게 이용하실 수 있습니다.<br /><br />EMS(Express Mail Service) 는 전세계 59개국(계속 확대중)으로 배송하며, 외국 우편당국과 체결한 특별협정에 따라 취급합니다."
				ElseIf TItemdiv = "09" Then
					oJson(null)("delivery")("type") = GetDeliveryName(TDeliverytype, TSellcash, TDefaultfreebeasonglimit, TMakerid)
					oJson(null)("delivery")("info") = "해당 상품은 10X10 Present 상품으로 주문 건당 2,000원의 배송비가 부과됩니다."
				Else
					oJson(null)("delivery")("type") = GetDeliveryName(TDeliverytype, TSellcash, TDefaultfreebeasonglimit, TMakerid)
					oJson(null)("delivery")("info") = ""
				End If

				oJson(null)("priority")		= TItemscore

				'/키워드 시작
				Set oJson(null)("tags")	= jsArray()
				If IsArray(arrTag) Then
					For lp = 1 to Ubound(ArrTag)
						If Trim(ArrTag(lp)) <> "" Then
							oJson(null)("tags")(null) = ArrTag(lp)
						End If
					Next
				Else
					oJson(null)("tags")		 = ""
				End If

				oJson(null)("from")			= ""&TSourcearea&""
		Next
		'oJson(null)("end")			= "end"
		oJson.Flush
	end if
End Function

'/비트윈 카테고리	'//2015.08.26 한용민 생성
Function getbetweencategory(Tcatecode, Tcatename, Taddcatearr, byref oJson)
	dim tempcatearr

	'/기본카테고리
	If not isNull(Tcatecode) and Tcatecode <> "" Then
		oJson(null)("catetory")(null) = "[기본카테고리]["& Tcatecode &"]["& Tcatename &"]"
	End If

	'/현재 추가 카테고리 1까지만 사용하고 있으나 어드민 무한 증가로 카테고리 등록이 가능하도록 개발완료가 되어 있음
	'/현재는 디비 아낄려고 쿼리 안하지만, 차후 우리측에서 추가 카테고리가 더 늘어 난다면 프로시져에서 구분자 추가하고 데이터 가져와서 구분자로 짤라서 밑에 이어 붙일것
	'/추가카테고리
	If not isNull(Taddcatearr) and Taddcatearr <> "" Then
		tempcatearr = split(Taddcatearr,"!@#")

		if isarray(tempcatearr) then
			oJson(null)("catetory")(null) = "[추가카테고리]["& tempcatearr(0) &"]["& tempcatearr(1) &"]"
		end if
	End If
End Function

'해당 상품이미지 출력 함수
Function getItemImages(iitemid, ibasicImage, ibasicImage600, imaskImage, byref oJson)
	Dim strSql, iRows, lp, vImage, lp2

'	strSql = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & CStr(iitemid)
'	rsget.Open strSql, dbget
'	If Not rsget.EOF Then
'		iRows = rsget.GetRows
'	End if
'	rsget.Close
'	If isArray(iRows) Then
'		For lp = 0 to Ubound(iRows, 2)
'			If isNull(iRows(3, lp)) OR iRows(3, lp) = "" Then
'				vImage = iRows(2, lp)
'			Else
'				vImage = iRows(3, lp)
'			End IF
'
'			IF iRows(1, lp)="1" Then
'				oJson(null)("images")(null) = "http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(iitemid) & "/" & vImage
'			Else
'				oJson(null)("images")(null) = "http://webimage.10x10.co.kr/item/add" & GetImageSubFolderByItemid(iitemid) & "/" & vImage
'			End IF
'		Next
'	End If

	If isNull(ibasicImage600) OR ibasicImage600 = "" Then
		If NOT isNull(ibasicImage) OR ibasicImage = "" Then
			oJson(null)("images")(null) = "http://webimage.10x10.co.kr/image/basic/" + GetImageSubFolderByItemid(iitemid) + "/" + ibasicImage
		End if
	Else
		oJson(null)("images")(null) = "http://webimage.10x10.co.kr/image/basic600/" + GetImageSubFolderByItemid(iitemid) + "/" + ibasicImage600
	End If

	If Not(isNull(imaskImage) OR imaskImage = "") Then
		oJson(null)("images")(null) = "http://webimage.10x10.co.kr/image/mask/" + GetImageSubFolderByItemid(iitemid) + "/" + imaskImage
	End If
End Function

Function getItemInfoCode(iitemid, byref oJson, iitemsource, iitemsise, isafetyYN, isafetyNum, iExplain)
	Dim strSql, iRows, lp, lj, tempExplainarr1, tempExplainarr2, tempExplainarr3

'	strSql = "exec [db_item].[dbo].[sp_Ten_CategoryPrd_AddExplain] " & CStr(iitemid)
'
'	'response.write strSql & "<br>"
'	rsget.CursorLocation = adUseClient
'	rsget.CursorType=adOpenStatic
'	rsget.Locktype=adLockReadOnly
'	rsget.Open strSQL, dbget
'	If Not rsget.EOF Then
'		iRows 	= rsget.GetRows
'	End if
'	rsget.close
'	If isArray(iRows) Then
'		For lp = 0 to Ubound(iRows, 2)
'			if lp > UBound(iRows, 2) then exit For
'
'			Set oJson(null)("info")(null) = jsObject()
'			If iRows(2, lp) = "35005" Then
'				If iitemsource <> "" Then
'					oJson(null)("info")(null)("name") = "재질"
'					oJson(null)("info")(null)("value") = iitemsource
'				End If
'
'				If iitemsise <> "" Then
'					oJson(null)("info")(null)("name") = "사이즈"
'					oJson(null)("info")(null)("value") = iitemsise
'				End If
'			End If
'			oJson(null)("info")(null)("name") = iRows(0, lp)
'			oJson(null)("info")(null)("value") = iRows(1, lp)
'		Next
'		If isafetyYN = "Y" Then
'			oJson(null)("info")(null)("name") = "안전인증대상"
'			oJson(null)("info")(null)("value") = isafetyNum
'		End If
'	End If

	If not isNull(iExplain) and iExplain <> "" Then
		tempExplainarr1 = split(iExplain,"!@#")

		if isarray(tempExplainarr1) then
			for lp = 0 to ubound(tempExplainarr1)-1
				tempExplainarr2 = tempExplainarr1(lp)
				tempExplainarr3 = split(tempExplainarr2,"|^|")

				if isarray(tempExplainarr3) then
					if ubound(tempExplainarr3)>1 then
						Set oJson(null)("info")(null) = jsObject()
						If tempExplainarr3(2) = "35005" Then
							If iitemsource <> "" Then
								oJson(null)("info")(null)("name") = "재질"
								oJson(null)("info")(null)("value") = iitemsource
							End If

							If iitemsise <> "" Then
								oJson(null)("info")(null)("name") = "사이즈"
								oJson(null)("info")(null)("value") = iitemsise
							End If
						End If
						oJson(null)("info")(null)("name") = tempExplainarr3(0)
						oJson(null)("info")(null)("value") = tempExplainarr3(1)
					end if
				end if
			next

			If isafetyYN = "Y" Then
				oJson(null)("info")(null)("name") = "안전인증대상"
				oJson(null)("info")(null)("value") = isafetyNum
			End If
		end if
	End If
End Function

Function getItemContents(iitemid, icontents, iusingHtml, imainimage, imainimage2, imainimage3, contentsimage400)
	Dim strRst, strSQL, tempcontentsimagearr, lp

	strRst = ("<div align=""center"">")
	'#기본 상품설명
	Select Case iusingHtml
		Case "Y"
			strRst = strRst & (icontents & "<br>")
		Case "H"
			strRst = strRst & (nl2br(icontents) & "<br>")
		Case Else
			strRst = strRst & (nl2br(ReplaceBracket(icontents)) & "<br>")
	End Select

'	'# 추가 상품 설명이미지 접수
'	strSQL = "exec [db_item].[dbo].sp_Ten_CategoryPrd_AddImage @vItemid =" & iitemid
'	rsget.CursorLocation = adUseClient
'	rsget.CursorType=adOpenStatic
'	rsget.Locktype=adLockReadOnly
'	rsget.Open strSQL, dbget
'	If Not(rsget.EOF or rsget.BOF) Then
'		Do Until rsget.EOF
'			If rsget("imgType") = "1" Then
'				strRst = strRst & ("<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(iitemid) & "/" & rsget("addimage_400") & """ border=""0"" style=""width:100%""><br>")
'			End If
'			rsget.MoveNext
'		Loop
'	End If
'	rsget.Close

	If not isNull(contentsimage400) and contentsimage400 <> "" Then
		tempcontentsimagearr = split(contentsimage400,"!@#")

		if isarray(tempcontentsimagearr) then
			for lp = 0 to ubound(tempcontentsimagearr)-1
			strRst = strRst & "<img src=""http://webimage.10x10.co.kr/item/contentsimage/" & GetImageSubFolderByItemid(iitemid) & "/" & tempcontentsimagearr(lp) & """ border=""0"" style=""width:100%""><br>"
			next
		end if
	End If

	'#기본 상품 설명이미지
	If ImageExists(imainimage) Then strRst = strRst & ("<img src=""" & imainimage & """ border=""0"" style=""width:100%""><br>")
	If ImageExists(imainimage2) Then strRst = strRst & ("<img src=""" & imainimage2 & """ border=""0"" style=""width:100%""><br>")
	If ImageExists(imainimage3) Then strRst = strRst & ("<img src=""" & imainimage3 & """ border=""0"" style=""width:100%""><br>")

	getItemContents = strRst
End Function

'// 상품이미지 존재여부 검사
Function ImageExists(byval iimg)
	If (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") Then
		ImageExists = false
	Else
		ImageExists = true
	End If
End Function

''// 업체별 배송비 부과 상품(업체 조건 배송)
Public Function IsUpcheParticleDeliverItem(qdefaultfreebeasonglimit, qdefaultdeliverpay, qdeliveryType)
	IsUpcheParticleDeliverItem = (qdefaultfreebeasonglimit>0) and (qdefaultdeliverpay>0) and (qdeliveryType="9")
End Function

''// 업체착불 배송여부
Public Function IsUpcheReceivePayDeliverItem(qdeliveryType)
	IsUpcheReceivePayDeliverItem = (qdeliveryType="7")
End Function

Public Function getDeliverNoticsStr(ibrandname, ibrandname_kor, idefaultfreebeasonglimit, idefaultdeliverpay, ideliveryType)
	getDeliverNoticsStr = ""
	If (IsUpcheParticleDeliverItem(idefaultfreebeasonglimit, idefaultdeliverpay, ideliveryType)) then
		getDeliverNoticsStr = ibrandname & "(" & ibrandname_kor & ") 제품으로만" & "<br>"
		getDeliverNoticsStr = getDeliverNoticsStr & FormatNumber(idefaultfreebeasonglimit,0) & "원 이상 구매시 무료배송 됩니다."
		getDeliverNoticsStr = getDeliverNoticsStr & "배송비(" & FormatNumber(idefaultdeliverpay,0) & "원)"
	ElseIf (IsUpcheReceivePayDeliverItem(ideliveryType)) then
		getDeliverNoticsStr = "착불 배송비는 지역에 따라 차이가 있습니다. "
		getDeliverNoticsStr = getDeliverNoticsStr & " 상품설명의 '배송안내'를 꼭 읽어보세요." & "<br>"
	End If
End Function

'// 해외 배송 여부(텐배 + 해외여부 + 상품무게)
public Function IsAboardBeasong(ideliverOverseas, iitemWeight, ideliverytype)
	If ideliverOverseas="Y" and iitemWeight>0 and (ideliverytype="1" or ideliverytype="3" or ideliverytype="4") then
		IsAboardBeasong = true
	Else
		IsAboardBeasong = false
	End If
end function

'// 무료 배송 여부
public Function IsFreeBeasong(isellcash, idefaultfreebeasonglimit, ideliverytype)
	if (cLng(isellcash) >= cLng(getFreeBeasongLimitByUserLevel(idefaultfreebeasonglimit, ideliverytype))) then
		IsFreeBeasong = true
	else
		IsFreeBeasong = false
	end if

	if (ideliverytype="2") or (ideliverytype="4") or (ideliverytype="5") or (ideliverytype="6") then
		IsFreeBeasong = true
	end if

	''//착불 배송은 무료배송이 아님
	if (ideliverytype="7") then
	    IsFreeBeasong = false
	end if
end Function

public Function getFreeBeasongLimitByUserLevel(qdefaultfreebeasonglimit, qdeliverytype)
	dim ulevel
	if (qdeliverytype="9") then
	    If (IsNumeric(qdefaultfreebeasonglimit)) and (qdefaultfreebeasonglimit<>0) then
	        getFreeBeasongLimitByUserLevel = qdefaultfreebeasonglimit
	    else
	        getFreeBeasongLimitByUserLevel = 50000
	    end if
	else
	    getFreeBeasongLimitByUserLevel = 30000
	end if
end Function

'// 배송구분 : 무료배송은 따로 처리  '!
public Function GetDeliveryName(ideliverytype, isellcash, idefaultfreebeasonglimit, imakerid)
	Select Case ideliverytype
		Case "1"
			if IsFreeBeasong(isellcash, idefaultfreebeasonglimit, ideliverytype) then
				GetDeliveryName="텐바이텐무료배송"
			else
				GetDeliveryName="텐바이텐배송"
			end if
		Case "2"
			if imakerid="goodovening" then
				GetDeliveryName="업체배송"
			else
				GetDeliveryName="업체무료배송"
			end if
		'Case "3"
		'		GetDeliveryName="텐바이텐배송"
		Case "4"
				GetDeliveryName="텐바이텐무료배송"
		Case "5"
				GetDeliveryName="업체무료배송"
		Case "6"
				GetDeliveryName="현장수령상품"
		Case "7"
			GetDeliveryName="업체착불배송"
		Case "9"
			if Not IsFreeBeasong(isellcash, idefaultfreebeasonglimit, ideliverytype) then
				GetDeliveryName="업체조건배송"
			else
				GetDeliveryName="업체무료배송"
			end if
		Case Else
			GetDeliveryName="텐바이텐배송"
	End Select
End Function

'/상품 옵션 가져오기	'/2015.05.01 김진영 생성
Function fnBetweenItemoptionJsonFlush(iitemid)
	Dim strSql, ArrRows, oJson, i, lp, ArrTag
	Dim optcount, ilimityn, ilimitno, ilimitsold, iitemdiv
	Dim chkMultiOpt : chkMultiOpt = False
	Dim optioncode, optTypeNm, optionname, optLimit, optaddprice
	strSql = ""
	strSql = strSql & " SELECT TOP 1 optioncnt, limityn, limitno, limitsold, itemdiv FROM db_AppWish.dbo.tbl_item where itemid = '"&iitemid&"' "
	rsCTget.Open strSql, dbCTget
	If Not rsCTget.EOF Then
		optcount	= rsCTget("optioncnt")
		ilimityn	= rsCTget("limityn")
		ilimitno	= rsCTget("limitno")
		ilimitsold	= rsCTget("limitsold")
		iitemdiv	= rsCTget("itemdiv")
	End if
	rsCTget.Close

	Set oJson = jsObject()
		Set oJson("options")	= jsArray()

		If optcount > 0 Then
			strSql = "exec [db_appwish].[dbo].sp_Ten_ItemOptionMultipleTypeList " & iitemid
		    rsCTget.CursorLocation = adUseClient
			rsCTget.CursorType = adOpenStatic
			rsCTget.LockType = adLockOptimistic
			rsCTget.Open strSql, dbCTget
			If Not(rsCTget.EOF or rsCTget.BOF) Then
				chkMultiOpt = True
				Do until rsCTget.EOF
					optTypeNm = optTypeNm & db2Html(rsCTget("optionTypeName"))
					rsCTget.MoveNext
					If Not(rsCTget.EOF) Then optTypeNm = optTypeNm & ","
				Loop
			End If
			rsCTget.Close

			If chkMultiOpt Then
				strSql = ""
				strSql = strSql & " SELECT itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_Appwish].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & iitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsCTget.Open strSql, dbCTget, 1
				If Not(rsCTget.EOF or rsCTget.BOF) Then
					Set oJson("options")(null) = jsObject()
						oJson("options")(null)("name") = optTypeNm
						Set oJson("options")(null)("values")	= jsArray()
					Do until rsCTget.EOF
						Set oJson("options")(null)("values")(null)	= jsObject()
					    optLimit = rsCTget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) then optLimit = 0
					    If (ilimityn <> "Y") Then optLimit = CDEFALUT_STOCK
						optionname	= db2Html(rsCTget("optionname"))
						optioncode	= rsCTget("itemoption")
						optaddprice	= rsCTget("optaddprice")

						oJson("options")(null)("values")(null)("code")		= optioncode
						oJson("options")(null)("values")(null)("name")		= optionname
						oJson("options")(null)("values")(null)("count")		= optLimit
						oJson("options")(null)("values")(null)("price")	= optaddprice
						rsCTget.MoveNext
					Loop
				End If
				rsCTget.Close
			Else
				strSql = ""
				strSql = strSql & " SELECT itemoption, optionTypeName, optionname, (optlimitno-optlimitsold) as optLimit, optaddprice "
				strSql = strSql & " FROM [db_Appwish].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & iitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsCTget.Open strSql, dbCTget, 1
				If Not(rsCTget.EOF or rsCTget.BOF) Then
					If db2Html(rsCTget("optionTypeName")) <> "" Then
						optTypeNm = db2Html(rsCTget("optionTypeName"))
					Else
						optTypeNm = "옵션"
					End If

					Set oJson("options")(null) = jsObject()
						oJson("options")(null)("name") = optTypeNm
						Set oJson("options")(null)("values")	= jsArray()
					Do until rsCTget.EOF
						Set oJson("options")(null)("values")(null)	= jsObject()
					    optLimit = rsCTget("optLimit")
					    optLimit = optLimit - 5
					    If (optLimit < 1) then optLimit = 0
					    If (ilimityn <> "Y") Then optLimit = CDEFALUT_STOCK
						optionname = db2Html(rsCTget("optionname"))
						optioncode	= rsCTget("itemoption")
						optaddprice	= rsCTget("optaddprice")

						oJson("options")(null)("values")(null)("code")		= optioncode
						oJson("options")(null)("values")(null)("name")		= optionname
						oJson("options")(null)("values")(null)("count")		= optLimit
						oJson("options")(null)("values")(null)("price")	= optaddprice
						rsCTget.MoveNext
					Loop
				End If
				rsCTget.Close
			End If

		Else
			oJson("options")		 = ""
		End If

		oJson("code") = iitemid
		oJson("is_text") = Chkiif(iitemdiv = "06", True, False)
		oJson("is_limit") = Chkiif(ilimityn = "Y", True, False)
		oJson("is_multiple") = chkMultiOpt

		if ilimitno - ilimitsold < 1 then
			oJson("count") = Chkiif(ilimityn = "Y", 0, CDEFALUT_STOCK)
		else
			oJson("count") = Chkiif(ilimityn = "Y", ilimitno - ilimitsold - 5, CDEFALUT_STOCK)
		end if

		oJson.Flush
End Function

Function FnURLDecode(sStr)
	Dim sRet, reEncode, sChar
	Dim i

	If isnull(sStr) Then
		sStr = ""
	Else
		Set reEncode = New RegExp
		reEncode.IgnoreCase = True
		reEncode.Pattern = "^%[0-9a-f][0-9a-f]$"
		sStr = Replace(sStr, "+", " ")
		sRet = ""

		For i = 1 To Len(sStr)
			sChar = Mid(sStr, i, 3)
			If reEncode.Test(sChar) Then
				If CInt("&H" & Mid(sStr, i + 1, 2)) < 128 Then
					sRet = sRet & Chr(CInt("&H" & Mid(sStr, i + 1, 2)))
					i = i + 2
				Elseif mid(sStr, i+3, 1) ="%" Then
					sRet = sRet & Chr(CInt("&H" & Mid(sStr, i + 1, 2) & Mid(sStr, i + 4, 2)))
					i = i + 5
				Else
					sRet = sRet & Chr(CInt("&H" & Mid(sStr, i + 1, 2) & "00") + asc(mid(sStr,i+3,1)))
					i = i + 3
					'이부분이 중요하다.
					'기존 urldecode함수가 몇 몇 글자들에서 에러를 내는 이유는 3바이트로 인코딩 되어있는 부분이 있기 때문. ascII로 변형되어 인코딩 된 부분이 존재한다.
				End If
			Else
				sRet = sRet & Mid(sStr, i, 1)
			End If
		Next
	End If
	FnURLDecode = sRet
End Function
%>