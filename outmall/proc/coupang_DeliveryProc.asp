<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/coupang/coupangItemcls.asp"-->
<!-- #include virtual="/outmall/coupang/incCoupangFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Public Function fnDeliveryReg(iMakerid, iMaeipdiv, iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, isRegYn, strObj, phoneNumner, phoneNumner2
    isRegYn = "N"
	If iMaeipdiv = "U" Then
	    istrParam = "makerID="&iMakerid
		'/////// 우리DB에 일단 저장.. 누락 분이 있다면 통신하지 말고 에러처리 ///////
		strSql = "EXEC [db_etcmall].[dbo].[usp_API_Coupang_deliveryInfo_Add] '"&iMakerid&"' "
		dbget.Execute strSql

		'////// 우편번호에 하이픈(-)이 있으면 수정 처리
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " SET returnZipCode =  "
		strSql = strSql & " 	Case WHEN charindex('-',returnZipCode) > 0 THEN replace(returnZipCode, '-', '')  "
		strSql = strSql & " 	ELSE returnZipCode END "
		strSql = strSql & " WHERE makerid = '"& iMakerid &"' "
		dbget.Execute strSql

		strSql = ""
		strSql = strSql & " SELECT top 1 len(companyContactNumber) as phoneNumner, len(phoneNumber2) as phoneNumner2 "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " WHERE makerid = '"&iMakerid&"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			phoneNumner = rsget("phoneNumner")
			phoneNumner2 = rsget("phoneNumner2")
		End If
		rsget.Close

		If (phoneNumner < 11) or (phoneNumner > 14) Then
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 전화번호 길이 오류 11~14자리 요망"
			Exit Function
		End If

		If (phoneNumner2 < 11) or (phoneNumner2 > 14) Then
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 전화번호 길이 오류2 11~14자리 요망"
			Exit Function
		End If

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as cnt "
		strSql = strSql & " FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " WHERE makerid = '"&iMakerid&"' "
		strSql = strSql & " and isNull(companyContactNumber, '') <> '' "
		strSql = strSql & " and isNull(phoneNumber2, '') <> '' "
		strSql = strSql & " and isNull(returnZipCode, '') <> '' "
		strSql = strSql & " and isNull(returnAddress, '') <> '' "
		strSql = strSql & " and isNull(returnAddressDetail, '') <> '' "
		strSql = strSql & " and isNull(deliveryCode, '') <> '' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("cnt") > 0 Then
			isRegYn = "Y"
		End If
		rsget.Close
		'//////////////////////////////////////////////////////////////////////

		If isRegYn = "Y" Then
			On Error Resume Next
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
				objXML.open "POST", "http://xapi.10x10.co.kr:8080/Deliveries/Coupang/origin", false
				objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				objXML.Send(istrParam)

				If Err.number <> 0 Then
					iErrStr = "ERR||"&iMakerid&"||실패[출고지] " & Err.Description
					Exit Function
				End If
'rw BinaryToText(objXML.ResponseBody,"utf-8")
				If objXML.Status = "200" OR objXML.Status = "201" Then
					iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
					Set strObj = JSON.parse(iRbody)
						'rw strObj.outboundShippingPlaceCode 이걸로 DB업데이트 하려했는 데, 이미 API서버에서 구현한듯..
						iErrStr = "OK||"&iMakerid&"||성공[출고지]"
					Set strObj = nothing
				Else
					iErrStr = "ERR||"&iMakerid&"||실패[출고지] 통신오류"
				End If
			Set objXML = nothing
		Else
			iErrStr = "ERR||"&iMakerid&"||실패[출고지] 정보누락"
		End If
	Else		'매입 or 특정이라면 출고지는 도봉물류로 통일
		strSql = ""
		strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.dbo.tbl_coupang_branddelivery_mapping WHERE makerid='"&iMakerid&"' )"
		strSql = strSql & " BEGIN "
		strSql = strSql & " INSERT INTO db_etcmall.dbo.tbl_coupang_branddelivery_mapping "
		strSql = strSql & " (makerid, vendorId, deliveryCode, companyContactNumber, notJeju, outboundShippingPlaceCode, regdate ) VALUES "
		strSql = strSql & " ('"&iMakerid&"', '', 'HANJIN', '1644-6035', '3000', '122412', getdate()) END "
		dbget.Execute strSql
		iErrStr = "OK||"&iMakerid&"||성공[출고지]"
	End If
End Function

Dim maeipdiv, makerid, iErrStr, failCnt
Dim resultCode, lastErrMsg, strSql
makerid = request("makerid")
maeipdiv = fnBrandmaeipdiv(makerid)
Call fnDeliveryReg(makerid, maeipdiv, iErrStr)
If iErrStr <> "" Then
	resultCode = Split(iErrStr, "||")(0)
	lastErrMsg = Split(iErrStr, "||")(2)
	failCnt = 0
	If resultCode = "ERR" Then
		failCnt = 1
	End If

	strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_coupang_delivery_LOG] WHERE makerid='"&makerid&"' )"
	strSql = strSql & " 	BEGIN "
	strSql = strSql & " 		INSERT INTO db_etcmall.[dbo].[tbl_coupang_delivery_LOG] (makerid, lastErrMsg, resultCode, regdate, failCnt) VALUES "
	strSql = strSql & " 		('"& makerid &"', '"& lastErrMsg &"', '"& resultCode &"', getdate(), '"& failCnt &"') "
	strSql = strSql & " 	END "
	strSql = strSql & " ELSE "
	strSql = strSql & " 	BEGIN "
	strSql = strSql & " 		UPDATE db_etcmall.[dbo].[tbl_coupang_delivery_LOG] SET "
	If failCnt = 0 Then
		strSql = strSql & " 		failCnt = 0 "
	Else
		strSql = strSql & " 		failCnt = failCnt + 1 "
	End If
	strSql = strSql & " 		WHERE makerid = '"&makerid&"' "
	strSql = strSql & " 	END "
	dbget.Execute strSql
	rw lastErrMsg
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->