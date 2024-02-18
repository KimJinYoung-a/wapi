<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 기본정보 등록
Public Function fnhalfclubItemReg(iitemid, istrParam, byRef iErrStr, imustprice, iimageNm)
    Dim objXML, xmlDOM, strSql, iResult, halfclubGoodno, i, tenOptCnt
    Dim iRbody, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Goods/Goods.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPMethodName", "Set_GoodsRegister"
		objXML.send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품등록] " & Err.Description
			Exit Function
		End If

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
'response.write replace(objXML.responseText, "utf-8","euc-kr")
'response.end
			iResult = xmlDOM.getElementsByTagName("ResultCode").Item(0).Text
			If iResult = "0000" Then
				halfclubGoodno = xmlDOM.getElementsByTagName("PCode").Item(0).Text
				strSql = ""
				strSql = strSql & " UPDATE R" & VbCrlf
				strSql = strSql & " SET HalfClubRegdate = getdate()" & VbCrlf
				If (halfclubGoodno <> "") Then
				    strSql = strSql & "	, HalfClubStatCd = '7'"& VbCRLF
				End If
				strSql = strSql & " ,HalfClubGoodNo = '" & halfclubGoodno & "'" & VbCrlf
				strSql = strSql & " ,HalfClublastupdate = getdate()"
				strSql = strSql & " ,HalfClubPrice = '"&imustprice&"' " & VbCrlf
				strSql = strSql & " ,HalfClubsellYn = 'Y' "& VbCrlf
				strSql = strSql & " ,accFailCNT = 0" & VbCrlf                		'실패회수 초기화
				strSql = strSql & " ,regimageName = '"&iimageNm&"'"& VbCrlf
				strSql = strSql & " FROM db_etcmall.dbo.tbl_halfclub_regitem R" & VbCrlf
				strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
				strSql = strSql & " where R.itemid = " & iitemid
				dbget.execute strSql

				strSql = ""
				strSql = strSql &  "SELECT count(*) as cnt "
				strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
				strSql = strSql & " WHERE itemid=" & iitemid
				strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
				rsget.Open strSql,dbget,1
					tenOptCnt = rsget("cnt")
				rsget.Close

				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_halfclub_regitem SET "
				strSql = strSql & " regedOptCnt = " & tenOptCnt
				strSql = strSql & " WHERE itemid = " & iitemid
				dbget.Execute strSql
				iErrStr =  "OK||"&iitemid&"||등록성공(상품등록)"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("ResultMsg").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품등록)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 수정
Public Function fnhalfclubItemEdit(iitemid, ihalfclubGoodNo, iErrStr, istrParam, imustprice, iItemName, ichgSellyn, ichgImageNm)
    Dim objXML, xmlDOM, strSql, iResult, halfclubGoodno, i, sellStatStr
    Dim iRbody, iMessage, ocount
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Goods/Goods.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPMethodName", "Set_GoodsRegister"
		objXML.send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품수정] " & Err.Description
			Exit Function
		End If

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
' if session("ssBctID")="kjy8517" then
' 	response.write replace(objXML.responseText, "utf-8","euc-kr")
' 	response.write replace(objXML.responseText, "?xml","aaaass")
' end if
'response.end
			iResult = xmlDOM.getElementsByTagName("ResultCode").Item(0).Text
			If iResult = "0000" Then
				strSql = ""
				strSql = strSql & " UPDATE R " & VbCrlf
				strSql = strSql & " SET HalfClublastupdate = getdate()" & VbCrlf
				strSql = strSql & " ,accFailCNT=0" & VbCrlf
				strSql = strSql & " ,HalfClubSellYn = '" & ichgSellYn & "'" & VbCRLF
				strSql = strSql & " ,HalfClubPrice = '" & imustprice & "'" & VbCRLF
				strSql = strSql & " ,regitemname = '"&iItemName&"' " & VbCRLF
				If (ichgImageNm <> "N") Then
					strSql = strSql & " ,regimageName='"&ichgImageNm&"'"& VbCrlf
				End If
				strSql = strSql & " from db_etcmall.dbo.tbl_halfclub_regitem R" & VbCrlf
				strSql = strSql & " JOIN db_item.dbo.tbl_item i on R.itemid = i.itemid" & VbCrlf
				strSql = strSql & " WHERE R.itemid = " & iitemid
				dbget.execute strSql

				If ichgSellYn = "Y" Then
					strSql = ""
					strSql = strSql &  "SELECT count(*) as cnt "
					strSql = strSql & " FROM [db_item].[dbo].tbl_item_option "
					strSql = strSql & " WHERE itemid=" & iitemid
					strSql = strSql & " and isUsing='Y' and optsellyn='Y' "
					rsget.Open strSql,dbget,1
						ocount = rsget("cnt")
					rsget.Close

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_halfclub_regItem SET "
					strSql = strSql & " regedOptCnt = " & ocount
					strSql = strSql & " WHERE itemid = " & iitemid
					dbget.Execute strSql
				End If

				If ichgSellYn = "N" Then
					iErrStr = "OK||"&iitemid&"||품절처리"
				Else
					iErrStr = "OK||"&iitemid&"||수정성공"
				End If
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("ResultMsg").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품수정)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

'상품 가격 변경
Public Function fnhalfclubItemEditPrice(iitemid, ihalfclubGoodNo, iErrStr, istrParam, imustprice)
    Dim objXML, xmlDOM, strSql, iResult, halfclubGoodno, i
    Dim iRbody, iMessage
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Goods/Goods.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPMethodName", "Set_Good_Price_Change"
		objXML.send(istrParam)
		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||[상품가격] " & Err.Description
			Exit Function
		End If

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
response.write replace(objXML.responseText, "utf-8","euc-kr")
response.end
			iResult = xmlDOM.getElementsByTagName("ResultCode").Item(0).Text
			If iResult = "0000" Then
			    strSql = ""
    			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_halfclub_regitem " & VbCRLF
    			strSql = strSql & "	SET HalfClublastupdate = getdate() " & VbCRLF
    			strSql = strSql & "	,HalfClubPrice = " & imustprice & VbCRLF
    			strSql = strSql & "	,accFailCnt = 0"& VbCRLF
    			strSql = strSql & " Where itemid='" & iitemid & "'"& VbCRLF
				dbget.execute strSql
				iErrStr =  "OK||"&iitemid&"||수정성공(상품가격)"
			Else
				iMessage = replaceMsg(xmlDOM.getElementsByTagName("ResultMsg").item(0).text)
				iErrStr = "ERR||"&iitemid&"||"&iMessage&"(상품가격)"
			End If
		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0
End Function

Public Function fnhalfclubDeliveryCode()
    Dim objXML, xmlDOM, strSql, iResult, istrParam
    Dim iRbody, iMessage
	istrParam = ""
	istrParam = istrParam & "<?xml version=""1.0"" encoding=""utf-8""?>"
	istrParam = istrParam & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	istrParam = istrParam & "  <soap:Body>"
	istrParam = istrParam & "    <Get_DeliveryAgencyList xmlns=""http://api.tricycle.co.kr/"" />"
	istrParam = istrParam & "  </soap:Body>"
	istrParam = istrParam & "</soap:Envelope>"

	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.open "POST", "" & APIURL&"/Delivery/Delivery.asmx"
		objXML.setRequestHeader "Host", "api.tricycle.co.kr"
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", LenB(istrParam)
		objXML.setRequestHeader "SOAPMethodName", "Get_DeliveryAgencyList"
		objXML.send(istrParam)

        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			response.write replace(objXML.responseText, "utf-8","euc-kr")
			response.end

		Set xmlDOM = nothing
	Set objXML = nothing
	On Error Goto 0

End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>