<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/sabangnet/sabangnetItemcls.asp"-->
<!-- #include virtual="/outmall/sabangnet/incSabangnetFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<%
Dim itemid, mallid, action, failCnt, oSabangnet, getMustprice, chgSellYn, vOptCnt
Dim iErrStr, strParam, mustPrice, strSql, SumErrStr, SumOKStr, chgImageNm
Dim jenkinsBatchYn, idx, lastErrStr
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
jenkinsBatchYn	= request("jenkinsBatchYn")
idx				= request("idx")
lastErrStr		= ""
If itemid="" or itemid="0" Then
	response.write "<script>alert('상품번호가 없습니다.')</script>"
	response.end
ElseIf Not(isNumeric(itemid)) Then
	response.write "<script>alert('잘못된 상품번호입니다.')</script>"
	response.end
Else
	'정수형태로 변환
	itemid=CLng(getNumeric(itemid))
End If
'######################################################## Sabangnet API ########################################################
If mallid = "sabangnet" Then
	If action = "SOLDOUT" Then				'상태변경
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetSimpleEditOneItem
		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[요약수정상태] 수정 가능한 상품이 아닙니다."
				response.write "ERR||"&itemid&"||[요약수정상태] 수정 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter("N")
				Call fnSabangnetSimpleEdit(itemid, "N", oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "sellyn")
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=SOLDOUT
	ElseIf action = "PRICE" Then			'가격수정
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetSimpleEditOneItem
		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[요약수정가격] 수정 가능한 상품이 아닙니다."
				response.write "ERR||"&itemid&"||[요약수정가격] 수정 가능한 상품이 아닙니다."
			Else
				strParam = ""
				If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
					chgSellYn = "N"
					strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
				Else
					chgSellYn = "Y"
					strParam = oSabangnet.FOneItem.getSabangnetSimpleEditItemParameter(chgSellYn)
				End If

				Call fnSabangnetSimpleEdit(itemid, chgSellYn, oSabangnet.FOneItem.MustPrice, html2db(oSabangnet.FOneItem.FItemName), strParam, iErrStr, "price")
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=PRICE
	ElseIf action = "EDIT" Then				'상품수정
		SET oSabangnet = new CSabangnet
			oSabangnet.FRectItemID	= itemid
			oSabangnet.getSabangnetEditOneItem

		    If (oSabangnet.FResultCount < 1) Then
				lastErrStr = "ERR||"&itemid&"||[전체수정] 수정 가능한 상품이 아닙니다."
				response.write "ERR||"&itemid&"||[전체수정] 수정 가능한 상품이 아닙니다."
			Else
				If (oSabangnet.FOneItem.FmaySoldOut = "Y") OR (oSabangnet.FOneItem.IsMayLimitSoldout = "Y") OR (oSabangnet.FOneItem.IsSoldOut) Then
					chgSellYn = "N"
				Else
					chgSellYn = "Y"
				End If
				strParam = ""
				strParam = oSabangnet.FOneItem.getSabangnetItemRegParameter(True, chgSellYn)
				Call fnSabangnetItemEdit(itemid, strParam, iErrStr, oSabangnet.FOneItem.MustPrice, chgImageNm, oSabangnet.FOneItem.FLimityn, oSabangnet.FOneItem.FLimitno, oSabangnet.FOneItem.FLimitsold, chgSellYn)
				lastErrStr = iErrStr
				response.write iErrStr
			End If
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("sabangnet", itemid, iErrStr)
			End If
		SET oSabangnet = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/sabangnetProc.asp?itemid=1649348&mallid=sabangnet&action=PRICE
	End If
End If
'###################################################### Sabangnet API END #######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->