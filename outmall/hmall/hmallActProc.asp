<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 60 * 15
%>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/hmall/hmallItemcls.asp"-->
<!-- #include virtual="/outmall/hmall/inchmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oHmall, failCnt, chgSellYn, arrRows, isItemIdChk, maeipdiv, mustPrice, getMustprice
Dim iErrStr, strSql, SumErrStr, SumOKStr, i, tHmallGoodno, isChkStat, strparam, mrgnRate, chgImageNm, endItemErrMsgReplace
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
failCnt			= 0

If action <> "SECTView" and action <> "infoDivView" Then
	If isItemIdChk = "Y" Then
		If itemid="" or itemid="0" Then
			response.write "<script>alert('상품번호가 없습니다.')</script>"
			response.end
		ElseIf Not(isNumeric(itemid)) Then
			response.write "<script>alert('잘못된 상품번호입니다.')</script>"
			response.end
		Else
			'정수형태로 변환
			itemid = CLng(getNumeric(itemid))
		End If
	End If
End If

'######################################################## HMall API ########################################################
If action = "REG" Then									'상품등록
'	SET oHmall = new CHMall
'		oHmall.FRectItemID	= itemid
'		oHmall.getHmallNotRegOneItem
'	    If (oHmall.FResultCount < 1) Then
'			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
'		Else
'			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
'			dbget.execute strSql
'
'			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
'			If oHmall.FOneItem.checkTenItemOptionValid Then
'				Call fnHmallItemReg(itemid, iErrStr)
'			Else
'				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
'			End If
'		End If
'		If LEFT(iErrStr, 2) <> "OK" Then
'			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
'		End If
'		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
'	SET oHmall = nothing
'###################################################	여기까지 구 버전 ##########################################
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		'oHmall.getHmallNotRegOneItem
		oHmall.getHmallNotRegOnlyOneItem
		If oHmall.FResultCount > 0 Then
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'If oHmall.FOneItem.fnCheckMakerid Then
			'	iErrStr = "ERR||"&itemid&"||[상품등록add] 간이과세 등록불가"
			If oHmall.FOneItem.checkTenItemOptionValid Then
'				Call fnHmallOnlyItemReg(itemid, iErrStr)
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemRegParameter()

				getMustprice = ""
				getMustprice = oHmall.FOneItem.MustPrice()
				Call fnHmallItemOnlyReg(itemid, strParam, iErrStr, getMustprice, oHmall.FOneItem.gethmallSellYn, oHmall.FOneItem.FLimityn, oHmall.FOneItem.FLimitNo, oHmall.FOneItem.FLimitSold, html2db(oHmall.FOneItem.FItemName), oHmall.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록add] 옵션검사 실패"
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				tHmallGoodno = getHmallGoodno(itemid)
				If tHmallGoodno <> "" Then
					chgImageNm = oHmall.FOneItem.getBasicImage
					Call fnHmallImage(itemid, chgImageNm, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If

			If failCnt = 0 Then
				tHmallGoodno = getHmallGoodno2(itemid)
				If tHmallGoodno <> "" Then
					Call fnHmallImageConfirm(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
			End If
		Else
			failCnt = 1
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql
			SumErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oHmall = nothing
ElseIf action = "REGAddItem" Then						'상품만 등록
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallNotRegOneItem
	    If (oHmall.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oHmall.FOneItem.checkTenItemOptionValid Then
				Call fnHmallOnlyItemReg(itemid, iErrStr)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "IMAGE" Then							'이미지 등록 & 확인
	tHmallGoodno = getHmallGoodno(itemid)
	If tHmallGoodno = "" Then
		failCnt = 1
		SumErrStr = "ERR||"&itemid&"||상품부터 등록 하셔야 됩니다."
	Else
		chgImageNm = getTenBasicImage(itemid)
		Call fnHmallImage(itemid, chgImageNm, iErrStr)
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If

	If failCnt = 0 Then
		tHmallGoodno = getHmallGoodno2(itemid)
		If tHmallGoodno = "" Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||상품 및 이미지부터 등록 하셔야 됩니다."
		Else
			Call fnHmallImageConfirm(itemid, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If
	End If

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
		Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "REGImage" Then							'이미지 등록
	tHmallGoodno = getHmallGoodno(itemid)
	If tHmallGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||상품부터 등록 하셔야 됩니다."
	Else
		chgImageNm = getTenBasicImage(itemid)
		Call fnHmallImage(itemid, chgImageNm, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "REGImageConfirm" Then					'이미지 확인
	tHmallGoodno = getHmallGoodno2(itemid)
	If tHmallGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||상품 및 이미지부터 등록 하셔야 됩니다."
	Else
		Call fnHmallImageConfirm(itemid, iErrStr)
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	End If
ElseIf action = "EditSellYn" Then						'상품 상태 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||상태수정 할 상품이 등록되어 있지 않습니다."
		Else
			Call fnHmallSellYN(itemid, chgSellYn, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "PRICE" Then							'상품 가격 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			mustPrice = ""
			mustPrice = oHmall.FOneItem.MustPrice()

			mrgnRate = ""
			mrgnRate = oHmall.FOneItem.FMrgnRate
			strParam = oHmall.FOneItem.getHmallPriceParameter()
			Call fnHmallPrice(itemid, mustPrice, mrgnRate, strParam, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||가격수정 할 상품이 등록되어 있지 않습니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
' ElseIf action = "CHKSTAT" Then							'상품 상세 조회	+ 재고 조회			.NET 버전 / 옵션이 많으면 너무
' 	SET oHmall = new CHMall
' 		oHmall.FRectItemID	= itemid
' 		oHmall.getHmallEditOneItem
' 		If oHmall.FResultCount > 0 Then
' 			Call fnHmallStatChk(itemid, iErrStr)
' 			If Left(iErrStr, 2) <> "OK" Then
' 				failCnt = failCnt + 1
' 				SumErrStr = SumErrStr & iErrStr
' 			Else
' 				SumOKStr = SumOKStr & iErrStr
' 			End If

' 			If INSTR(iErrStr, "승인완료") > 0 AND failCnt = 0 Then
' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If

' 				strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
' 				Call fnHmallOptionEdit(itemid, strparam, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If

' 				Call fnHmallOptionStatCheck(itemid, iErrStr)
' 				If Left(iErrStr, 2) <> "OK" Then
' 					failCnt = failCnt + 1
' 					SumErrStr = SumErrStr & iErrStr
' 				Else
' 					SumOKStr = SumOKStr & iErrStr
' 				End If
' 			End If
' 		Else
' 			failCnt = 1
' 			SumErrStr = "ERR||"&itemid&"||상세조회 할 상품이 등록되어 있지 않습니다."
' 		End If

' 		If failCnt > 0 Then
' 			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
' 			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
' 			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
' 			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

' 			iErrStr = "ERR||"&itemid&"||"&SumErrStr
' 		Else
' 			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
' 			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
' 			iErrStr = "OK||"&itemid&"||"&SumOKStr
' 		End If
' 	SET oHmall = nothing
ElseIf action = "OPTSTAT" Then							'상품 재고 조회
	'Call fnHmallOptionStatChk(itemid, iErrStr)
	Call fnHmallOptionStatCheck(itemid, iErrStr)
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "OPTEDIT" Then							'상품 옵션 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
			Call fnHmallOptionEdit(itemid, strparam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If failCnt = 0 Then
				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			failCnt = 1
			SumErrStr = "ERR||"&itemid&"||옵션수정 할 상품이 등록되어 있지 않습니다."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "EDITItem" Then							'상품 정보만 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			Call fnHmallOnlyItemEdit(itemid, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||상품정보 수정 할 상품이 등록되어 있지 않습니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
ElseIf action = "REGOnly" Then						'상품만 등록
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallNotRegOnlyOneItem
	    If (oHmall.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Hmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"'"
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oHmall.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemRegParameter()

				getMustprice = ""
				getMustprice = oHmall.FOneItem.MustPrice()
				Call fnHmallItemOnlyReg(itemid, strParam, iErrStr, getMustprice, oHmall.FOneItem.gethmallSellYn, oHmall.FOneItem.FLimityn, oHmall.FOneItem.FLimitNo, oHmall.FOneItem.FLimitSold, html2db(oHmall.FOneItem.FItemName), oHmall.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[상품등록] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=REGOnly&itemid=3887719
ElseIf action = "EDITonly" Then							'상품 정보만 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strParam = oHmall.FOneItem.gethmallItemEditParameter()
			Call fnHmallItemOnlyEdit(itemid, strParam, iErrStr)
		Else
			iErrStr = "ERR||"&itemid&"||상품정보 수정 할 상품이 등록되어 있지 않습니다."
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("hmall1010", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("hmall1010", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oHmall = nothing
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=EDITonly&itemid=2952364
ElseIf action = "CHKSTAT" Then							'상품 상세 조회	+ 재고 조회			.ASP버전(승인처리만) / 2022-01-05 추가
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount > 0 Then
			strParam = ""
			strParam = oHmall.FOneItem.getHmallItemConfirmParameter()
			Call fnHmallStatChk2(itemid, strParam, iErrStr)
			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If

			If INSTR(iErrStr, "승인완료") > 0 AND failCnt = 0 Then
				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
				Call fnHmallOptionEdit(itemid, strparam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				Call fnHmallOptionStatCheck(itemid, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If
		Else
			failCnt = 1
			SumErrStr = "ERR||"&itemid&"||상세조회 할 상품이 등록되어 있지 않습니다."
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
			Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oHmall = nothing
ElseIf action = "EDIT" Then								'상품 수정
	SET oHmall = new CHMall
		oHmall.FRectItemID	= itemid
		oHmall.getHmallEditOneItem
		If oHmall.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||수정 할 상품이 등록되어 있지 않습니다."
		Else
			'If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") OR (oHmall.FOneItem.IsAllOptionChange = "Y") OR (oHmall.FOneItem.fnCheckMakerid) Then
            If (oHmall.FOneItem.FmaySoldOut = "Y") OR (oHmall.FOneItem.IsMayLimitSoldout = "Y") Then
				Call fnHmallSellYN(itemid, "N", iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
			'############## Hmall 상품 수정 #################
'2022-05-09 김진영 하단 수정
'				Call fnHmallOnlyItemEdit(itemid, iErrStr)
'				If Left(iErrStr, 2) <> "OK" Then
'					failCnt = failCnt + 1
'					SumErrStr = SumErrStr & iErrStr
'				Else
'					SumOKStr = SumOKStr & iErrStr
'				End If
				strParam = ""
				strParam = oHmall.FOneItem.gethmallItemEditParameter()
				Call fnHmallItemOnlyEdit(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			'############## Hmall 이미지 수정 #################
				If oHmall.FOneItem.isImageChanged Then
					chgImageNm = oHmall.FOneItem.getBasicImage
					Call fnHmallImage(itemid, chgImageNm, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If

					Call fnHmallImageConfirm(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall 가격 수정 #################
				If failCnt = 0 Then
					mustPrice = ""
					mustPrice = oHmall.FOneItem.MustPrice()

					mrgnRate = ""
					mrgnRate = oHmall.FOneItem.FMrgnRate
					strParam = ""
					strParam = oHmall.FOneItem.getHmallPriceParameter()
					Call fnHmallPrice(itemid, mustPrice, mrgnRate, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall 옵션 수정 #################
				If failCnt = 0 Then
					strparam = oHmall.FOneItem.fngetOptionEditParam(itemid)
					Call fnHmallOptionEdit(itemid, strparam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall 재고 조회 #################
				If failCnt = 0 Then
					Call fnHmallOptionStatCheck(itemid, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

			'############## Hmall 판매 상태 수정 #################
				If failCnt = 0 Then
					Call fnHmallSellYN(itemid, "Y", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				endItemErrMsgReplace = replace(SumErrStr, "OK||"&itemid&"||", "")
				endItemErrMsgReplace = replace(SumErrStr, "ERR||"&itemid&"||", "")

				If (oHmall.FOneItem.IsAllOptionChange = "Y") OR (Instr(endItemErrMsgReplace, "판매가능한 속성 정보가 없습니다") > 0) OR (Instr(endItemErrMsgReplace, "판매가능한속성정보가없습니다") > 0) Then
					Call fnHmallSellYN(itemid, "N", iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					strSql = "	DECLARE @Temp CHAR(1) " & _
								"	If NOT EXISTS(SELECT * FROM [db_temp].dbo.tbl_jaehyumall_not_in_itemid Where mallgubun = 'hmall1010' AND itemid = '"& itemid &"') " & _
								"		BEGIN " & _
								"			INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_itemid(itemid,mallgubun,bigo) VALUES('"& itemid &"','hmall1010', '옵션전체변경품절(system)') " & _
								"		END	"
					dbget.execute strSql
				End If
			End If
		End If
	SET oHmall = nothing

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("hmall1010", itemid, SumErrStr)
		Call SugiQueLogInsert("hmall1010", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))

		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_hmall_regItem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("hmall1010", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
ElseIf action = "SECTView" Then
	Call fnHmallSectView()
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=SECTView
ElseIf action = "infoDivView" Then
	Call fnHmallInfoDivView()
	'http://localhost:11117/outmall/hmall/hmallActProc.asp?act=infoDivView
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&iErrStr&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->