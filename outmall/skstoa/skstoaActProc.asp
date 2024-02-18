<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/outmall/skstoa/skstoaItemcls.asp"-->
<!-- #include virtual="/outmall/skstoa/incskstoaFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, oSkstoa, failCnt, chgSellYn, arrRows, getMustprice, interfaceId
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, salegb, iskGoodNo
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
interfaceId		= request("interfaceId")
failCnt			= 0
iskGoodNo		= ""

Select Case action
	Case "commonCode", "cateList", "certList", "shipCost", "offerList"
		isItemIdChk = "N"
	Case Else
		isItemIdChk = "Y"
End Select

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
'######################################################## skstoa API ########################################################
'상품 등록 PROCESS
'IF_API_10_005 >> IF_API_10_001 / IF_API_10_002 / IF_API_10_003 / IF_API_10_006 /  IF_API_10_028 / >> IF_API_10_011
'기초정보등록		기술서등록			단품등록		이미지URL		정보고시	  KC 인증 추가(Optional)	승인요청
' 	reglevel1		reglevel2		reglevel3		reglevel4		reglevel5		reglevel6			reglevel7
If action = "REG" Then
	'##################################### 기초정보등록 시작 #####################################
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getskstoaNotRegOneItem
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oSkstoa.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaItemRegParameter()

				getMustprice = ""
				getMustprice = oSkstoa.FOneItem.MustPrice()
				Call fnskstoaItemReg(itemid, strParam, iErrStr, getMustprice, oSkstoa.FOneItem.getskstoaSellYn, oSkstoa.FOneItem.FLimityn, oSkstoa.FOneItem.FLimitNo, oSkstoa.FOneItem.FLimitSold, html2db(oSkstoa.FOneItem.FItemName), oSkstoa.FOneItem.FbasicimageNm)
				rw "기초정보등록"
				response.flush
				response.clear
			Else
				iErrStr = "ERR||"&itemid&"||[임시등록] 옵션검사 실패"
			End If
		End If
	SET oSkstoa = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### 기초정보등록 끝 #######################################

	'##################################### 기술서등록 시작 #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf oSkstoa.FOneItem.FReglevel <> 1 Then
				iErrStr = "ERR||"&itemid&"||상품부터 등록하세요."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaContentParameter()
				Call fnSkstoaContentReg(itemid, strParam, iErrStr)
				rw "기술서등록"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 기술서등록 끝 #######################################

	'##################################### 단품등록 시작 #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf oSkstoa.FOneItem.FReglevel <> 2 Then
				failCnt = failCnt + 1
				iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				arrRows = getOptionList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaOptParameter(arrRows(0, i), arrRows(1, i))
						Call fnSkstoaOptReg(itemid, strParam, iErrStr)
						If iErrStr <> "" Then
							SumErrStr = SumErrStr & arrRows(2, i) & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("REGOpt", itemid, SumErrStr)
					rw "단품등록"
					response.flush
					response.clear
				End If
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 단품등록 끝 #########################################

	'##################################### 이미지URL 시작 ######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf oSkstoa.FOneItem.FReglevel <> 3 Then
				iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaImageParameter()
				Call fnSkstoaImageReg(itemid, strParam, iErrStr)
				rw "이미지URL 등록"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 이미지URL 끝 ########################################

	'##################################### 정보고시 시작 #######################################
	If failCnt = 0 Then
		iErrStr = ""
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf oSkstoa.FOneItem.FReglevel <> 4 Then
				failCnt = failCnt + 1
				SumErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				arrRows = getInfoCodeMapList(itemid)
				If isArray(arrRows) Then
					For i = 0 To UBound(arrRows,2)
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
						Call fnSkstoaGosiReg(itemid, strParam, iErrStr)
						If iErrStr <> "" Then
							SumErrStr = SumErrStr & arrRows(0, i) & ","
						End If
					Next
					iErrStr = ArrErrStrInfo("REGGosi", itemid, SumErrStr)
					rw "정보고시 등록"
					response.flush
					response.clear
				End If
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 정보고시 끝 ##########################################

	'##################################### 안전인증등록 시작 ####################################
	If (failCnt = 0) AND (getMayCertYn(itemid)) = "Y" Then
		iErrStr = ""
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaCertParameter() 
				Call fnSkstoaCert(itemid, strParam, iErrStr)
				rw "안전인증 등록"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 안전인증등록 시작 #####################################

	'##################################### 승인요청 시작 #########################################
	If (failCnt = 0) Then
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
			ElseIf oSkstoa.FOneItem.FReglevel <> 5 AND oSkstoa.FOneItem.FReglevel <> 6 Then
				iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
				Call fnSkstoaConfirm(itemid, strParam, iErrStr)
				rw "승인요청"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 승인요청 끝 ###########################################
	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
		Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REG&itemid=4412416
ElseIf action = "CONFIRM" Then
	'##################################### 임시상품 조회(상세) 시작 ###############################
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("Y")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnskstoaRegChkstat(itemid, strParam, iErrStr, iskGoodNo)
			rw "승인조회"
			response.flush
			response.clear
		End If
	SET oSkstoa = nothing
	If Left(iErrStr, 2) <> "OK" Then
		failCnt = failCnt + 1
		SumErrStr = SumErrStr & iErrStr
	Else
		SumOKStr = SumOKStr & iErrStr
	End If
	'##################################### 임시상품 조회(상세) 끝 #################################

	'##################################### 판매상품 조회(상세) 시작 ###############################
	If failCnt = 0 AND iskGoodNo <> "" Then
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaEditOneItem

			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
				Call fnSkstoaItemView(itemid, strParam, iErrStr)
				rw "판매상품 조회"
				response.flush
				response.clear
			End If
		SET oSkstoa = nothing
		If Left(iErrStr, 2) <> "OK" Then
			failCnt = failCnt + 1
			SumErrStr = SumErrStr & iErrStr
		Else
			SumOKStr = SumOKStr & iErrStr
		End If
	End If
	'##################################### 판매상품 조회(상세) 끝 #################################

	If failCnt > 0 Then
		SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
		SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
		CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
		Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
		iErrStr = "ERR||"&itemid&"||"&SumErrStr
	Else
		strSql = ""
		strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regitem SET " & VBCRLF
		strSql = strSql & " accFailcnt = 0  " & VBCRLF
		strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
		dbget.Execute strSql

		SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
		Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
		iErrStr = "OK||"&itemid&"||"&SumOKStr
	End If
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=CONFIRM&itemid=4412416
ElseIf action = "PRICE" Then							'IF_API_10_031 / 판매상품 상품가격 수정
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||가격 수정 가능한 상품이 아닙니다."
		Else
			getMustprice = ""
			getMustprice = oSkstoa.FOneItem.MustPrice()
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditPriceParameter()
			Call fnSkstoaEditPrice(itemid, strParam, getMustprice, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=PRICE&itemid=3853757
ElseIf action = "EDIT" Then
'상품 수정 PROCESS
' 조회 > 판매여부 N이면 종료
' 조회 > 가격수정 > 기본정보수정 > 기술서수정 > 이미지변경여부 검색후 수정 > 판매상태수정 > 재고수정 > 옵션추가  > 조회
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

		If oSkstoa.FResultCount = 0 Then
			iErrStr = "ERR||"&itemid&"||수정 할 상품이 등록되어 있지 않습니다."
		Else
			'checkTenItemOptionValid2 => 옵션 사용여부와 옵션추가금액 체크
			If (oSkstoa.FOneItem.FmaySoldOut = "Y") OR (oSkstoa.FOneItem.IsMayLimitSoldout = "Y") OR (oSkstoa.FOneItem.checkTenItemOptionValid2 <> "True") OR (oSkstoa.FOneItem.FLimityn = "Y" AND (oSkstoa.FOneItem.getiszeroWonSoldOut(itemid) = "Y")) Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaSellynParameter("N")
				Call fnSkstoaSellyn(itemid, "N", strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
	'##################################### 판매상품 조회 시작 #######################################
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
				Call fnSkstoaItemView(itemid, strParam, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
				rw "판매상품 조회"
				response.flush
				response.clear
	'##################################### 판매상품 조회 끝 #########################################

	'##################################### 협력사 가격등록 시작 #####################################
				If failCnt = 0 Then
					iErrStr = ""
					getMustprice = ""
					getMustprice = oSkstoa.FOneItem.MustPrice()
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaEditPriceParameter()
					Call fnskstoaEditPrice(itemid, strParam, getMustprice, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "판매상품 가격수정"
					response.flush
					response.clear
				End If
	'##################################### 협력사 가격등록 끝 #######################################

	'##################################### 협력사 기초정보 수정 시작 ################################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaItemEditParameter()
					Call fnSkstoaItemEdit(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "판매상품 기초정보 수정"
					response.flush
					response.clear
				End If
	'##################################### 협력사 기초정보 수정 끝 ##################################

	'##################################### 협력사 기술서 수정 시작 ##################################
				If failCnt = 0 Then
					iErrStr = ""
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaEditContentParameter()
					Call fnSkstoaEditContentReg(itemid, strParam, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
					rw "판매상품 기술서 수정"
					response.flush
					response.clear
				End If
	'##################################### 협력사 기술서 수정 끝 ####################################

	'##################################### 이미지 수정 시작 #########################################
				If failCnt = 0 Then
					If oSkstoa.FOneItem.isImageChanged = True Then
						iErrStr = ""
						strParam = ""
						strParam = oSkstoa.FOneItem.getskstoaEditImageParameter()
						Call fnSkstoaEditImage(itemid, strParam, iErrStr)

						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
						rw "이미지 수정"
						response.flush
						response.clear
					End If
				End If
	'##################################### 이미지 수정 끝 ###########################################

	'##################################### 옵션 판매상태 수정 시작 ###################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMapList(itemid, "S")
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'재고가 1개 이하거나 옵션명이 다르거나 옵션사용여부N 이거나 옵션판매여부 N이거나
								salegb = "11"
							Else
								salegb = "00"
							End If

							strParam = ""
							strParam = oSkstoa.FOneItem.geskstoaOptionStatParam(arrRows(2, i), salegb)
							Call fnSkstoaOptSellyn(itemid, strParam, iErrStr)
							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITSTAT", itemid, SumErrStr)
						rw "옵션 판매상태 수정"
						response.flush
						response.clear
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
	'##################################### 옵션 판매상태 수정 끝 #####################################

	'##################################### 옵션 재고 수정 시작 #######################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMapList(itemid, "C")
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							strParam = ""
							strParam = oSkstoa.FOneItem.geskstoaOptionQtyParam(arrRows(2, i), arrRows(7, i))
							Call fnSkstoaQtyEdit(itemid, strParam, iErrStr)
							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITQTY", itemid, SumErrStr)
						rw "옵션 재고 수정"
						response.flush
						response.clear
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
					End If
				End If
	'##################################### 옵션 재고 수정 끝 #########################################

	'##################################### 판매상품 옵션추가 시작 #####################################
				If failCnt = 0 Then
					iErrStr = ""
					arrRows = ""
					arrRows = getOptiopnMayAddList(itemid)
					If isArray(arrRows) Then
						For i = 0 To UBound(arrRows,2)
							strParam = ""
							strParam = oSkstoa.FOneItem.getskstoaOptionAddParam(arrRows(0, i), arrRows(1, i))
							Call fnSkstoaOptAdd(itemid, strParam, iErrStr)

							If iErrStr <> "" Then
								SumErrStr = SumErrStr & arrRows(2, i) & ","
							End If
						Next
						iErrStr = ArrErrStrInfo("EDITADDOPT", itemid, SumErrStr)
					Else
						iErrStr = "OK||"&itemid&"||성공[옵션추가x]]"
					End If
					rw "옵션 추가"
					response.flush
					response.clear

					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If
	'##################################### 판매상품 옵션추가 끝 #######################################

	'##################################### 판매상품 조회 시작 #######################################
'				If failCnt = 0 Then
					If InStr(SumErrStr, "관리자 종료처리") < 0 Then
						strParam = ""
						strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
						Call fnSkstoaItemView(itemid, strParam, iErrStr)
						If Left(iErrStr, 2) <> "OK" Then
							failCnt = failCnt + 1
							SumErrStr = SumErrStr & iErrStr
						Else
							SumOKStr = SumOKStr & iErrStr
						End If
						rw "판매상품 조회"
						response.flush
						response.clear
					End If
'				End If
	'##################################### 판매상품 조회 끝 #########################################
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			If InStr(SumErrStr, "관리자 종료처리") < 0 Then
				CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			strSql = ""
			strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regItem SET " & VBCRLF
			strSql = strSql & " accFailcnt = 0  " & VBCRLF
			strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
			dbget.Execute strSql

			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDIT&itemid=4466673
ElseIf action = "REGAddItem" Then					'IF_API_10_005 / 임시상품 기초정보 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getskstoaNotRegOneItem
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strSql = "EXEC [db_etcmall].[dbo].[usp_API_Outmall_RegItem_Add] '"&itemid&"', '"&session("SSBctID")&"', '"&CMALLNAME&"' "
			dbget.execute strSql

			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oSkstoa.FOneItem.checkTenItemOptionValid Then
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaItemRegParameter()

				getMustprice = ""
				getMustprice = oSkstoa.FOneItem.MustPrice()
				Call fnskstoaItemReg(itemid, strParam, iErrStr, getMustprice, oSkstoa.FOneItem.getskstoaSellYn, oSkstoa.FOneItem.FLimityn, oSkstoa.FOneItem.FLimitNo, oSkstoa.FOneItem.FLimitSold, html2db(oSkstoa.FOneItem.FItemName), oSkstoa.FOneItem.FbasicimageNm)
			Else
				iErrStr = "ERR||"&itemid&"||[임시등록] 옵션검사 실패"
			End If
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGAddItem&itemid=3853757
ElseIf action = "REGContent" Then					'IF_API_10_001 / 임시상품 기술서 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oSkstoa.FOneItem.FReglevel <> 1 Then
			iErrStr = "ERR||"&itemid&"||상품부터 등록하세요."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaContentParameter()
			Call fnSkstoaContentReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGContent&itemid=3853757
ElseIf action = "REGOpt" Then						'IF_API_10_002 / 임시상품 단품정보 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oSkstoa.FOneItem.FReglevel <> 2 Then
			failCnt = failCnt + 1
			iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			arrRows = getOptionList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaOptParameter(arrRows(0, i), arrRows(1, i))
					Call fnSkstoaOptReg(itemid, strParam, iErrStr)

					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGOpt&itemid=3853757
ElseIf action = "REGImg" Then						'IF_API_10_003 / 임시상품 이미지 등록(URL)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oSkstoa.FOneItem.FReglevel <> 3 Then
			iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaImageParameter()
			Call fnSkstoaImageReg(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGImg&itemid=3853757
ElseIf action = "REGGosi" Then						'IF_API_10_006 / 임시상품 정보제공고시 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oSkstoa.FOneItem.FReglevel <> 4 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaGosiRegParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnSkstoaGosiReg(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(0, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGGosi&itemid=3853757
ElseIf action = "REGCert" Then						'IF_API_10_028 / KC 인증 추가
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaCertParameter() 
			Call fnSkstoaCert(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGCert&itemid=3853757
ElseIf action = "REGConfirm" Then					'IF_API_10_011 / 임시상품 승인요청
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		ElseIf oSkstoa.FOneItem.FReglevel <> 5 AND oSkstoa.FOneItem.FReglevel <> 6 Then
			iErrStr = "ERR||"&itemid&"||등록 레벨 확인 (현재 : "& oSkstoa.FOneItem.FReglevel &") "
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnSkstoaConfirm(itemid, strParam, iErrStr)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGConfirm&itemid=3853757
ElseIf action = "REGCHKSTAT" Then					'IF_API_10_013 / 임시상품 조회(상세) -> 여기서 판매상품코드를 얻는다.
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaTmpRegedOneItem("Y")
	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
			Call fnskstoaRegChkstat(itemid, strParam, iErrStr, iskGoodNo)
			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
			Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
		End If
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=REGCHKSTAT&itemid=3853757
ElseIf action = "EditSellYn" Then					'IF_API_10_023 / 상품 판매중단 처리
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getSkstoaSellynParameter(chgSellYn)
			Call fnSkstoaSellyn(itemid, chgSellYn, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EditSellYn&itemid=3853757&chgsellyn=N
ElseIf action = "CHKSTAT" Then						'IF_API_10_015 / 판매상품 조회(상세)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
			Call fnSkstoaItemView(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=CHKSTAT&itemid=3853757
ElseIf action = "EDITINFO" Then						'IF_API_10_017 / 판매상품 기초정보 수정
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||기초정보 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaItemEditParameter()
			Call fnSkstoaItemEdit(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITINFO&itemid=3853757
ElseIf action = "EDITContent" Then					'IF_API_10_019 / 판매상품 기술서 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||기술서 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditContentParameter()
			Call fnSkstoaEditContentReg(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITContent&itemid=3853757
ElseIf action = "EDITImage" Then					'IF_API_10_021 / 판매상품 이미지 등록(URL)
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||이미지 수정 가능한 상품이 아닙니다."
		Else
			strParam = ""
			strParam = oSkstoa.FOneItem.getskstoaEditImageParameter()
			Call fnSkstoaEditImage(itemid, strParam, iErrStr)
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITImage&itemid=3853757
ElseIf action = "EDITQTY" Then						'IF_API_10_018 / 판매상품 재고변경
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||재고 수정 가능한 상품이 아닙니다."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||조회부터 진행하세요."
		Else
			arrRows = getOptiopnMapList(itemid, "C")
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.geskstoaOptionQtyParam(arrRows(2, i), arrRows(7, i))
					Call fnSkstoaQtyEdit(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITQTY&itemid=3853757
ElseIf action = "EDITSTAT" Then						'IF_API_10_023 / 상품 판매중단 처리
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||조회부터 진행하세요."
		Else
			arrRows = getOptiopnMapList(itemid, "S")
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					If (arrRows(7, i) < 1) OR (arrRows(11, i)= "1") OR (arrRows(9, i)= "N") OR (arrRows(10, i) = "N") Then	'재고가 1개 이하거나 옵션명이 다르거나 옵션사용여부N 이거나 옵션판매여부 N이거나
						salegb = "11"
					Else
						salegb = "00"
					End If

					strParam = ""
					strParam = oSkstoa.FOneItem.geskstoaOptionStatParam(arrRows(2, i), salegb)
					Call fnSkstoaOptSellyn(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITSTAT&itemid=4466673
ElseIf action = "EDITADDOPT" Then					'IF_API_10_020 / 판매상품 단품정보 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
		ElseIf getSkstoaOptCnt(itemid) = 0 Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||조회부터 진행하세요."
		Else
			arrRows = getOptiopnMayAddList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaOptionAddParam(arrRows(0, i), arrRows(1, i))
					Call fnSkstoaOptAdd(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(2, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)
			Else
				iErrStr = "OK||"&itemid&"||성공[옵션추가x]]"
			End If

			If Left(iErrStr, 2) <> "OK" Then
				failCnt = failCnt + 1
				SumErrStr = SumErrStr & iErrStr
			Else
				SumOKStr = SumOKStr & iErrStr
			End If
		End If

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITADDOPT&itemid=4466673
ElseIf action = "EDITGosi" Then					'IF_API_10_016 / 판매상품 정보제공고시 등록
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

	    If (oSkstoa.FResultCount < 1) Then
			failCnt = failCnt + 1
			SumErrStr = "ERR||"&itemid&"||정보제공고시 수정 가능한 상품이 아닙니다."
		Else
			arrRows = getInfoCodeMapList(itemid)
			If isArray(arrRows) Then
				For i = 0 To UBound(arrRows,2)
					strParam = ""
					strParam = oSkstoa.FOneItem.getskstoaGosiEditParameter(arrRows(0, i), arrRows(1, i), arrRows(2, i))
					Call fnSkstoaGosiEdit(itemid, strParam, iErrStr)
					If iErrStr <> "" Then
						SumErrStr = SumErrStr & arrRows(0, i) & ","
					End If
				Next
				iErrStr = ArrErrStrInfo(action, itemid, SumErrStr)

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
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			Call SugiQueLogInsert("skstoa", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
			iErrStr = "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			Call SugiQueLogInsert("skstoa", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
			iErrStr = "OK||"&itemid&"||"&SumOKStr
		End If
 	SET oSkstoa = nothing
 	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITGosi&itemid=4466673
 ElseIf action = "EDITCert" Then					'IF_API_10_030 / 판매상품 KC인증정보 수정
	SET oSkstoa = new CSkstoa
		oSkstoa.FRectItemID	= itemid
		oSkstoa.getSkstoaEditOneItem

 	    If (oSkstoa.FResultCount < 1) Then
 			iErrStr = "ERR||"&itemid&"||인증정보 수정 가능한 상품이 아닙니다."
 		Else
 			strParam = ""
 			strParam = oSkstoa.FOneItem.getskstoaEditCertParameter()
 			Call fnSkstoaEditCert(itemid, strParam, iErrStr)
 		End If

 		If LEFT(iErrStr, 2) <> "OK" Then
 			CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
 		End If
 		Call SugiQueLogInsert("skstoa", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
 	SET oSkstoa = nothing
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=EDITCert&itemid=4466673
ElseIf action = "commonCode" Then					'IF_API_00_001 ~ 
	Call fnGetCommonCodeList(interfaceId)
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=commonCode&interfaceId=IF_API_00_001
ElseIf action = "cateList" Then						'IF_API_00_028 / 상품 세분류 조회
	Call fnGetGoodsDgroupList()
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=cateList
ElseIf action = "offerList" Then					'IF_API_00_023 / 상품정보제공고시 항목 조회
	Call fnGetOfferList()
	response.end
	'http://localhost:11117/outmall/skstoa/skstoaActProc.asp?act=offerList
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str = '"&iErrStr&"<br>' + str " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### skstoa API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
