<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/incJenkinsCommon.asp" -->
<!-- #include virtual="/outmall/skstoa/skstoaItemcls.asp"-->
<!-- #include virtual="/outmall/skstoa/incskstoaFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/outmall/batch/batchfunction.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, mallid, action, oSkstoa, failCnt, chgSellYn, arrRows, getMustprice, interfaceId, getShipCostCode
Dim iErrStr, strParam, strSql, SumErrStr, SumOKStr, isItemIdChk, grpVal, rSkip, rLimit, i, salegb, iskGoodNo
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
'######################################################## skstoa API ########################################################
If mallid = "skstoa" Then
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
		'##################################### 안전인증등록 끝 #####################################

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
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
	'http://localhost:11117/outmall/proc/skstoaProc.asp?mallid=skstoa&action=REG&itemid=4371598
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
				If (oSkstoa.FOneItem.FmaySoldOut = "Y") OR (oSkstoa.FOneItem.IsMayLimitSoldout = "Y") OR (oSkstoa.FOneItem.checkTenItemOptionValid2 <> "True") Then
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
'					If failCnt = 0 Then
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
							response.flush
							response.clear
						End If
'					End If
		'##################################### 판매상품 조회 끝 #########################################
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				If InStr(SumErrStr, "관리자 종료처리") < 0 Then
					CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
				End If
				lastErrStr = "ERR||"&itemid&"||"&SumErrStr
				response.write "ERR||"&itemid&"||"&SumErrStr
			Else
				strSql = ""
				strSql = strSql & " UPDATE db_etcmall.dbo.tbl_skstoa_regItem SET " & VBCRLF
				strSql = strSql & " accFailcnt = 0  " & VBCRLF
				strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
				dbget.Execute strSql

				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				lastErrStr = "OK||"&itemid&"||"&SumOKStr
				response.write "OK||"&itemid&"||"&SumOKStr
			End If
		SET oSkstoa = nothing
	ElseIf action = "CONFIRM" Then
		'##################################### 승인요청 시작 #####################################
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaTmpRegedOneItem("Y")
			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다.1"
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getskstoaConfirmParameter()
				Call fnskstoaRegChkstat(itemid, strParam, iErrStr, iskGoodNo)
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
		'##################################### 승인요청 끝 #######################################

		'##################################### 판매상품 조회(상세) 시작 ###########################
		If failCnt = 0 AND iskGoodNo <> "" Then
			SET oSkstoa = new CSkstoa
				oSkstoa.FRectItemID	= itemid
				oSkstoa.getSkstoaEditOneItem

				If (oSkstoa.FResultCount < 1) Then
					iErrStr = "ERR||"&itemid&"||조회 가능한 상품이 아닙니다.2"
				Else
					strParam = ""
					strParam = oSkstoa.FOneItem.getSkstoaItemViewParameter()
					Call fnSkstoaItemView(itemid, strParam, iErrStr)
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
		'##################################### 판매상품 조회(상세) 끝 #############################

		If failCnt > 0 Then
			SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
			SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
			CALL Fn_AcctFailTouch("skstoa", itemid, SumErrStr)
			lastErrStr = "ERR||"&itemid&"||"&SumErrStr
			response.write "ERR||"&itemid&"||"&SumErrStr
		Else
			SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
			lastErrStr = "OK||"&itemid&"||"&SumOKStr
			response.write "OK||"&itemid&"||"&SumOKStr
		End If
		'http://localhost:11117/outmall/proc/skstoaProc.asp?mallid=skstoa&action=CONFIRM&itemid=4466673
	ElseIf action = "SOLDOUT" Then
		SET oSkstoa = new CSkstoa
			oSkstoa.FRectItemID	= itemid
			oSkstoa.getSkstoaEditOneItem

			If (oSkstoa.FResultCount < 1) Then
				iErrStr = "ERR||"&itemid&"||수정 가능한 상품이 아닙니다."
			Else
				strParam = ""
				strParam = oSkstoa.FOneItem.getSkstoaSellynParameter("N")
				Call fnSkstoaSellyn(itemid, "N", strParam, iErrStr)
			End If
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
		SET oSkstoa = nothing
		'http://localhost:11117/outmall/proc/skstoaProc.asp?mallid=skstoa&action=SOLDOUT&itemid=4466673
	ElseIf action = "PRICE" Then
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
			lastErrStr = iErrStr
			response.write iErrStr

			If LEFT(iErrStr, 2) <> "OK" Then
				CALL Fn_AcctFailTouch("skstoa", itemid, iErrStr)
			End If
		SET oSkstoa = nothing
		'http://localhost:11117/outmall/proc/skstoaProc.asp?mallid=skstoa&action=PRICE&itemid=4466673
	End If
End If
'###################################################### skstoa API END ######################################################
If jenkinsBatchYn = "Y" and lastErrStr <> "" Then
	strSql = ""
	strSql = "db_etcmall.[dbo].[sp_Ten_OutMall_API_Que_ResultWrite] "&idx&","&itemid&",'"&Split(lastErrStr, "||")(0)&"','"&html2DB(Split(lastErrStr, "||")(2))&"'"
	dbget.Execute strSql
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
