<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/11stmy/my11stItemcls.asp"-->
<!-- #include virtual="/outmall/11stmy/incmy11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
Dim itemid, action, omy11st, failCnt, chgSellYn, arrRows, isItemIdChk
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, ccd, vMy11stGoodno
Dim isoptionyn, isText, i, chgOptCnt, vOrgprice, mayOptSoldOut, vExchangeRate, vMultiplerate, vMaySellPrice
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
ccd				= request("ccd")
failCnt			= 0
Select Case action
	Case "my11stCommonCode"			isItemIdChk = "N"
	Case Else						isItemIdChk = "Y"
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
		itemid=CLng(getNumeric(itemid))
	End If
End If
'######################################################## 11번가 API ########################################################
If action = "REG" Then									'상품 등록
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stNotRegOneItem
		If (omy11st.FResultCount < 1) Then
			iErrStr = "ERR||"&itemid&"||등록가능한 상품이 아닙니다."
		Else
			chgOptCnt = omy11st.getchangeOptionNameCnt(itemid)
			If (omy11st.FOneItem.FOptioncnt > 0) AND (chgOptCnt = 0) Then
				iErrStr = "ERR||"&itemid&"||옵션 번역 및 옵션 사용여부 확인하세요."
			Else
				If (omy11st.FOneItem.FOptioncnt > 0) AND omy11st.FOneItem.IsMayLimitSoldout = "Y" Then
					iErrStr = "ERR||"&itemid&"||옵션 수량 부족"
				Else
					'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
					If omy11st.FOneItem.checkTenItemOptionValid Then
						strSql = ""
						strSql = strSql & " IF NOT Exists(SELECT * FROM db_etcmall.[dbo].[tbl_my11st_regItem] where itemid="&itemid&")"
						strSql = strSql & " BEGIN"& VbCRLF
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_my11st_regItem] "
						strSql = strSql & " (itemid, regdate, reguserid, my11ststatCD, transItemname, regitemname)"
						strSql = strSql & " VALUES ("&itemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(omy11st.FOneitem.FTransItemName)&"',  '"&html2db(omy11st.FOneitem.FItemName)&"')"
						strSql = strSql & " END "
						dbget.Execute strSql

						strParam = ""
						strParam = omy11st.FOneItem.getMy11stItemRegXML("")

						Call fnMy11stItemReg(itemid, strParam, omy11st.FOneItem.FOrgprice, omy11st.FOneItem.FMaySellPrice, omy11st.FOneItem.FOptRecordCnt, omy11st.FOneItem.FMultiplerate, omy11st.FOneItem.FExchangeRate, iErrStr)
					Else
						iErrStr = "ERR||"&itemid&"||옵션검사 실패"
					End If
				End If
			End If
		End If

		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET omy11st = nothing
ElseIf action = "EDIT" Then								'상품 수정
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stlEditOneItem
		If omy11st.FResultCount > 0 Then
			If omy11st.FOneItem.FOptioncnt > 0 Then
				mayOptSoldOut = omy11st.FOneItem.IsMayLimitSoldout
			End If

			If (omy11st.FOneItem.FMaySoldOut = "Y") OR (omy11st.FOneItem.IsSoldOutLimit5Sell) OR (mayOptSoldOut = "Y") Then
				Call fnMy11stSoldOut(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			Else
				If (omy11st.FOneItem.FMy11stSellYn = "N" AND omy11st.FOneItem.IsSoldOut = False) Then
					iErrStr = ""
					Call fnMy11stOnSale(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
					If Left(iErrStr, 2) <> "OK" Then
						failCnt = failCnt + 1
						SumErrStr = SumErrStr & iErrStr
					Else
						SumOKStr = SumOKStr & iErrStr
					End If
				End If

				'상품 수정
				strParam = ""
				strParam = omy11st.FOneItem.getMy11stItemRegXML(omy11st.FOneItem.FMy11stGoodNo)
				Call fnMy11stItemEdit(itemid, omy11st.FOneItem.FMy11stGoodNo, strParam, omy11st.FOneItem.FOrgprice, omy11st.FOneItem.FExchangeRate, omy11st.FOneItem.FMultiplerate, omy11st.FOneItem.FMaySellPrice, omy11st.FOneItem.FOptRecordCnt, omy11st.FOneItem.FNotdb2HTMLitemname, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If

				'옵션 조회
				Call fnMy11stOptView(itemid, omy11st.FOneItem.FMy11stGoodNo, iErrStr)
				If Left(iErrStr, 2) <> "OK" Then
					failCnt = failCnt + 1
					SumErrStr = SumErrStr & iErrStr
				Else
					SumOKStr = SumOKStr & iErrStr
				End If
			End If

			If failCnt > 0 Then
				SumErrStr = replace(SumErrStr, "'", "")
				SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
				SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
				CALL Fn_AcctFailTouch("11stmy", itemid, SumErrStr)
				Call SugiQueLogInsert("11stmy", action, itemid, "ERR", "ERR||"&itemid&"||"&SumErrStr, session("ssBctID"))
				iErrStr = "ERR||"&itemid&"||"&SumErrStr
			Else
				SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
				Call SugiQueLogInsert("11stmy", action, itemid, "OK", "OK||"&itemid&"||"&SumOKStr, session("ssBctID"))
				iErrStr = "OK||"&itemid&"||"&SumOKStr
			End If
		End If
	SET omy11st = nothing
ElseIf action = "PRICE" Then							'판매 가격 수정
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
	Else
		Call getMy11stRatePrice(itemid, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice)
		Call fnMy11stPrice(itemid, vMy11stGoodno, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "SOLDOUT" Then							'판매 상태 변경 N
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
	Else
		Call fnMy11stSoldOut(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "ONSALE" Then							'판매 상태 변경 Y
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
	Else
		Call fnMy11stOnSale(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "EDITOPT" Then							'옵션 수정
	SET omy11st = new CMy11st
		omy11st.FRectItemID	= itemid
		omy11st.getmy11stlEditOneItem
		If omy11st.FResultCount > 0 Then
			strParam = ""
			strParam = omy11st.FOneItem.getMy11stOptEditXML()
			Call fnMy11stOptEdit(itemid, omy11st.FOneItem.FMy11stGoodNo, strParam, omy11st.FOneItem.FOptRecordCnt, iErrStr)
		End If
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
	SET omy11st = nothing
ElseIf action = "VIEW" Then								'상품 조회
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
	Else
		Call fnMy11stView(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "VIEWOPT" Then							'옵션 조회
	vMy11stGoodno = getMy11stGoodNo(itemid)
	If vMy11stGoodno = "" Then
		iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
	Else
		Call fnMy11stOptView(itemid, vMy11stGoodno, iErrStr)
	End If
	If LEFT(iErrStr, 2) <> "OK" Then
		CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
	End If
	Call SugiQueLogInsert("11stmy", action, itemid, Split(iErrStr,"||")(0), iErrStr, session("ssBctID"))
ElseIf action = "my11stCommonCode" Then					'공통코드 검색
	If ccd = "CATEGORYLIST" Then
		strParam = ""
		strParam = getCommCode(ccd)
	End If
End If

If iErrStr <> "" Then
	response.write  "<script>" & vbCrLf &_
					"	var str, t; " & vbCrLf &_
					"	t = parent.document.getElementById('actStr') " & vbCrLf &_
					"	str = t.innerHTML; " & vbCrLf &_
					"	str += '"&iErrStr&"<br>' " & vbCrLf &_
					"	t.innerHTML = str; " & vbCrLf &_
					"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
					"</script>"
End If
'###################################################### 11번가 API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->