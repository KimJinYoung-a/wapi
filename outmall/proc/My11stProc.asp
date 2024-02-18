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
Dim itemid, mallid, action, omy11st, failCnt, chgSellYn, arrRows, skipItem, sellgubun, getMustprice, mayOptSoldOut
Dim iErrStr, strParam, mustPrice, ret1, strSql, SumErrStr, SumOKStr, iitemname, optReset, optString
Dim vMy11stGoodno, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice
itemid			= requestCheckVar(request("itemid"),9)
mallid			= request("mallid")
action			= request("action")
failCnt			= 0
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
'######################################################## 11st API ########################################################
If mallid = "11stmy" Then
	If action = "SOLDOUT" Then													'상태 변경
		vMy11stGoodno = getMy11stGoodNo(itemid)
		If vMy11stGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
		Else
			Call fnMy11stSoldOut(itemid, vMy11stGoodno, iErrStr)
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/My11stProc.asp?itemid=282197&mallid=11stmy&action=SOLDOUT
	ElseIf action = "PRICE" Then												'가격 수정
		vMy11stGoodno = getMy11stGoodNo(itemid)
		If vMy11stGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
		Else
			Call getMy11stRatePrice(itemid, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice)
			Call fnMy11stPrice(itemid, vMy11stGoodno, vOrgprice, vExchangeRate, vMultiplerate, vMaySellPrice, iErrStr)
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/My11stProc.asp?itemid=282197&mallid=11stmy&action=PRICE
	ElseIf action = "VIEWOPT" Then												'옵션 조회
		vMy11stGoodno = getMy11stGoodNo(itemid)
		If vMy11stGoodno = "" Then
			iErrStr = "ERR||"&itemid&"||11번가 상품코드 없음"
		Else
			Call fnMy11stOptView(itemid, vMy11stGoodno, iErrStr)
		End If
		response.write iErrStr
		If LEFT(iErrStr, 2) <> "OK" Then
			CALL Fn_AcctFailTouch("11stmy", itemid, iErrStr)
		End If
		'http://testwapi.10x10.co.kr/outmall/proc/My11stProc.asp?itemid=282197&mallid=11stmy&action=VIEWOPT
	ElseIf action = "EDIT" Then													'상품 수정
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

					'OK던 ERR이던 editQuecnt에 + 1을 시킴..
					'스케줄링에서 editQuecnt ASC, i.lastupdate DESC로 중복을 막자
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_my11st_regItem SET " & VBCRLF
					strSql = strSql & " EditQueCnt = isnull(editQuecnt, 0) + 1 " & VBCRLF
					strSql = strSql & " ,my11stLastUpdate = getdate()  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
				End If

				If failCnt > 0 Then
					SumErrStr = replace(SumErrStr, "OK||"&itemid&"||", "")
					SumErrStr = replace(SumErrStr, "ERR||"&itemid&"||", "")
					CALL Fn_AcctFailTouch("11stmy", itemid, SumErrStr)
					response.write "ERR||"&itemid&"||"&SumErrStr
				Else
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_my11st_regItem SET " & VBCRLF
					strSql = strSql & " accFailcnt = 0  " & VBCRLF
					strSql = strSql & " WHERE itemid = '"&itemid&"' " & VBCRLF
					dbget.Execute strSql
					SumOKStr = replace(SumOKStr, "OK||"&itemid&"||", "")
					response.write "OK||"&itemid&"||"&SumOKStr
				End If
			End If
		SET omy11st = nothing
		'http://testwapi.10x10.co.kr/outmall/proc/My11stProc.asp?itemid=436497&mallid=11stmy&action=EDIT
	End If
End If
'###################################################### 11st API END #######################################################
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->