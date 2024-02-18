<%
'// 한글 한글 한글

'// 품절로 인한 전체취소인지
Function IsAllStockOutCancel(orderserial)
	Dim vQuery, arr
	dim reducedPriceSUM, cancelReducedPriceSUM
	IsAllStockOutCancel = True

	if orderserial="" then exit Function

    '// 출고완료 내역은 품절등록되어 있어도 취소상품이 아니다.

	vQuery = " select "
	vQuery = vQuery & " 	IsNull(sum(case when d.itemid <> 0 then d.reducedPrice*d.itemno else 0 end),0) as reducedPriceSUM "
	vQuery = vQuery & " 	, IsNull(sum(case when d.itemid <> 0 and IsNull(m.code, '') in ('05','06') and IsNull(d.currstate, '0') < '7' then d.reducedPrice*IsNull(m.itemlackno,0) else 0 end),0) as cancelReducedPriceSUM "
	'vQuery = vQuery & " 	, IsNull(sum(case when d.itemid <> 0 and (IsNull(m.code, '') in ('05') or (IsNull(m.code, '') in ('03') and d.isupchebeasong='N')) and IsNull(d.currstate, '0') < '7' then d.reducedPrice*IsNull(m.itemlackno,0) else 0 end),0) as cancelReducedPriceSUM "
	vQuery = vQuery & " from "
	vQuery = vQuery & " [db_order].[dbo].[tbl_order_detail] d "
	vQuery = vQuery & " left join db_temp.dbo.tbl_mibeasong_list m "
	vQuery = vQuery & " on "
	vQuery = vQuery & " 	d.idx = m.detailidx "
	vQuery = vQuery & " where "
	vQuery = vQuery & " 	1 = 1 "
	vQuery = vQuery & " 	and d.orderserial = '" & orderserial & "' "
	vQuery = vQuery & " 	and d.cancelyn <> 'Y' "						'// 절대 출고완료된 상품 빼면 안된다.(참조 : /cscenter/lib/csAsfunction.asp : RegWebCSDetailAllCancel)
	'response.write vQuery & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.eof then
		reducedPriceSUM = rsget("reducedPriceSUM")
		cancelReducedPriceSUM = rsget("cancelReducedPriceSUM")
	end if
	rsget.close

	if (reducedPriceSUM <> cancelReducedPriceSUM) then
		IsAllStockOutCancel = False
		exit Function
	end if
End Function

function ChkStockoutItemExist(myorderdetail)
	ChkStockoutItemExist = False
	for i=0 to myorderdetail.FResultCount-1
		if (myorderdetail.FItemList(i).Fmibeasoldoutyn = "Y") then
			ChkStockoutItemExist = True
			exit for
		end if
		'if (myorderdetail.FItemList(i).Fmibeadelayyn = "Y") then
		'	ChkStockoutItemExist = True
		'	exit for
		'end if
		if (myorderdetail.FItemList(i).FmibeaDeliveryStrikeyn = "Y") then
			ChkStockoutItemExist = True
			exit for
		end if
	next
end function

function ChkStockoutItemExist_Proc(orderserial)
	dim vQuery

	ChkStockoutItemExist_Proc = False
	'// response.write "111"

	vQuery = " exec [db_order].[dbo].[sp_Ten_MyOrderStockOutItemCnt] '" & orderserial & "' "
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.eof then
		ChkStockoutItemExist_Proc = (rsget("stockoutCnt") > 0)
	end if
	rsget.close
end function

function OrderCancelValidMSG(myorder, myorderdetail, IsAllCancelProcess, IsPartCancelProcess, IsStockoutCancelProcess)
	dim IsCancelOK, CancelFailMSG
	IsCancelOK = True
	CancelFailMSG = ""

	if (IsCancelOK and (myorder.FResultCount < 1)) then
		IsCancelOK = False
		CancelFailMSG = "주문 내역이 없거나 취소된 거래건 입니다."
	elseif (IsCancelOK and (myorderdetail.FResultCount<1) and (myorder.FOneItem.Fipkumdiv >= "4")) then
		IsCancelOK = False
		CancelFailMSG = "배송비 추가결제건은 결제이후 취소할 수 없습니다."
	end if

	if IsCancelOK and IsStockoutCancelProcess = True then
		if ChkStockoutItemExist(myorderdetail) = False then
			IsCancelOK = False
			CancelFailMSG = "품절/택배파업취소 상품이 없습니다."
		end if
	end if

	if IsCancelOK and Not myorder.FOneItem.IsValidOrder then
		IsCancelOK = False
		CancelFailMSG = "취소된 주문입니다."
	end if

	if IsCancelOK = False then
		OrderCancelValidMSG = CancelFailMSG
		exit function
		''ShowAlertAndClosePopup(CancelFailMSG)
	end if

	if IsAllCancelProcess = True then
		if IsStockoutCancelProcess then
			'// 품절취소
			if IsCancelOK and Not myorder.FOneItem.IsWebStockOutItemCancelEnable then
				IsCancelOK = False
				CancelFailMSG = "웹취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
			end if

			if IsCancelOK and Not myorder.FOneItem.IsDirectStockOutPartialCancelEnable(myorderdetail) then
				IsCancelOK = False
				CancelFailMSG = "웹취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
			end if
		else
			'// 일반 전체취소
			if IsCancelOK and Not myorder.FOneItem.IsWebOrderCancelEnable then
				IsCancelOK = False
				CancelFailMSG = "잘못된 주문상태입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
				if (CStr(myorder.FOneItem.FIpkumdiv) = "6") then
					CancelFailMSG = "업체확인중인 상품이 있습니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
				end if
				if (CStr(myorder.FOneItem.FIpkumdiv) > "6") then
					CancelFailMSG = "이미 출고된 상품이 있습니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 취소 또는 반품을 문의주세요."
				end if
			end if

			if IsCancelOK and Not myorder.FOneItem.IsDirectALLCancelEnable(myorderdetail) then
				IsCancelOK = False
				CancelFailMSG = "웹취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
			end if
		end if
	elseif IsPartCancelProcess = True then
		'// 품절상품취소
		if IsCancelOK and myorder.FOneItem.FOrderSheetYN="P" then
			IsCancelOK = False
			CancelFailMSG = "선물포장 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
		end if

		if IsCancelOK and Not IsStockoutCancelProcess then
			IsCancelOK = False
			CancelFailMSG = "부분취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
		end if

		if IsCancelOK and IsStockoutCancelProcess and Not myorder.FOneItem.IsWebStockOutItemCancelEnable then
			IsCancelOK = False
			CancelFailMSG = "웹취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
		end if

		if IsCancelOK and IsStockoutCancelProcess and Not myorder.FOneItem.IsDirectStockOutPartialCancelEnable(myorderdetail) then
			IsCancelOK = False
			CancelFailMSG = "웹취소 불가 주문입니다. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 상담</font></a> 또는 고객센터로 문의주세요."
		end if

		if Not IsStockoutCancelProcess then
			ShowAlertAndClosePopup("잘못된 접근입니다.")
		end if
	else
		ShowAlertAndClosePopup("잘못된 접근입니다.")
	end if

	OrderCancelValidMSG = CancelFailMSG
end function

Function GetIsCancelOrderByOne(myorder, mode)
	'// 한번에 전체취소인지(취소금액 = 최초결제금액)
	dim vQuery, reducedPriceSUM
	reducedPriceSUM = 0
	vQuery = " select "
	if (mode = "stockoutcancel") or (mode = "socancelorder") then
		vQuery = vQuery & "		IsNull(sum(d.reducedPrice*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end)),0) as reducedPriceSUM "
	else
		vQuery = vQuery & "		IsNull(sum(d.reducedPrice*d.itemno),0) as reducedPriceSUM "
	end if
	vQuery = vQuery & "	from "
	vQuery = vQuery & "		[db_order].[dbo].[tbl_order_detail] d "
	vQuery = vQuery & "		left join db_temp.dbo.tbl_mibeasong_list m "
	vQuery = vQuery & "		on "
	vQuery = vQuery & "			d.idx = m.detailidx "
	vQuery = vQuery & "	where "
	vQuery = vQuery & "		1 = 1 "
	vQuery = vQuery & "		and d.orderserial = '" & myorder.FRectOrderserial & "' "
	vQuery = vQuery & "		and d.cancelyn <> 'Y' "
	vQuery = vQuery + " 	and IsNull(d.currstate, '0') < '7' "
	if (mode = "stockoutcancel") or (mode = "socancelorder") then
		vQuery = vQuery & "		and ( "
		'vQuery = vQuery & "			((d.itemid <> 0) and (IsNull(m.code, '') = '05')) "
		vQuery = vQuery & "			((d.itemid <> 0) and (IsNull(m.code, '') in ('05','06'))) "
		vQuery = vQuery & "			or "
		vQuery = vQuery & "			((d.itemid = 0) and (d.makerid in ( "
		vQuery = vQuery & "				select "
		vQuery = vQuery & "					(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) as makerid "
		vQuery = vQuery & "				from "
		vQuery = vQuery & "				[db_order].[dbo].[tbl_order_detail] d "
		vQuery = vQuery & "				left join db_temp.dbo.tbl_mibeasong_list m "
		vQuery = vQuery & "				on "
		vQuery = vQuery & "					d.idx = m.detailidx "
		vQuery = vQuery & "				where "
		vQuery = vQuery & "					1 = 1 "
		vQuery = vQuery & "					and d.orderserial = '" & myorder.FRectOrderserial & "' "
		vQuery = vQuery & "					and d.cancelyn <> 'Y' "
		vQuery = vQuery & "				group by "
		vQuery = vQuery & "					(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) "
		vQuery = vQuery & "				having "
		'vQuery = vQuery & "					sum(case when d.itemid <> 0 then 1 else 0 end) = sum(case when d.itemid <> 0 and IsNull(m.code, '') = '05' then 1 else 0 end) "
		vQuery = vQuery & "					sum(case when d.itemid <> 0 then 1 else 0 end) = sum(case when d.itemid <> 0 and IsNull(m.code, '') in ('05','06') then 1 else 0 end) "
		vQuery = vQuery & "			))) "
		vQuery = vQuery & "		) "
	end if
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.Eof Then
		reducedPriceSUM = rsget("reducedPriceSUM")
	End IF
	rsget.close

	Dim vPrice : vPrice = 0
	vQuery = "select IsNull(sum(acctamount),0) as acctamount from [db_order].[dbo].[tbl_order_PaymentEtc] "
	vQuery = vQuery & "where orderserial = '" & myorder.FRectOrderserial & "'"
	rsget.CursorLocation = adUseClient
	rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.Eof Then
		vPrice = rsget("acctamount")
	End IF
	rsget.close

	GetIsCancelOrderByOne = (reducedPriceSUM = (vPrice + myorder.FOneItem.FMileTotalPrice))
end function

function GetValidReturnMethod(myorder, IsCancelOrderByOne)
	GetValidReturnMethod = "R000"

	if Not myorder.FOneItem.IsPayed then
		exit function
	end if

	dim mainpaymentorg, cardPartialCancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcodeall
	cardPartialCancelok = "N"

	select case myorder.FOneItem.Faccountdiv
		case "100"
			'// 신용카드(일반, 네이버페이, 페이코)
			if IsCancelOrderByOne then
				GetValidReturnMethod = "R100"
			else
				Call myorder.getMainPaymentInfo(myorder.FOneItem.Faccountdiv, mainpaymentorg, cardPartialCancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcodeall)
				if cardPartialCancelok = "Y" then
					GetValidReturnMethod = "R120"
				else
					GetValidReturnMethod = "FAIL"
				end if
			end if
		case "400"
			'// 휴대폰
			if IsCancelOrderByOne then
				if DateDiff("m", myorder.FOneItem.FIpkumDate, Now) <= 0 then
					'// 이번달 결제
					GetValidReturnMethod = "R400"
				else
					GetValidReturnMethod = "R007"
				end if
			else
				GetValidReturnMethod = "R007"
			end if
		case "20"
			'// 실시간(일반, 네이버페이)
			if IsCancelOrderByOne then
				GetValidReturnMethod = "R020"
			else
				Call myorder.getMainPaymentInfo(myorder.FOneItem.Faccountdiv, mainpaymentorg, cardPartialCancelok, cardcancelerrormsg, cardcancelcount, cardcancelsum, cardcodeall)
				if cardPartialCancelok = "Y" then
					GetValidReturnMethod = "R022"
				else
					GetValidReturnMethod = "FAIL"
				end if
			end if
		case "7"
			'// 무통장
			GetValidReturnMethod = "R007"
		case "50"
			'// 외부몰 환불
			GetValidReturnMethod = "R050"
		case "14"
			'// 편의점결제
			GetValidReturnMethod = "R007"	' 편의점결제는 무통장 환불한다.
		case else
			'// 기타
			GetValidReturnMethod = "FAIL"
	end select
end function

function ShowAlertAndClosePopup(msg)
    response.write "<script language='javascript'>alert(' " & msg & "');</script>"
    response.write "<script language='javascript'>window.close();</script>"
    dbget.close()	:	response.End
end function

%>
