<%
'// �ѱ� �ѱ� �ѱ�

'// ǰ���� ���� ��ü�������
Function IsAllStockOutCancel(orderserial)
	Dim vQuery, arr
	dim reducedPriceSUM, cancelReducedPriceSUM
	IsAllStockOutCancel = True

	if orderserial="" then exit Function

    '// ���Ϸ� ������ ǰ����ϵǾ� �־ ��һ�ǰ�� �ƴϴ�.

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
	vQuery = vQuery & " 	and d.cancelyn <> 'Y' "						'// ���� ���Ϸ�� ��ǰ ���� �ȵȴ�.(���� : /cscenter/lib/csAsfunction.asp : RegWebCSDetailAllCancel)
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
		CancelFailMSG = "�ֹ� ������ ���ų� ��ҵ� �ŷ��� �Դϴ�."
	elseif (IsCancelOK and (myorderdetail.FResultCount<1) and (myorder.FOneItem.Fipkumdiv >= "4")) then
		IsCancelOK = False
		CancelFailMSG = "��ۺ� �߰��������� �������� ����� �� �����ϴ�."
	end if

	if IsCancelOK and IsStockoutCancelProcess = True then
		if ChkStockoutItemExist(myorderdetail) = False then
			IsCancelOK = False
			CancelFailMSG = "ǰ��/�ù��ľ���� ��ǰ�� �����ϴ�."
		end if
	end if

	if IsCancelOK and Not myorder.FOneItem.IsValidOrder then
		IsCancelOK = False
		CancelFailMSG = "��ҵ� �ֹ��Դϴ�."
	end if

	if IsCancelOK = False then
		OrderCancelValidMSG = CancelFailMSG
		exit function
		''ShowAlertAndClosePopup(CancelFailMSG)
	end if

	if IsAllCancelProcess = True then
		if IsStockoutCancelProcess then
			'// ǰ�����
			if IsCancelOK and Not myorder.FOneItem.IsWebStockOutItemCancelEnable then
				IsCancelOK = False
				CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
			end if

			if IsCancelOK and Not myorder.FOneItem.IsDirectStockOutPartialCancelEnable(myorderdetail) then
				IsCancelOK = False
				CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
			end if
		else
			'// �Ϲ� ��ü���
			if IsCancelOK and Not myorder.FOneItem.IsWebOrderCancelEnable then
				IsCancelOK = False
				CancelFailMSG = "�߸��� �ֹ������Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
				if (CStr(myorder.FOneItem.FIpkumdiv) = "6") then
					CancelFailMSG = "��üȮ������ ��ǰ�� �ֽ��ϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
				end if
				if (CStr(myorder.FOneItem.FIpkumdiv) > "6") then
					CancelFailMSG = "�̹� ���� ��ǰ�� �ֽ��ϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� ��� �Ǵ� ��ǰ�� �����ּ���."
				end if
			end if

			if IsCancelOK and Not myorder.FOneItem.IsDirectALLCancelEnable(myorderdetail) then
				IsCancelOK = False
				CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
			end if
		end if
	elseif IsPartCancelProcess = True then
		'// ǰ����ǰ���
		if IsCancelOK and myorder.FOneItem.FOrderSheetYN="P" then
			IsCancelOK = False
			CancelFailMSG = "�������� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
		end if

		if IsCancelOK and Not IsStockoutCancelProcess then
			IsCancelOK = False
			CancelFailMSG = "�κ���� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
		end if

		if IsCancelOK and IsStockoutCancelProcess and Not myorder.FOneItem.IsWebStockOutItemCancelEnable then
			IsCancelOK = False
			CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
		end if

		if IsCancelOK and IsStockoutCancelProcess and Not myorder.FOneItem.IsDirectStockOutPartialCancelEnable(myorderdetail) then
			IsCancelOK = False
			CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�. <a href='javascript:GotoCSCenter()'><font color='blue'>1:1 ���</font></a> �Ǵ� �����ͷ� �����ּ���."
		end if

		if Not IsStockoutCancelProcess then
			ShowAlertAndClosePopup("�߸��� �����Դϴ�.")
		end if
	else
		ShowAlertAndClosePopup("�߸��� �����Դϴ�.")
	end if

	OrderCancelValidMSG = CancelFailMSG
end function

Function GetIsCancelOrderByOne(myorder, mode)
	'// �ѹ��� ��ü�������(��ұݾ� = ���ʰ����ݾ�)
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
			'// �ſ�ī��(�Ϲ�, ���̹�����, ������)
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
			'// �޴���
			if IsCancelOrderByOne then
				if DateDiff("m", myorder.FOneItem.FIpkumDate, Now) <= 0 then
					'// �̹��� ����
					GetValidReturnMethod = "R400"
				else
					GetValidReturnMethod = "R007"
				end if
			else
				GetValidReturnMethod = "R007"
			end if
		case "20"
			'// �ǽð�(�Ϲ�, ���̹�����)
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
			'// ������
			GetValidReturnMethod = "R007"
		case "50"
			'// �ܺθ� ȯ��
			GetValidReturnMethod = "R050"
		case "14"
			'// ����������
			GetValidReturnMethod = "R007"	' ������������ ������ ȯ���Ѵ�.
		case else
			'// ��Ÿ
			GetValidReturnMethod = "FAIL"
	end select
end function

function ShowAlertAndClosePopup(msg)
    response.write "<script language='javascript'>alert(' " & msg & "');</script>"
    response.write "<script language='javascript'>window.close();</script>"
    dbget.close()	:	response.End
end function

%>
