<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ �ڵ���� ó��
' History : 2020.10.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/maillib2.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/cscenter/lib/CSFunction.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cancelOrderLib.asp" -->
<!-- #include virtual="/lib/classes/ordercls/sp_myordercls.asp" -->
<%
'///////// �ش��������� ��� ���� �����ؼ� ������ �о� ����� �ؿ� �������� �ݵ�� ��� ���� �����ؾ� �մϴ�.
' WAPI : /autojob/cs_cancel_autojob.asp , �������, ����lib, ����Ŭ����
' WWW : /my10x10/orderPopup/CancelOrder_process.asp , �������, ����lib, ����Ŭ����
' M : /my10x10/order/CancelOrder_process.asp , �������, ����lib, ����Ŭ����
' APP : /apps/appCom/wish/web2014/my10x10/order/CancelOrder_process.asp , �������, ����lib, ����Ŭ����
'////////////////////////////////////
'dbget.Close() : response.end
dim webImgUrl : webImgUrl		= "http://webimage.10x10.co.kr"
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
	response.write "���� IP�� �ƴմϴ�."
    dbget.Close() : response.end
end if

dim mode     : mode = requestCheckVar(request("mode"),32)
dim sqlStr, i, j, arrOrderserial, orderserial, itemCnt, itemName, orderdate, sitename, mastertelno
dim rebankownername, encmethod, encaccount
dim mibeasongidxArr, buyhp, smstext, successCnt, cancelmode, myorder, myorderdetail, arruserid, userid, rebankname, rebankaccount
dim IsChangeOrder, IsCancelOK, CancelFailMSG, IsCancelOrderByOne, validReturnMethod, vIsMobileCancelDateUpDown, isCsMailSend
dim IsAllCancelProcess, IsPartCancelProcess, IsStockoutCancelProcess, isEvtGiftDisplay, ismoneyrefundok, IsSoldOutCancel, contents_finish
dim orgsubtotalprice, orgitemcostsum, orgbeasongpay, orgmileagesum, orgcouponsum, orgallatdiscountsum, gubun01, gubun02, contents_jupsu
Dim remainsubtotalprice, remainitemcostsum, remainbeasongpay, remainmileagesum, remaincouponsum, remainallatdiscountsum, remaindepositsum, remaingiftcardsum
dim retVal,IsCyberAcctCancel, vIsPacked, vQuery, orderusingmsg, ScanErr, errcode, CsId, ResultMsg, ipkumdiv, reguserid, finishuser, title
dim refundgiftcardsum, refunddepositsum, returnmethod, orgdepositsum, orggiftcardsum, refundrequire, newasid, modeflag2, divcd, id
Dim canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, refundallatsubtractsum, refundbeasongpay, refunddeliverypay, refundadjustpay, paygatetid
dim CancelValidResultMessage, intloop
dim finishArrOrderserial
'���� �����
dim copyitemcouponinfo, resultItemCouponCount
resultItemCouponCount=0
Const CFINISH_SYSTEM = "system"
	successCnt = 0
	intloop=0
	rebankname=""
	rebankownername=""
	encaccount=""
	rebankaccount=""

select Case mode
	'// ǰ����ǰ/�ù��ľ� �ڵ����(����)
    Case "cssoldoutitemcancel"
		'////////////////// ǰ����ǰ/�ù��ľ� �ڵ����(����)
		successCnt=0
		arrOrderserial=""
		arruserid=""
        finishArrOrderserial = ""
		' �� ���� ������. /cscenter/lib/csAsfunction.asp �� �Լ� RegmibesongCanceldate�� ���� ������ �ּ���. ������ �����ɰ�� �ڵ���ҿϷ� ó���� ���� �ʽ��ϴ�.
        sqlStr = " select distinct top 100 l.orderserial"
		sqlStr = sqlStr + " , isnull((select userid from [db_order].[dbo].[tbl_order_master] with (nolock) where l.orderserial=orderserial),'') as userid"
		sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		'sqlStr = sqlStr + " where l.code = '05' "
		sqlStr = sqlStr + " where l.code in ('05','06') "
		sqlStr = sqlStr + " 	and l.state <= '4' "
		sqlStr = sqlStr + " 	and l.isSendSMS = 'Y' "
		sqlStr = sqlStr + " 	and l.isSendEmail = 'Y' "
		sqlStr = sqlStr + " 	and l.sendCount > 0 "
		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		'sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('10x10_cs')) " '// ���� : RegmibesongCanceldate() �� ���� ������ ��!!!
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		sqlStr = sqlStr + " 	and l.isSendSMSdate is not null"
		sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) > 24"	' ���ڹ߼۵���24�ð�������
		'sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) < 72"	' ���ڹ߼۵��� 3�� �������� ������ �ʴ´�
		'sqlStr = sqlStr + " 	and d.isupchebeasong='N'"
		sqlStr = sqlStr + " 	and l.isautocanceldate is null"		' �ڵ���ҾȵȰ�
		sqlStr = sqlStr + " group by l.orderserial "
		sqlStr = sqlStr + " order by l.orderserial asc"

		' �׽�Ʈ�� ��ü 16033187645 , �Ϻ� 15031855057 20091692316

		''response.write sqlStr
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        if Not rsget.Eof then
            do until rsget.eof
            arrOrderserial = arrOrderserial & rsget("orderserial") & ","
			if rsget("userid")="" then
				arruserid = arruserid & "GuestOrder,"
			else
				arruserid = arruserid & rsget("userid") & ","
			end if
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if Right(arrOrderserial,1)="," then arrOrderserial=Left(arrOrderserial,Len(arrOrderserial)-1)
		if Right(arruserid,1)="," then arruserid=Left(arruserid,Len(arruserid)-1)

        arrOrderserial = split(arrOrderserial,",")
		arruserid = split(arruserid,",")

        if UBound(arrOrderserial)>-1 then
            for intloop=0 to UBound(arrOrderserial)
				orderusingmsg=""
				userid=""
				orderserial=""
				ResultMsg=""
				orderserial = arrOrderserial(intloop)
				userid = arruserid(intloop)

                response.write "<br />�ֹ���ȣ :" & orderserial & "<br />"

				cancelmode = "stockoutcancel"
				if IsAllStockOutCancel(orderserial) = True then
					cancelmode = "socancelorder"
				end if

                response.write "��Ҹ�� :" & cancelmode & "<br />"

				IsAllCancelProcess = (cancelmode = "socancelorder")
				IsPartCancelProcess = (cancelmode = "stockoutcancel")
				IsStockoutCancelProcess = ((cancelmode = "socancelorder") or (cancelmode = "stockoutcancel"))
				isEvtGiftDisplay = IsAllCancelProcess

				set myorder = new CMyOrder
				if userid="GuestOrder" then
					''��ȸ���ֹ�
					myorder.FRectOrderserial = orderserial
					if (orderserial<>"") then
						myorder.GetOneOrder
					end if
				else
					''ȸ���ֹ�
					myorder.FRectUserID = userid
					myorder.FRectOrderserial = orderserial

					if (userid<>"") and (orderserial<>"") then
						myorder.GetOneOrder
					end if
				end if

				IsChangeOrder = myorder.FOneItem.Fjumundiv = "6"

				set myorderdetail = new CMyOrder
				myorderdetail.FRectOrderserial = orderserial

				if (myorder.FResultCount>0) then
					myorderdetail.GetOrderDetail
				end if

				IsCancelOK = True
				CancelFailMSG = ""

				'// �ֹ����� üũ
				CancelFailMSG = OrderCancelValidMSG(myorder, myorderdetail, IsAllCancelProcess, IsPartCancelProcess, IsStockoutCancelProcess)
				if CancelFailMSG <> "" then
					IsCancelOK = False
				end if

                response.write "CancelFailMSG : " & CancelFailMSG & "<br />"

				'// ============================================================================
				'// ȯ�� ��������
				IsCancelOrderByOne = False
				if IsCancelOK then
					'// �ѹ� �ֹ� ��ü�������
					IsCancelOrderByOne = GetIsCancelOrderByOne(myorder, cancelmode) and Not IsPartCancelProcess
				end if

                if IsCancelOrderByOne then
                    response.write "IsCancelOrderByOne : Y" & "<br />"
                else
                    response.write "IsCancelOrderByOne : N" & "<br />"
                end if

				validReturnMethod = "R000"
				if IsCancelOK then
					validReturnMethod = GetValidReturnMethod(myorder, IsCancelOrderByOne)
				end if

                response.write "validReturnMethod : " & validReturnMethod & "<br />"

				if (validReturnMethod = "FAIL") then
					IsCancelOK = False
					CancelFailMSG = "����� �Ұ� �ֹ��Դϴ�."
				end if

				rebankname=""
				rebankownername=""
				encaccount=""
				rebankaccount=""
				if userid="GuestOrder" then
					''��ȸ���ֹ�
					if (orderserial<>"") then
						' ������ �ϰ�� ���� ȯ�Ұ��¸� �޾ƿ´�.
						if validReturnMethod = "R007" then
							returnmethod = "R007"
							fnSoldOutMyRefundInfo orderserial, rebankname, rebankownername, encaccount
							if isnull(encaccount) then encaccount = ""
							if myorder.FOneItem.FAccountDiv <> "7" then encaccount = ""		' ���������� ������ �ϰ�쿡�� ȯ�Ұ��¸� �����´�.
							rebankaccount = encaccount
						end if
					end if
				else
					''ȸ���ֹ�
					if (userid<>"") then
						' ������ �ϰ�� ���� ȯ�Ұ��¸� �޾ƿ´�.
						if validReturnMethod = "R007" then
							returnmethod = "R007"
							fnSoldOutMyRefundInfo userid, rebankname, rebankownername, encaccount
							if isnull(encaccount) then encaccount = ""
							if myorder.FOneItem.FAccountDiv <> "7" then encaccount = ""		' ���������� ������ �ϰ�쿡�� ȯ�Ұ��¸� �����´�.
							rebankaccount = encaccount
						end if
					end if
				end if

				'// �ڵ��� ���� ����ϰ� ������ ��. UP�� ��ҿ��� ���������� ��
				If myorder.FOneItem.Faccountdiv = "400" AND DateDiff("m", myorder.FOneItem.FIpkumDate, Now) > 0 Then
					vIsMobileCancelDateUpDown = "UP"
				Else
					vIsMobileCancelDateUpDown = "DOWN"
				End If

				if IsCancelOK then
					if validReturnMethod = "R007" then
						if (returnmethod <> "R007") and (returnmethod <> "R910") and (returnmethod <> "R000") then
							orderusingmsg="�߸��� �����Դϴ�.(ȯ�ҹ�� ����[0])"
						end if

                        '// ������ȯ���� ��� ��ġ��ȯ�ҷ� ��ȯ, 2021-01-18, skyer9
                        if (userid <> "GuestOrder") and ((rebankname = "") or (rebankownername = "") or (rebankaccount = "")) then
                            returnmethod = "R910"
                        end if
					else
						returnmethod = validReturnMethod
					end if
				else
					orderusingmsg="�߸��� �����Դϴ�.(ȯ�ҹ�� ����[1])"
				end if

                response.write " : returnmethod" & returnmethod & "<br />"
                response.write "validReturnMethod : " & validReturnMethod & "<br />"

				ismoneyrefundok = false
				if returnmethod = "R007" then
					ismoneyrefundok = true
				end if

				'### ǰ����ҽ� 1�� �ֹ��� ��ü��ǰ�� ǰ���� ��� cancelorder ��ü��� �� �¿�.
				IsSoldOutCancel = false
				if (cancelmode = "stockoutcancel") or (cancelmode = "socancelorder") then
					IsSoldOutCancel = true
				end if

				''�޴��� ���� �߰� 2015/04/21 IsINIMobile
				Dim IsINIMobile : IsINIMobile = false
				if (myorder.FOneItem.Faccountdiv = "400") and (Len(myorder.FOneItem.Fpaygatetid)=40) then
					IsINIMobile = (LEFT(myorder.FOneItem.Fpaygatetid,LEN("IniTechPG_"))="IniTechPG_") or (LEFT(myorder.FOneItem.Fpaygatetid,LEN("INIMX_HPP_"))="INIMX_HPP_") or (LEFT(myorder.FOneItem.Fpaygatetid,LEN("StdpayHPP_"))="StdpayHPP_")
				end if

				Dim IsDacomMobile : IsDacomMobile = false
				if (NOT IsINIMobile) then
					if (myorder.FOneItem.Faccountdiv = "400") and (Len(myorder.FOneItem.Fpaygatetid)>=31) then
						IsDacomMobile = True        ''46~49 Tradeid(23) & "|" & vTID(24)
					else
						IsDacomMobile = False       ''32~35 Tradeid(23) & "|" & vTID(10)
					end if
				end if

				'// ���ֹ�
				orgsubtotalprice		= myorder.FOneItem.Fsubtotalprice
				orgitemcostsum			= myorder.FOneItem.Ftotalsum - myorder.FOneItem.FDeliverprice
				orgbeasongpay			= myorder.FOneItem.FDeliverPrice
				orgmileagesum			= myorder.FOneItem.FMileTotalPrice
				orgcouponsum			= myorder.FOneItem.FTenCardSpend
				orgallatdiscountsum		= myorder.FOneItem.FAllatDiscountPrice
				orgdepositsum			= myorder.FOneItem.Fspendtencash
				orggiftcardsum			= myorder.FOneItem.Fspendgiftmoney
				paygatetid				= myorder.FOneItem.Fpaygatetid

				remainsubtotalprice		= orgsubtotalprice
				remainitemcostsum		= orgitemcostsum
				remainbeasongpay		= orgbeasongpay
				remainmileagesum		= orgmileagesum
				remaincouponsum			= orgcouponsum
				remainallatdiscountsum	= orgallatdiscountsum
				remaindepositsum		= orgdepositsum
				remaingiftcardsum		= orggiftcardsum

				refunditemcostsum		= 0
				refundmileagesum		= 0
				refundcouponsum			= 0
				refundallatsubtractsum	= 0
				refundbeasongpay		= 0
				refunddeliverypay		= 0
				refundadjustpay			= 0
				refundgiftcardsum		= 0
				refunddepositsum		= 0

				'������ �����������.
				IsCyberAcctCancel = myorder.FOneItem.IsDacomCyberAccountPay
				IsCyberAcctCancel = IsCyberAcctCancel And (Not myorder.FOneItem.IsPayed)

				vIsPacked = (myorder.FOneItem.FOrderSheetYN="P")

				if (cancelmode="socancelorder") then
					'' ��ü ���
					vQuery = " select "
					vQuery = vQuery & "		sum(case when d.itemid <> 0 then d.itemcost*d.itemno else 0 end) as refunditemcostsum "
					vQuery = vQuery & "		, sum(d.itemcost*d.itemno - (d.reducedPrice + IsNull(d.etcDiscount,0))*d.itemno) as refundcouponsum "
					vQuery = vQuery & "		, sum(IsNull(d.etcDiscount,0)*d.itemno) as refundallatsubtractsum "
					vQuery = vQuery & "		, sum(case when d.itemid = 0 then d.itemcost*d.itemno else 0 end) as refundbeasongpay "
					vQuery = vQuery & "	from "
					vQuery = vQuery & "		[db_order].[dbo].[tbl_order_detail] d with (nolock)"
					vQuery = vQuery & "	where "
					vQuery = vQuery & "		1 = 1 "
					vQuery = vQuery & "		and d.orderserial = '" & orderserial & "' "
					vQuery = vQuery & "		and d.cancelyn <> 'Y' "
					rsget.CursorLocation = adUseClient
					rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
					If not rsget.Eof Then
						refunditemcostsum = rsget("refunditemcostsum")
						refundcouponsum = rsget("refundcouponsum")
						refundallatsubtractsum = rsget("refundallatsubtractsum")
						refundbeasongpay = rsget("refundbeasongpay")
					End IF
					rsget.close

					if (remainitemcostsum < refunditemcostsum) or (remaincouponsum < refundcouponsum) or (remainbeasongpay < refundbeasongpay) then
						orderusingmsg="������� �� �� �����ϴ�.[�ڵ��ȣ:3-3]"
					end if

					'��Ÿ����, �ۼ�Ʈ���� �翬����
					refundrequire = refunditemcostsum - refundallatsubtractsum - refundcouponsum + refundbeasongpay

					'���ϸ���, ��ġ��, ����Ʈī�� ����
					'// 2018-02-22, skyer9, ���ϸ��� �̹� ��������.
					remainsubtotalprice = remainsubtotalprice - 0 - remaindepositsum - remaingiftcardsum

					'���ϸ���
					if (remainsubtotalprice < refundrequire) then
						if (remainmileagesum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remainmileagesum) then
								refundmileagesum = remainmileagesum
							else
								refundmileagesum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refundmileagesum
						end if
					end if

					'����Ʈī��
					if (remainsubtotalprice < refundrequire) then
						if (remaingiftcardsum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remaingiftcardsum) then
								refundgiftcardsum = remaingiftcardsum
							else
								refundgiftcardsum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refundgiftcardsum
						end if
					end if

					'��ġ��
					if (remainsubtotalprice < refundrequire) then
						if (remaindepositsum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remaindepositsum) then
								refunddepositsum = remaindepositsum
							else
								refunddepositsum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refunddepositsum
						end if
					end if

					if (remainsubtotalprice < refundrequire) then
						orderusingmsg="������� �� �� �����ϴ�.[�ڵ��ȣ:4-1]"
					end if
					if refundrequire < 0 then
						orderusingmsg="������� �� �� �����ϴ�.[�ڵ��ȣ:4-2]"
					end if

					canceltotal = refundrequire
					newasid 		= -1
					modeflag2   	= "regcsas"
					divcd       	= "A008"
					id          	= 0
					ipkumdiv    	= myorder.FOneItem.FIpkumDiv
					reguserid   	= CFINISH_SYSTEM
					finishuser  	= CFINISH_SYSTEM
					title       	= "[�ڵ���ü���]" & GetDefaultTitle(divcd, 0, orderserial)
					gubun01     	= "C004"  ''����

					If IsSoldOutCancel Then
						gubun02     	= "CD05"  ''ǰ��
					Else
						gubun02     	= "CD01"  ''�ܼ�����
					End If

					contents_jupsu  = ""
					contents_finish = ""
					isCsMailSend 	= "on"
					refundrequire	= myorder.FOneItem.Fsubtotalprice - myorder.FOneItem.FsumPaymentEtc
					if (myorder.FOneItem.Fipkumdiv < 4) then
						refundrequire = "0"
					end if

					if orderusingmsg="" then
						dbget.beginTrans

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "001"
							'' CS Master ����
							CsId = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "002"
							'' CS Detail ����
							Call RegWebCSDetailAllCancel(CsId, orderserial)
						end if

						' �ڵ���ҳ�¥�� �ִ´�.
						RegmibesongCanceldate(orderserial)

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "003"
							'' ȯ�� �������� (��)����
							'// ������ ����Ѵ�. 2019-01-10, skyer9
							''if (refundrequire<>"0") and (returnmethod<>"R000") then
								refundcouponsum = refundcouponsum * -1
								refundmileagesum = refundmileagesum * -1
								refundgiftcardsum = refundgiftcardsum * -1
								refunddepositsum = refunddepositsum * -1

								'CS Master ȯ�� �������� ����	''# RegCSMasterRefundInfo, AddCSMasterRefundInfo -> /cscenter/lib/csAsfunction.asp
								Call RegCSMasterRefundInfo(CsId, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, refundallatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
								Call AddCSMasterRefundInfo(CsId, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

								'''���� ��ȣȭ �߰�.
								Call EditCSMasterRefundEncInfo(CsId, encmethod, rebankaccount)
							''end if
						End if

						' ��ǰ����ȯ�޿���
						if itemCouponRefundYN="Y" and userid<>"GuestOrder" and userid<>"" then
							If (Err.Number = 0) and (ScanErr="") Then
								errcode = "004"

								' �ֹ� ��ǰ���� ����� üũ
								resultItemCouponCount = ItemCouponCount(CsId, "P", userid)
								if resultItemCouponCount>0 then
									copyitemcouponinfo="Y"
								else
									copyitemcouponinfo="N"
								end if

								' ��ǰ���� ���������
								Call EditCSCopyItemCouponInfo(CsId, copyitemcouponinfo)
							end if
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							dbget.CommitTrans

                            finishArrOrderserial = finishArrOrderserial + "," + orderserial
							successCnt = successCnt + 1
							response.write "��ü��������Ϸ� : " & orderserial & "<br /><br />"

							'########################################### �������� ���� ���. ��ü��Ҹ� ��. ###########################################
							If vIsPacked = "Y" Then
								sqlStr = "UPDATE [db_order].[dbo].[tbl_order_pack_master] SET cancelyn = 'Y' WHERE orderserial = '" & orderserial & "' " & vbCrLf
								sqlStr = sqlStr & "UPDATE [db_order].[dbo].[tbl_order_pack_detail] SET cancelyn = 'Y' "
								sqlStr = sqlStr & "WHERE midx IN(select midx from [db_order].[dbo].[tbl_order_pack_master] where orderserial = '" & orderserial & "')"
								dbget.Execute sqlStr
							End If
							'########################################### �������� ���� ���. ��ü��Ҹ� ��. ###########################################

						Else
							dbget.RollBackTrans
							response.write "����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�. ������ ���� ���.(�����ڵ� : " & CStr(errcode) & ")" & "<br /><br />"
						End If
					else
						response.write orderserial & " : " & orderusingmsg & "<br /><br />"
					end if

				'### ǰ��/�ù��ľ� ���(�κ�����϶�) ���μ��� ###
				elseif (cancelmode="stockoutcancel") then
					If vIsPacked = "Y" Then
						orderusingmsg="���������� �ֹ��� �Դϴ�.\n�������� �ֹ��� ��ü��Ҹ� �����մϴ�."
					End If

					vQuery = " select "
					vQuery = vQuery & "		sum(case when d.itemid <> 0 then d.itemcost*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end) else 0 end) as refunditemcostsum "
					vQuery = vQuery & "		, sum(d.itemcost*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end) - (d.reducedPrice + IsNull(d.etcDiscount,0))*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end)) as refundcouponsum "
					vQuery = vQuery & "		, sum(IsNull(d.etcDiscount,0)*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end)) as refundallatsubtractsum "
					vQuery = vQuery & "		, sum(case when d.itemid = 0 then d.itemcost*(case when d.itemid = 0 then d.itemno else IsNull(m.itemlackno,0) end) else 0 end) as refundbeasongpay "
					vQuery = vQuery & "	from "
					vQuery = vQuery & "		[db_order].[dbo].[tbl_order_detail] d with (nolock)"
					vQuery = vQuery & "		left join db_temp.dbo.tbl_mibeasong_list m with (nolock)"
					vQuery = vQuery & "		on "
					vQuery = vQuery & "			d.idx = m.detailidx "
					vQuery = vQuery & "	where "
					vQuery = vQuery & "		1 = 1 "
					vQuery = vQuery & "		and d.orderserial = '" & orderserial & "' "
					vQuery = vQuery & "		and d.cancelyn <> 'Y' "
					vQuery = vQuery + " 	and IsNull(d.currstate, '0') < '7' "
					vQuery = vQuery & " 	and ((IsNull(m.itemlackno,0) > 0) or (d.itemid = 0)) "
					vQuery = vQuery & "		and ( "
					'vQuery = vQuery & "			((d.itemid <> 0) and (IsNull(m.code, '') = '05')) "					'// ����� ǰ���� �ڵ����, skyer9, 2020-11-10
					vQuery = vQuery & "				((d.itemid <> 0) and (IsNull(m.code, '') in ('05','06'))) "
					vQuery = vQuery & "			or "
					vQuery = vQuery & "			((d.itemid = 0) and (d.makerid in ( "
					vQuery = vQuery & "				select "
					vQuery = vQuery & "					(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) as makerid "
					vQuery = vQuery & "				from "
					vQuery = vQuery & "				[db_order].[dbo].[tbl_order_detail] d with (nolock)"
					vQuery = vQuery & "				left join db_temp.dbo.tbl_mibeasong_list m with (nolock)"
					vQuery = vQuery & "				on "
					vQuery = vQuery & "					d.idx = m.detailidx "
					vQuery = vQuery & "				where "
					vQuery = vQuery & "					1 = 1 "
					vQuery = vQuery & "					and d.orderserial = '" & orderserial & "' "
					vQuery = vQuery & "					and d.cancelyn <> 'Y' "
					vQuery = vQuery & "				group by "
					vQuery = vQuery & "					(case when d.isupchebeasong = 'Y' or d.itemid = 0 then d.makerid else '' end) "
					vQuery = vQuery & "				having "
					'vQuery = vQuery & "					sum(case when d.itemid <> 0 then d.itemno else 0 end) = sum(case when d.itemid <> 0 and IsNull(m.code, '') = '05' then IsNull(m.itemlackno,0) else 0 end) "
					vQuery = vQuery & "					sum(case when d.itemid <> 0 then d.itemno else 0 end) = sum(case when d.itemid <> 0 and IsNull(m.code, '') in ('05','06') then IsNull(m.itemlackno,0) else 0 end) "
					vQuery = vQuery & "			))) "
					vQuery = vQuery & "		) "
					rsget.CursorLocation = adUseClient
					rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
					If not rsget.Eof Then
						refunditemcostsum = rsget("refunditemcostsum")
						refundcouponsum = rsget("refundcouponsum")
						refundallatsubtractsum = rsget("refundallatsubtractsum")
						refundbeasongpay = rsget("refundbeasongpay")
					End IF
					rsget.close

					if (remainitemcostsum < refunditemcostsum) or (remaincouponsum < refundcouponsum) or (remainbeasongpay < refundbeasongpay) then
						orderusingmsg="ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:3-3]"
					end if

					'��Ÿ����, �ۼ�Ʈ���� �翬����
					refundrequire = refunditemcostsum - refundallatsubtractsum - refundcouponsum + refundbeasongpay

					'���ϸ���, ��ġ��, ����Ʈī�� ����
					'// 2018-02-22, skyer9, ���ϸ��� �̹� ��������.
					remainsubtotalprice = remainsubtotalprice - 0 - remaindepositsum - remaingiftcardsum

					'���ϸ���
					if (remainsubtotalprice < refundrequire) then
						if (remainmileagesum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remainmileagesum) then
								refundmileagesum = remainmileagesum
							else
								refundmileagesum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refundmileagesum
						end if
					end if

					'����Ʈī��
					if (remainsubtotalprice < refundrequire) then
						if (remaingiftcardsum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remaingiftcardsum) then
								refundgiftcardsum = remaingiftcardsum
							else
								refundgiftcardsum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refundgiftcardsum
						end if
					end if

					'��ġ��
					if (remainsubtotalprice < refundrequire) then
						if (remaindepositsum > 0) then
							if ((refundrequire - remainsubtotalprice) >= remaindepositsum) then
								refunddepositsum = remaindepositsum
							else
								refunddepositsum = (refundrequire - remainsubtotalprice)
							end if
							refundrequire = refundrequire - refunddepositsum
						end if
					end if

					if (remainsubtotalprice < refundrequire) then
						orderusingmsg="ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:4-1]"
					end if
					if refundrequire < 0 then
						orderusingmsg="ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:4-2]"
					end if

					canceltotal = refundrequire

					newasid 		= -1

					modeflag2   	= "regcsas"
					divcd       	= "A008"
					id          	= 0
					ipkumdiv    	= myorder.FOneItem.FIpkumDiv
					reguserid   	= CFINISH_SYSTEM
					finishuser  	= CFINISH_SYSTEM
					title       	= "[�ڵ��κ����]" & GetDefaultTitle(divcd, 0, orderserial)
					gubun01     	= "C004"  ''����
					gubun02     	= "CD05"  ''ǰ��
					ScanErr = ""

					contents_jupsu  = ""
					contents_finish = ""
					isCsMailSend 	= "on"

					if (myorder.FOneItem.Fipkumdiv < 4) then
						refundrequire = "0"
					end if

					if (reguserid = "") then
						reguserid="GuestOrder"
					end if

					if orderusingmsg="" then
						dbget.beginTrans

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "001"
							'' CS Master ����
							id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "002"
							'' CS Detail ����
							Call RegWebCSDetailStockoutCancel(id, orderserial)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "003"
							'' ȯ�� �������� (��)����

							'// ������ ����Ѵ�. 2019-01-10, skyer9
							''if (refundrequire<>"0") and (returnmethod<>"R000") then
								refundcouponsum = refundcouponsum * -1
								refundmileagesum = refundmileagesum * -1
								refundgiftcardsum = refundgiftcardsum * -1
								refunddepositsum = refunddepositsum * -1

								'CS Master ȯ�� �������� ����	''# RegCSMasterRefundInfo, AddCSMasterRefundInfo -> /cscenter/lib/csAsfunction.asp
								Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, refundallatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
								Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

								'''���� ��ȣȭ �߰�.
								Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
							''end if
						End if

						' ��ǰ����ȯ�޿���
						if itemCouponRefundYN="Y" and userid<>"GuestOrder" and userid<>"" then
							If (Err.Number = 0) and (ScanErr="") Then
								errcode = "004"

								' �ֹ� ��ǰ���� ����� üũ
								resultItemCouponCount = ItemCouponCount(id, "P", userid)
								if resultItemCouponCount>0 then
									copyitemcouponinfo="Y"
								else
									copyitemcouponinfo="N"
								end if

								' ��ǰ���� ���������
								Call EditCSCopyItemCouponInfo(id, copyitemcouponinfo)
							end if
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "005"

							CancelValidResultMessage = GetPartialCancelRegValidResult(id, orderserial)

							if (CancelValidResultMessage <> "") then
								ScanErr = CancelValidResultMessage
								response.write orderserial & " : ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:6]" & CancelValidResultMessage & "<br />"
							end if
						End If

						'���Ϸ� �Ǵ� ��ҵ� ��ǰ�� ���� ���, ��������(�ֹ���� �Ұ�)
						'���Ϸ�� ��ǰ�� ��ǰ�� �����ϴ�.
						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "006"

							''��� �Ϸ� �Ǵ� ��ҵ� ������ �ִ��� Ȯ��
							if Not (IsCancelValidState(id, orderserial)) then
								dbget.RollBackTrans
								response.write orderserial & " : ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:5]" & "<br />"
							end if
						end if

						'' �Ϸ�ó�� �ٷ� �������� ����
						'' ��ü Ȯ���� ���°� �ִ°�� - > �����θ� ����
						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "007"
							contents_finish = ""
						End If

						ResultMsg = ResultMsg + "->. [�ֹ� ��� CS] ����\n\n"

						If (Err.Number = 0) and (ScanErr="") Then
							'���Ϸ� �Ǵ� ��ҵ� ��ǰ�� ���� ���, ��������(�ֹ���� �Ұ�)
							'���Ϸ�� ��ǰ�� ��ǰ�� �����ϴ�.
							''��� �Ϸ� �Ǵ� ��ҵ� ������ �ִ��� Ȯ��
							if Not (IsCancelValidState(id, orderserial)) then
								errcode = "006"
								dbget.RollBackTrans
								response.write orderserial & " : ǰ��/�ù��ľ�������� �� �� �����ϴ�.[�ڵ��ȣ:5]" & "<br /><br />"
							else
								dbget.CommitTrans
								response.write "ǰ��/�ù��ľ���������Ϸ� : " & orderserial & "<br /><br />"

                                finishArrOrderserial = finishArrOrderserial + "," + orderserial
								successCnt = successCnt + 1
							end if

						Else
							dbget.RollBackTrans

							response.write orderserial & " : ǰ��/�ù��ľ�������� �� �� �����ϴ�.[99-"&errcode&"]" & "<br /><br />"
						End If
					else
						response.write orderserial & " : " & orderusingmsg & "<br /><br />"
					end if
				end if

			set myorder = Nothing
			set myorderdetail = Nothing
			next
		end if

        if (finishArrOrderserial = "") then
            finishArrOrderserial = "-"
        else
            finishArrOrderserial = Mid(finishArrOrderserial, 2, 2000)
            finishArrOrderserial = Replace(finishArrOrderserial, ",", "','")
        end if

		' �׽�Ʈ
 		'sqlStr="update db_temp.dbo.tbl_mibeasong_list set code = '05', itemlackno=1 , itemno=1,isSendSMSdate='2020-12-08 09:30:33.250', isSendEmaildate='2020-12-08 09:30:33.250', isautocanceldate=NULL ,isSendSMS = 'Y',isSendEmail = 'Y' , sendCount=1 where idx in (664264,664265,664266,665965,667111)"
		'response.write sqlStr
		'dbget.Execute sqlStr

		if successCnt>0 then
			response.write "<br>��ǰǰ��/�ù��ľ� " & successCnt & "�� �ڵ���� ���� ���." & "<br />"
		else
			response.write "0"
		end if

	case else
		dbget.Close()
		response.end
end select

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","121.78.103.60","192.168.1.67","192.168.1.73", "::1")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

function AddCsMemo(orderserial,divcd,userid,writeuser,contents_jupsu)
    dim sqlStr
    dim mmgubun ''�޸𱸺�
	dim phoneNumber, startPhoneIdx, endPhoneIdx
    if (LCase(LEFT(contents_jupsu,5))="[sms ") then
    	mmgubun = "4"
		startPhoneIdx = Len("[sms ") + 1
		endPhoneIdx = InStr(contents_jupsu, "]")
		if (endPhoneIdx > 0) and ((endPhoneIdx - startPhoneIdx) < 16) then
			phoneNumber = Mid(contents_jupsu, startPhoneIdx, (endPhoneIdx - startPhoneIdx))
		end if
	elseif (LCase(LEFT(contents_jupsu,5))="[mail") then
		mmgubun = "5"
	else
		mmgubun = "0"
	end if

	if divcd="1" then
		''�Ϲݸ޸�
		sqlStr = "insert into [db_cs].[dbo].[tbl_cs_memo]"
		sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate, phoneNumber)"
		sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate(), '" + CStr(phoneNumber) + "')"

		'response.write sqlStr
		dbget.Execute sqlStr
	else
		''ó����û�޸�
		sqlStr = "insert into [db_cs].[dbo].[tbl_cs_memo]"
		sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
		sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

		'response.write sqlStr
		dbget.Execute sqlStr
	end if
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
