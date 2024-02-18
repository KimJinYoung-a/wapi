<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객센터 자동취소 처리
' History : 2020.10.27 한용민 생성
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
'///////// 해당페이지의 취소 로직 관련해서 수정이 읽어 날경우 밑에 페이지도 반드시 모두 같이 수정해야 합니다.
' WAPI : /autojob/cs_cancel_autojob.asp , 관련펑션, 관련lib, 관련클래스
' WWW : /my10x10/orderPopup/CancelOrder_process.asp , 관련펑션, 관련lib, 관련클래스
' M : /my10x10/order/CancelOrder_process.asp , 관련펑션, 관련lib, 관련클래스
' APP : /apps/appCom/wish/web2014/my10x10/order/CancelOrder_process.asp , 관련펑션, 관련lib, 관련클래스
'////////////////////////////////////
'dbget.Close() : response.end
dim webImgUrl : webImgUrl		= "http://webimage.10x10.co.kr"
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
	response.write "허용된 IP가 아닙니다."
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
'쿠폰 재발행
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
	'// 품절상품/택배파업 자동취소(접수)
    Case "cssoldoutitemcancel"
		'////////////////// 품절상품/택배파업 자동취소(접수)
		successCnt=0
		arrOrderserial=""
		arruserid=""
        finishArrOrderserial = ""
		' 요 쿼리 수정시. /cscenter/lib/csAsfunction.asp 에 함수 RegmibesongCanceldate도 같이 수정해 주세요. 한쪽이 누락될경우 자동취소완료 처리가 되지 않습니다.
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
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		'sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('10x10_cs')) " '// 주의 : RegmibesongCanceldate() 도 같이 수정할 것!!!
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		sqlStr = sqlStr + " 	and l.isSendSMSdate is not null"
		sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) > 24"	' 문자발송된지24시간지난거
		'sqlStr = sqlStr + " 	and datediff(hour,l.isSendSMSdate,getdate()) < 72"	' 문자발송된지 3일 지난건은 보내지 않는다
		'sqlStr = sqlStr + " 	and d.isupchebeasong='N'"
		sqlStr = sqlStr + " 	and l.isautocanceldate is null"		' 자동취소안된거
		sqlStr = sqlStr + " group by l.orderserial "
		sqlStr = sqlStr + " order by l.orderserial asc"

		' 테스트용 전체 16033187645 , 일부 15031855057 20091692316

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

                response.write "<br />주문번호 :" & orderserial & "<br />"

				cancelmode = "stockoutcancel"
				if IsAllStockOutCancel(orderserial) = True then
					cancelmode = "socancelorder"
				end if

                response.write "취소모드 :" & cancelmode & "<br />"

				IsAllCancelProcess = (cancelmode = "socancelorder")
				IsPartCancelProcess = (cancelmode = "stockoutcancel")
				IsStockoutCancelProcess = ((cancelmode = "socancelorder") or (cancelmode = "stockoutcancel"))
				isEvtGiftDisplay = IsAllCancelProcess

				set myorder = new CMyOrder
				if userid="GuestOrder" then
					''비회원주문
					myorder.FRectOrderserial = orderserial
					if (orderserial<>"") then
						myorder.GetOneOrder
					end if
				else
					''회원주문
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

				'// 주문상태 체크
				CancelFailMSG = OrderCancelValidMSG(myorder, myorderdetail, IsAllCancelProcess, IsPartCancelProcess, IsStockoutCancelProcess)
				if CancelFailMSG <> "" then
					IsCancelOK = False
				end if

                response.write "CancelFailMSG : " & CancelFailMSG & "<br />"

				'// ============================================================================
				'// 환불 가능한지
				IsCancelOrderByOne = False
				if IsCancelOK then
					'// 한방 주문 전체취소인지
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
					CancelFailMSG = "웹취소 불가 주문입니다."
				end if

				rebankname=""
				rebankownername=""
				encaccount=""
				rebankaccount=""
				if userid="GuestOrder" then
					''비회원주문
					if (orderserial<>"") then
						' 무통장 일경우 고객의 환불계좌를 받아온다.
						if validReturnMethod = "R007" then
							returnmethod = "R007"
							fnSoldOutMyRefundInfo orderserial, rebankname, rebankownername, encaccount
							if isnull(encaccount) then encaccount = ""
							if myorder.FOneItem.FAccountDiv <> "7" then encaccount = ""		' 결제수단이 무통장 일경우에만 환불계좌를 가져온다.
							rebankaccount = encaccount
						end if
					end if
				else
					''회원주문
					if (userid<>"") then
						' 무통장 일경우 고객의 환불계좌를 받아온다.
						if validReturnMethod = "R007" then
							returnmethod = "R007"
							fnSoldOutMyRefundInfo userid, rebankname, rebankownername, encaccount
							if isnull(encaccount) then encaccount = ""
							if myorder.FOneItem.FAccountDiv <> "7" then encaccount = ""		' 결제수단이 무통장 일경우에만 환불계좌를 가져온다.
							rebankaccount = encaccount
						end if
					end if
				end if

				'// 핸드폰 결제 취소일과 결제일 비교. UP이 취소월이 결제월보다 뒤
				If myorder.FOneItem.Faccountdiv = "400" AND DateDiff("m", myorder.FOneItem.FIpkumDate, Now) > 0 Then
					vIsMobileCancelDateUpDown = "UP"
				Else
					vIsMobileCancelDateUpDown = "DOWN"
				End If

				if IsCancelOK then
					if validReturnMethod = "R007" then
						if (returnmethod <> "R007") and (returnmethod <> "R910") and (returnmethod <> "R000") then
							orderusingmsg="잘못된 접근입니다.(환불방식 오류[0])"
						end if

                        '// 무통장환불일 경우 예치금환불로 전환, 2021-01-18, skyer9
                        if (userid <> "GuestOrder") and ((rebankname = "") or (rebankownername = "") or (rebankaccount = "")) then
                            returnmethod = "R910"
                        end if
					else
						returnmethod = validReturnMethod
					end if
				else
					orderusingmsg="잘못된 접근입니다.(환불방식 오류[1])"
				end if

                response.write " : returnmethod" & returnmethod & "<br />"
                response.write "validReturnMethod : " & validReturnMethod & "<br />"

				ismoneyrefundok = false
				if returnmethod = "R007" then
					ismoneyrefundok = true
				end if

				'### 품절취소시 1개 주문에 전체상품이 품절인 경우 cancelorder 전체취소 를 태움.
				IsSoldOutCancel = false
				if (cancelmode = "stockoutcancel") or (cancelmode = "socancelorder") then
					IsSoldOutCancel = true
				end if

				''휴대폰 결제 추가 2015/04/21 IsINIMobile
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

				'// 원주문
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

				'데이콤 가상계좌인지.
				IsCyberAcctCancel = myorder.FOneItem.IsDacomCyberAccountPay
				IsCyberAcctCancel = IsCyberAcctCancel And (Not myorder.FOneItem.IsPayed)

				vIsPacked = (myorder.FOneItem.FOrderSheetYN="P")

				if (cancelmode="socancelorder") then
					'' 전체 취소
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
						orderusingmsg="취소접수 할 수 없습니다.[코드번호:3-3]"
					end if

					'기타할인, 퍼센트쿠폰 당연차감
					refundrequire = refunditemcostsum - refundallatsubtractsum - refundcouponsum + refundbeasongpay

					'마일리지, 예치금, 기프트카드 제외
					'// 2018-02-22, skyer9, 마일리지 이미 빠져있음.
					remainsubtotalprice = remainsubtotalprice - 0 - remaindepositsum - remaingiftcardsum

					'마일리지
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

					'기프트카드
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

					'예치금
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
						orderusingmsg="취소접수 할 수 없습니다.[코드번호:4-1]"
					end if
					if refundrequire < 0 then
						orderusingmsg="취소접수 할 수 없습니다.[코드번호:4-2]"
					end if

					canceltotal = refundrequire
					newasid 		= -1
					modeflag2   	= "regcsas"
					divcd       	= "A008"
					id          	= 0
					ipkumdiv    	= myorder.FOneItem.FIpkumDiv
					reguserid   	= CFINISH_SYSTEM
					finishuser  	= CFINISH_SYSTEM
					title       	= "[자동전체취소]" & GetDefaultTitle(divcd, 0, orderserial)
					gubun01     	= "C004"  ''공통

					If IsSoldOutCancel Then
						gubun02     	= "CD05"  ''품절
					Else
						gubun02     	= "CD01"  ''단순변심
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
							'' CS Master 접수
							CsId = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "002"
							'' CS Detail 접수
							Call RegWebCSDetailAllCancel(CsId, orderserial)
						end if

						' 자동취소날짜를 넣는다.
						RegmibesongCanceldate(orderserial)

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "003"
							'' 환불 관련정보 (선)저장
							'// 언제나 등록한다. 2019-01-10, skyer9
							''if (refundrequire<>"0") and (returnmethod<>"R000") then
								refundcouponsum = refundcouponsum * -1
								refundmileagesum = refundmileagesum * -1
								refundgiftcardsum = refundgiftcardsum * -1
								refunddepositsum = refunddepositsum * -1

								'CS Master 환불 관련정보 저장	''# RegCSMasterRefundInfo, AddCSMasterRefundInfo -> /cscenter/lib/csAsfunction.asp
								Call RegCSMasterRefundInfo(CsId, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, refundallatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
								Call AddCSMasterRefundInfo(CsId, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

								'''계좌 암호화 추가.
								Call EditCSMasterRefundEncInfo(CsId, encmethod, rebankaccount)
							''end if
						End if

						' 상품쿠폰환급여부
						if itemCouponRefundYN="Y" and userid<>"GuestOrder" and userid<>"" then
							If (Err.Number = 0) and (ScanErr="") Then
								errcode = "004"

								' 주문 상품쿠폰 적용수 체크
								resultItemCouponCount = ItemCouponCount(CsId, "P", userid)
								if resultItemCouponCount>0 then
									copyitemcouponinfo="Y"
								else
									copyitemcouponinfo="N"
								end if

								' 상품쿠폰 재발행할지
								Call EditCSCopyItemCouponInfo(CsId, copyitemcouponinfo)
							end if
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							dbget.CommitTrans

                            finishArrOrderserial = finishArrOrderserial + "," + orderserial
							successCnt = successCnt + 1
							response.write "전체취소접수완료 : " & orderserial & "<br /><br />"

							'########################################### 선물포장 결제 취소. 전체취소만 됨. ###########################################
							If vIsPacked = "Y" Then
								sqlStr = "UPDATE [db_order].[dbo].[tbl_order_pack_master] SET cancelyn = 'Y' WHERE orderserial = '" & orderserial & "' " & vbCrLf
								sqlStr = sqlStr & "UPDATE [db_order].[dbo].[tbl_order_pack_detail] SET cancelyn = 'Y' "
								sqlStr = sqlStr & "WHERE midx IN(select midx from [db_order].[dbo].[tbl_order_pack_master] where orderserial = '" & orderserial & "')"
								dbget.Execute sqlStr
							End If
							'########################################### 선물포장 결제 취소. 전체취소만 됨. ###########################################

						Else
							dbget.RollBackTrans
							response.write "데이타를 저장하는 도중에 에러가 발생하였습니다. 관리자 문의 요망.(에러코드 : " & CStr(errcode) & ")" & "<br /><br />"
						End If
					else
						response.write orderserial & " : " & orderusingmsg & "<br /><br />"
					end if

				'### 품절/택배파업 취소(부분취소일때) 프로세스 ###
				elseif (cancelmode="stockoutcancel") then
					If vIsPacked = "Y" Then
						orderusingmsg="선물포장인 주문건 입니다.\n선물포장 주문은 전체취소만 가능합니다."
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
					'vQuery = vQuery & "			((d.itemid <> 0) and (IsNull(m.code, '') = '05')) "					'// 현재는 품절만 자동취소, skyer9, 2020-11-10
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
						orderusingmsg="품절/택배파업취소접수 할 수 없습니다.[코드번호:3-3]"
					end if

					'기타할인, 퍼센트쿠폰 당연차감
					refundrequire = refunditemcostsum - refundallatsubtractsum - refundcouponsum + refundbeasongpay

					'마일리지, 예치금, 기프트카드 제외
					'// 2018-02-22, skyer9, 마일리지 이미 빠져있음.
					remainsubtotalprice = remainsubtotalprice - 0 - remaindepositsum - remaingiftcardsum

					'마일리지
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

					'기프트카드
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

					'예치금
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
						orderusingmsg="품절/택배파업취소접수 할 수 없습니다.[코드번호:4-1]"
					end if
					if refundrequire < 0 then
						orderusingmsg="품절/택배파업취소접수 할 수 없습니다.[코드번호:4-2]"
					end if

					canceltotal = refundrequire

					newasid 		= -1

					modeflag2   	= "regcsas"
					divcd       	= "A008"
					id          	= 0
					ipkumdiv    	= myorder.FOneItem.FIpkumDiv
					reguserid   	= CFINISH_SYSTEM
					finishuser  	= CFINISH_SYSTEM
					title       	= "[자동부분취소]" & GetDefaultTitle(divcd, 0, orderserial)
					gubun01     	= "C004"  ''공통
					gubun02     	= "CD05"  ''품절
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
							'' CS Master 접수
							id = RegCSMaster(divcd, orderserial, reguserid, title, contents_jupsu, gubun01, gubun02)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "002"
							'' CS Detail 접수
							Call RegWebCSDetailStockoutCancel(id, orderserial)
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "003"
							'' 환불 관련정보 (선)저장

							'// 언제나 등록한다. 2019-01-10, skyer9
							''if (refundrequire<>"0") and (returnmethod<>"R000") then
								refundcouponsum = refundcouponsum * -1
								refundmileagesum = refundmileagesum * -1
								refundgiftcardsum = refundgiftcardsum * -1
								refunddepositsum = refunddepositsum * -1

								'CS Master 환불 관련정보 저장	''# RegCSMasterRefundInfo, AddCSMasterRefundInfo -> /cscenter/lib/csAsfunction.asp
								Call RegCSMasterRefundInfo(id, returnmethod, refundrequire , orgsubtotalprice, orgitemcostsum, orgbeasongpay , orgmileagesum, orgcouponsum, orgallatdiscountsum  , canceltotal, refunditemcostsum, refundmileagesum, refundcouponsum, refundallatsubtractsum*-1, refundbeasongpay, refunddeliverypay, refundadjustpay , rebankname, rebankaccount, rebankownername, paygateTid)
								Call AddCSMasterRefundInfo(id, orggiftcardsum, orgdepositsum, refundgiftcardsum, refunddepositsum)

								'''계좌 암호화 추가.
								Call EditCSMasterRefundEncInfo(id, encmethod, rebankaccount)
							''end if
						End if

						' 상품쿠폰환급여부
						if itemCouponRefundYN="Y" and userid<>"GuestOrder" and userid<>"" then
							If (Err.Number = 0) and (ScanErr="") Then
								errcode = "004"

								' 주문 상품쿠폰 적용수 체크
								resultItemCouponCount = ItemCouponCount(id, "P", userid)
								if resultItemCouponCount>0 then
									copyitemcouponinfo="Y"
								else
									copyitemcouponinfo="N"
								end if

								' 상품쿠폰 재발행할지
								Call EditCSCopyItemCouponInfo(id, copyitemcouponinfo)
							end if
						end if

						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "005"

							CancelValidResultMessage = GetPartialCancelRegValidResult(id, orderserial)

							if (CancelValidResultMessage <> "") then
								ScanErr = CancelValidResultMessage
								response.write orderserial & " : 품절/택배파업취소접수 할 수 없습니다.[코드번호:6]" & CancelValidResultMessage & "<br />"
							end if
						End If

						'출고완료 또는 취소된 상품이 있을 경우, 진행정지(주문취소 불가)
						'출고완료된 상품은 반품만 가능하다.
						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "006"

							''출고 완료 또는 취소된 내역이 있는지 확인
							if Not (IsCancelValidState(id, orderserial)) then
								dbget.RollBackTrans
								response.write orderserial & " : 품절/택배파업취소접수 할 수 없습니다.[코드번호:5]" & "<br />"
							end if
						end if

						'' 완료처리 바로 진행할지 검토
						'' 업체 확인중 상태가 있는경우 - > 접수로만 진행
						If (Err.Number = 0) and (ScanErr="") Then
							errcode = "007"
							contents_finish = ""
						End If

						ResultMsg = ResultMsg + "->. [주문 취소 CS] 접수\n\n"

						If (Err.Number = 0) and (ScanErr="") Then
							'출고완료 또는 취소된 상품이 있을 경우, 진행정지(주문취소 불가)
							'출고완료된 상품은 반품만 가능하다.
							''출고 완료 또는 취소된 내역이 있는지 확인
							if Not (IsCancelValidState(id, orderserial)) then
								errcode = "006"
								dbget.RollBackTrans
								response.write orderserial & " : 품절/택배파업취소접수 할 수 없습니다.[코드번호:5]" & "<br /><br />"
							else
								dbget.CommitTrans
								response.write "품절/택배파업취소접수완료 : " & orderserial & "<br /><br />"

                                finishArrOrderserial = finishArrOrderserial + "," + orderserial
								successCnt = successCnt + 1
							end if

						Else
							dbget.RollBackTrans

							response.write orderserial & " : 품절/택배파업취소접수 할 수 없습니다.[99-"&errcode&"]" & "<br /><br />"
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

		' 테스트
 		'sqlStr="update db_temp.dbo.tbl_mibeasong_list set code = '05', itemlackno=1 , itemno=1,isSendSMSdate='2020-12-08 09:30:33.250', isSendEmaildate='2020-12-08 09:30:33.250', isautocanceldate=NULL ,isSendSMS = 'Y',isSendEmail = 'Y' , sendCount=1 where idx in (664264,664265,664266,665965,667111)"
		'response.write sqlStr
		'dbget.Execute sqlStr

		if successCnt>0 then
			response.write "<br>상품품절/택배파업 " & successCnt & "건 자동취소 접수 등록." & "<br />"
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
    dim mmgubun ''메모구분
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
		''일반메모
		sqlStr = "insert into [db_cs].[dbo].[tbl_cs_memo]"
		sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,finishuser,contents_jupsu,finishyn,finishdate, phoneNumber)"
		sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + writeuser + "','" + html2db(contents_jupsu) + "','Y',getdate(), '" + CStr(phoneNumber) + "')"

		'response.write sqlStr
		dbget.Execute sqlStr
	else
		''처리요청메모
		sqlStr = "insert into [db_cs].[dbo].[tbl_cs_memo]"
		sqlStr = sqlStr + "(orderserial,divcd,userid,mmgubun,writeuser,contents_jupsu,finishyn)"
		sqlStr = sqlStr + " values('" + orderserial + "','" + divcd + "','" + userid + "','" + mmgubun + "','" + writeuser + "','" + html2db(contents_jupsu) + "','N')"

		'response.write sqlStr
		dbget.Execute sqlStr
	end if
end function
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
