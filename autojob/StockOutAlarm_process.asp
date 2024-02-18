<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 미출고상품 품절상품/출고지연 안내
' History : 이상구 생성
'           2020.10.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/email/maillib2.asp" -->
<!-- #include virtual="/lib/email/smslib.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cancelOrderLib.asp" -->
<!-- #include virtual="/lib/classes/ordercls/sp_myordercls.asp" -->
<%
'response.write "TT"
'response.end
dim webImgUrl : webImgUrl		= "http://webimage.10x10.co.kr"
dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

IF application("Svr_Info")<>"Dev" THEN
	if (Not CheckVaildIP(ref)) then
		response.write "허용된 IP가 아닙니다.[" & ref & "]"
		dbget.Close() : response.end
	end if
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
dim finishArrOrderserial, kakaomsgstr, btnJson, smstitlestr, smsmsgstr
Const CFINISH_SYSTEM = "system"
	successCnt = 0
	intloop=0

select Case mode
	'// 품절 알람(메일,SMS)
    Case "soalarm"
		' D+3초과 사유미입력분 출고지연으로 자동변경	' 2020.10.27 한용민
		'sqlStr="exec db_cs.dbo.usp_Ten_CS_michulgo_itemsoldout"

		'response.write sqlStr
		''dbget.Execute sqlStr		'// 자동입력 안함, skyer9, 2020-12-21

		'////////////////// 품절상품
        sqlStr = " select distinct top 100 l.orderserial "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		sqlStr = sqlStr + " where l.code = '05' "
		sqlStr = sqlStr + " 	and l.state <= '4' "						'// 출고지연 => 품절 등록시 상태값 변경안하고 있음, skyer9, 2020-12-21
		sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		sqlStr = sqlStr + " 	and l.sendCount = 0 "
		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌. cs박희연님 요청 발송 해달라고함.
		sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		sqlStr = sqlStr + " order by l.orderserial "

		''response.write sqlStr
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        if Not rsget.Eof then
            do until rsget.eof
            arrOrderserial = arrOrderserial & rsget("orderserial") & ","
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if Right(arrOrderserial,1)="," then arrOrderserial=Left(arrOrderserial,Len(arrOrderserial)-1)
        arrOrderserial = split(arrOrderserial,",")

        if UBound(arrOrderserial)>-1 then
            for i=0 to UBound(arrOrderserial)
				orderserial = arrOrderserial(i)

                sqlStr = " select l.idx as mibeasongidx, m.buyname, m.buyhp, m.buyemail, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, d.itemno, m.regdate, m.sitename "
		        sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock) "
		        sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		        sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
				sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
				sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		        sqlStr = sqlStr + " where l.orderserial = '" & orderserial & "' "
		        sqlStr = sqlStr + " 	and l.code = '05' "
		        sqlStr = sqlStr + " 	and l.state <= '4' "
		        sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		        sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		        sqlStr = sqlStr + " 	and l.sendCount = 0 "
				sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		        sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		        sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		        sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		        sqlStr = sqlStr + " 	and ( "
		        sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		        sqlStr = sqlStr + " 		or "
		        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		        sqlStr = sqlStr + " 	) "
				'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
				'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌. cs박희연님 요청 발송 해달라고함.
				sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		        sqlStr = sqlStr + " order by d.itemid, d.itemoption "

				''response.write sqlStr
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

				if Not rsget.Eof then
					itemCnt = 0
					itemName = ""
					orderdate = ""
                    sitename = ""
					mibeasongidxArr = ""
					buyhp = ""
					do until rsget.eof
						itemCnt = itemCnt + 1
						mibeasongidxArr = mibeasongidxArr & rsget("mibeasongidx") & ","
						if itemName = "" then
							buyhp = rsget("buyhp")
							itemName = db2html(rsget("itemname"))
							orderdate = Left(rsget("regdate"), 10)

                            sitename = rsget("sitename")
                            select case sitename
                                case "interpark"
                                    sitename = "인터파크"
                                    mastertelno = "1588-1555"
                                case "coupang"
                                    sitename = "쿠팡"
                                    mastertelno = "1577-7011"
                                case "11st1010"
                                    sitename = "11번가"
                                    mastertelno = "1599-0110"
                                case "auction1010"
                                    sitename = "옥션"
                                    mastertelno = "1588-0184"
                                case "gmarket1010"
                                    sitename = "지마켓"
                                    mastertelno = "1566-5701"
                                case "nvstorefarm"
                                    sitename = "네이버 스토어팜"
                                    mastertelno = "1588-3819"
                                case "lfmall"
                                    sitename = "lfmall"
                                    mastertelno = "1544-5114"
                                case "lotteon"
                                    sitename = "롯데온"
                                    mastertelno = "1899-7000"
                                case else
                                    sitename = ""
                                    mastertelno = ""
                            end select
						end if
						rsget.MoveNext
    				loop
					rsget.close

					if Right(mibeasongidxArr,1)="," then mibeasongidxArr=Left(mibeasongidxArr,Len(mibeasongidxArr)-1)

					if (itemCnt > 0) and (itemName <> "") and (buyhp <> "") then
                        '// 제휴몰 메일발송 스킵
                        if (sitename = "") then
						    Call sendmailStockOutAlarm(orderserial)
						    Call AddCsMemo(orderserial,"1","", "system","[MAIL] 품절안내 메일이 발송되었습니다.")
                        end if

						if (itemCnt > 1) then
							itemName = itemName & " 외 " & (itemCnt - 1) & "종"
						end if

                        if (sitename = "") then
						    ' smstext = ""
						    ' smstext = smstext + "[텐바이텐]죄송합니다. 고객님" + vbCrLf
						    ' smstext = smstext + "주문하신 상품의 재고를 확보하기 위해 노력하였으나" + vbCrLf
						    ' smstext = smstext + "안타깝게도 재고 부족으로 품절되어 안내드립니다." + vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "품절 상품은 익일 자동 취소 및 환불해드릴예정입니다." + vbCrLf
						    ' smstext = smstext + "(결제수단별 환불 소요일 상이)" + vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "주문상품: " & itemName & vbCrLf
						    ' smstext = smstext + "주문일자: " & orderdate & vbCrLf
						    ' smstext = smstext + "취소하기: http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx=" & orderserial & vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "추후에는 더욱 철저한 재고 관리로 고객님께 불편드리지 않도록 노력하겠습니다."
						    ' Call SendNormalLMSTimeFix(buyhp, "[텐바이텐] 주문하신 상품 품절 안내드립니다.", CNORMALCALLBAKC, smstext)
							smstitlestr = "[텐바이텐]주문하신 상품 품절 안내드립니다."
							smsmsgstr = "[10x10] 품절 안내" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "죄송합니다. 고객님" & vbCrLf
							smsmsgstr = smsmsgstr & "주문하신 상품의 재고를 확보하기 위해 노력하였으나" & vbCrLf
							smsmsgstr = smsmsgstr & "안타깝게도 재고 부족으로 품절되어 안내드립니다." & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "품절 상품은 익일 자동 취소 및 환불해드릴예정입니다." & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							smsmsgstr = smsmsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "추후에는 더욱 철저한 재고 관리로 고객님께 불편드리지 않도록 노력하겠습니다." & vbCrLf
							smsmsgstr = smsmsgstr & "감사합니다." & vbCrLf
							smsmsgstr = smsmsgstr & "취소하기 : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							btnJson = "{""button"":[{""name"":""취소하기"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							kakaomsgstr = "[10x10] 품절 안내" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "죄송합니다. 고객님" & vbCrLf
							kakaomsgstr = kakaomsgstr & "주문하신 상품의 재고를 확보하기 위해 노력하였으나" & vbCrLf
							kakaomsgstr = kakaomsgstr & "안타깝게도 재고 부족으로 품절되어 안내드립니다." & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "품절 상품은 익일 자동 취소 및 환불해드릴예정입니다." & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							kakaomsgstr = kakaomsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "추후에는 더욱 철저한 재고 관리로 고객님께 불편드리지 않도록 노력하겠습니다." & vbCrLf
							kakaomsgstr = kakaomsgstr & "감사합니다."
							Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0018",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")

						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [텐바이텐] 주문하신 상품 품절 안내드립니다.")
                        else
						    smstext = ""
						    smstext = smstext + "[텐바이텐] " & sitename & " 통해 주문하신 " & itemName & " 상품이 품절되어 문자안내드립니다." + vbCrLf
						    smstext = smstext + "번거로우시겠지만 구매진행하신 " & sitename & " 통해 취소접수 부탁드립니다." + vbCrLf
                            smstext = smstext + "쇼핑몰 이용시 불편드려 대단히 죄송합니다. [" & sitename & " : " & mastertelno & "]" + vbCrLf

						    Call SendNormalLMSTimeFix(buyhp, "[텐바이텐] 주문하신 상품 품절 안내드립니다.", mastertelno, smstext)
						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [텐바이텐] 주문하신 상품 품절 안내드립니다.")
                        end if

						sqlStr = " update db_temp.dbo.tbl_mibeasong_list "
						sqlStr = sqlStr + " set isSendSMS = 'Y', isSendEmail = 'Y', sendCount = sendCount + 1, state = 4 "
						sqlStr = sqlStr + " , isSendSMSdate=getdate(), isSendEmaildate=getdate()" & vbcrlf
						sqlStr = sqlStr + " where idx in (" & mibeasongidxArr & ") and isSendSMS <> 'Y' and state <= '4' "

						'response.write sqlStr
						dbget.Execute sqlStr

						successCnt = successCnt + 1
					end if
				else
					rsget.close
				end if
			next
		end if

		response.write "상품품절 " & successCnt & "건 전송완료.<br>"

		'////////////////// 택배파업	' 2022.01.17 한용민 생성
		successCnt=0
		arrOrderserial=""
		orderserial=""
        sqlStr = " select distinct top 100 l.orderserial "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		sqlStr = sqlStr + " where l.code = '06' "
		sqlStr = sqlStr + " 	and l.state <= '4' "						'// 출고지연 => 품절 등록시 상태값 변경안하고 있음, skyer9, 2020-12-21
		sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		sqlStr = sqlStr + " 	and l.sendCount = 0 "
		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌. cs박희연님 요청 발송 해달라고함.
		sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		sqlStr = sqlStr + " order by l.orderserial "

		''response.write sqlStr
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        if Not rsget.Eof then
            do until rsget.eof
            arrOrderserial = arrOrderserial & rsget("orderserial") & ","
            rsget.MoveNext
    		loop
        end if
        rsget.close

        if Right(arrOrderserial,1)="," then arrOrderserial=Left(arrOrderserial,Len(arrOrderserial)-1)
        arrOrderserial = split(arrOrderserial,",")

        if UBound(arrOrderserial)>-1 then
            for i=0 to UBound(arrOrderserial)
				orderserial = arrOrderserial(i)

                sqlStr = " select l.idx as mibeasongidx, m.buyname, m.buyhp, m.buyemail, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, d.itemno, m.regdate, m.sitename "
		        sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock) "
		        sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		        sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
				sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
				sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		        sqlStr = sqlStr + " where l.orderserial = '" & orderserial & "' "
		        sqlStr = sqlStr + " 	and l.code = '06' "
		        sqlStr = sqlStr + " 	and l.state <= '4' "
		        sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		        sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		        sqlStr = sqlStr + " 	and l.sendCount = 0 "
				sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		        sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		        sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		        sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		        sqlStr = sqlStr + " 	and ( "
		        sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		        sqlStr = sqlStr + " 		or "
		        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		        sqlStr = sqlStr + " 	) "
				'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
				'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌. cs박희연님 요청 발송 해달라고함.
				sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		        sqlStr = sqlStr + " order by d.itemid, d.itemoption "

				'response.write sqlStr & "<br>"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

				if Not rsget.Eof then
					itemCnt = 0
					itemName = ""
					orderdate = ""
                    sitename = ""
					mibeasongidxArr = ""
					buyhp = ""
					do until rsget.eof
						itemCnt = itemCnt + 1
						mibeasongidxArr = mibeasongidxArr & rsget("mibeasongidx") & ","
						if itemName = "" then
							buyhp = rsget("buyhp")
							itemName = db2html(rsget("itemname"))
							orderdate = Left(rsget("regdate"), 10)

                            sitename = rsget("sitename")
                            select case sitename
                                case "interpark"
                                    sitename = "인터파크"
                                    mastertelno = "1588-1555"
                                case "coupang"
                                    sitename = "쿠팡"
                                    mastertelno = "1577-7011"
                                case "11st1010"
                                    sitename = "11번가"
                                    mastertelno = "1599-0110"
                                case "auction1010"
                                    sitename = "옥션"
                                    mastertelno = "1588-0184"
                                case "gmarket1010"
                                    sitename = "지마켓"
                                    mastertelno = "1566-5701"
                                case "nvstorefarm"
                                    sitename = "네이버 스토어팜"
                                    mastertelno = "1588-3819"
                                case "lfmall"
                                    sitename = "lfmall"
                                    mastertelno = "1544-5114"
                                case "lotteon"
                                    sitename = "롯데온"
                                    mastertelno = "1899-7000"
                                case else
                                    sitename = ""
                                    mastertelno = ""
                            end select
						end if
						rsget.MoveNext
    				loop
					rsget.close

					if Right(mibeasongidxArr,1)="," then mibeasongidxArr=Left(mibeasongidxArr,Len(mibeasongidxArr)-1)

					if (itemCnt > 0) and (itemName <> "") and (buyhp <> "") then
                        '// 제휴몰 메일발송 스킵
                        if (sitename = "") then
						    Call sendmailDeliverystrikeAlarm(orderserial)
						    Call AddCsMemo(orderserial,"1","", "system","[MAIL] 택배파업안내 메일이 발송되었습니다.")
                        end if

						if (itemCnt > 1) then
							itemName = itemName & " 외 " & (itemCnt - 1) & "종"
						end if

                        if (sitename = "") then
						    smstitlestr = "[텐바이텐]주문하신 상품 택배파업 배송불가 안내"
							smsmsgstr = "[10x10] 택배파업 배송불가안내" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "죄송합니다. 고객님" & vbCrLf
							smsmsgstr = smsmsgstr & "택배파업으로 인해 고객님의 배송지로 택배발송이 어렵게되어 안내드립니다." & vbCrLf
							smsmsgstr = smsmsgstr & "현재 배송재개 가능 일정을 알수없는 상황으로" & vbCrLf
							smsmsgstr = smsmsgstr & "안타깝지만 주문취소 안내드리는 점 양해부탁드립니다." & vbCrLf
							smsmsgstr = smsmsgstr & "주문상품은 익일 자동 취소 및 환불예정입니다." & vbCrLf & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							smsmsgstr = smsmsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "감사합니다." & vbCrLf
							smsmsgstr = smsmsgstr & "취소하기 : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							btnJson = "{""button"":[{""name"":""취소하기"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							kakaomsgstr = "[10x10] 택배파업 배송불가안내" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "죄송합니다. 고객님" & vbCrLf
							kakaomsgstr = kakaomsgstr & "택배파업으로 인해 고객님의 배송지로 택배발송이 어렵게되어 안내드립니다." & vbCrLf
							kakaomsgstr = kakaomsgstr & "현재 배송재개 가능 일정을 알수없는 상황으로" & vbCrLf
							kakaomsgstr = kakaomsgstr & "안타깝지만 주문취소 안내드리는 점 양해부탁드립니다." & vbCrLf
							kakaomsgstr = kakaomsgstr & "주문상품은 익일 자동 취소 및 환불예정입니다." & vbCrLf & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "■ 주문번호 : "& orderserial &"" & vbCrLf
							kakaomsgstr = kakaomsgstr & "■ 상품명 : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "감사합니다."
							Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0025",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")

						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [텐바이텐]주문하신 상품 택배파업 배송불가 안내드립니다.")
                        else
							smstext = "[" & sitename & "]사이트 통해 주문하신 [" & itemName & "]상품이 택배파업으로 발송 어렵게되어 안내드립니다." & vbCrLf
							smstext = smstext & "번거로우시겠지만 구매진행하신 [" & sitename & "]사이트 통해 취소접수 부탁드립니다." & vbCrLf
							smstext = smstext & "쇼핑몰 이용시 불편드려 대단히 죄송합니다." & vbCrLf
							smstext = smstext & "[" & sitename & " : " & mastertelno & "]"
						    Call SendNormalLMSTimeFix(buyhp, "[10x10] 택배파업 배송불가안내", mastertelno, smstext)
						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [10x10] 택배파업 배송불가안내")
                        end if

						sqlStr = " update db_temp.dbo.tbl_mibeasong_list "
						sqlStr = sqlStr + " set isSendSMS = 'Y', isSendEmail = 'Y', sendCount = sendCount + 1, state = 4 "
						sqlStr = sqlStr + " , isSendSMSdate=getdate(), isSendEmaildate=getdate()" & vbcrlf
						sqlStr = sqlStr + " where idx in (" & mibeasongidxArr & ") and isSendSMS <> 'Y' and state <= '4' "

						'response.write sqlStr & "<br>"
						dbget.Execute sqlStr

						successCnt = successCnt + 1
					end if
				else
					rsget.close
				end if
			next
		end if

		response.write "택배파업 " & successCnt & "건 전송완료.<br>"

		' '////////////////// 출고지연	' 2020.10.27 한용민
		' successCnt=0
		' arrOrderserial=""
		' orderserial=""
        ' sqlStr = " select distinct top 100 l.orderserial "
		' sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
		' sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		' sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		' sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
		' sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		' sqlStr = sqlStr + " left join db_temp.dbo.tbl_mibeasong_list bl with (nolock)"
		' sqlStr = sqlStr + " 	on l.orderserial = bl.orderserial"
		' sqlStr = sqlStr + " 	and bl.code = '05'"
		' sqlStr = sqlStr + " left join [db_order].[dbo].[tbl_order_detail] dd with (nolock)"
		' sqlStr = sqlStr + " 		on bl.detailidx = dd.idx"
		' sqlStr = sqlStr + " 		and dd.isupchebeasong='N'"		' 텐배만체크
		' sqlStr = sqlStr + " where l.code = '03' "
		' sqlStr = sqlStr + " 	and l.state < '4' "
		' sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		' sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		' sqlStr = sqlStr + " 	and l.sendCount = 0 "
		' sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		' sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		' sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		' sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		' sqlStr = sqlStr + " 	and ( "
		' sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		' sqlStr = sqlStr + " 		or "
		' sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		' sqlStr = sqlStr + " 	) "
		' sqlStr = sqlStr + " and d.isupchebeasong='N'"	' 텐배만
		' sqlStr = sqlStr + " and dd.orderserial is null"		' 품절문자 발송이 있는경우, 출고지연 문자는 보내지 않는다.
		' 'sqlStr = sqlStr + " and m.sitename not in ('10x10_cs')"
		' sqlStr = sqlStr + " and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌.
		' sqlStr = sqlStr + " and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		' sqlStr = sqlStr + " order by l.orderserial "

		' ''response.write sqlStr
        ' rsget.CursorLocation = adUseClient
        ' rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

        ' if Not rsget.Eof then
        '     do until rsget.eof
        '     arrOrderserial = arrOrderserial & rsget("orderserial") & ","
        '     rsget.MoveNext
    	' 	loop
        ' end if
        ' rsget.close

        ' if Right(arrOrderserial,1)="," then arrOrderserial=Left(arrOrderserial,Len(arrOrderserial)-1)
        ' arrOrderserial = split(arrOrderserial,",")

        ' if UBound(arrOrderserial)>-1 then
        '     for i=0 to UBound(arrOrderserial)
		' 		orderserial = arrOrderserial(i)

        '         sqlStr = " select l.idx as mibeasongidx, m.buyname, m.buyhp, m.buyemail, d.itemid, d.itemoption, d.itemname, d.itemoptionname, d.itemcostCouponNotApplied, d.reducedPrice, d.itemno, m.regdate, m.sitename "
		'         sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock) "
		'         sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		'         sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		' 		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
		' 		sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		' 		sqlStr = sqlStr + " left join db_temp.dbo.tbl_mibeasong_list bl with (nolock)"
		' 		sqlStr = sqlStr + " 	on l.orderserial = bl.orderserial"
		' 		sqlStr = sqlStr + " 	and bl.code = '05'"
		' 		sqlStr = sqlStr + " left join [db_order].[dbo].[tbl_order_detail] dd with (nolock)"
		' 		sqlStr = sqlStr + " 		on bl.detailidx = dd.idx"
		' 		sqlStr = sqlStr + " 		and dd.isupchebeasong='N'"		' 텐배만체크
		'         sqlStr = sqlStr + " where l.orderserial = '" & orderserial & "' "
		'         sqlStr = sqlStr + " 	and l.code = '03' "
		'         sqlStr = sqlStr + " 	and l.state < '4' "
		'         sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		'         sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		'         sqlStr = sqlStr + " 	and l.sendCount = 0 "
		' 		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		'         sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		'         sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		'         sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// 출고완료 제외
		'         sqlStr = sqlStr + " 	and ( "
		'         sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		'         sqlStr = sqlStr + " 		or "
		'         sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// 인터파크,LF,지마켓,옥션,11번가,쿠팡,네이버 스토어팜은 우리가 먼저 SMS 발송후 고객이 동의해야 취소진행 가능
		'         sqlStr = sqlStr + " 	) "
		' 		sqlStr = sqlStr + " and d.isupchebeasong='N'"	' 텐배만
		' 		sqlStr = sqlStr + " and dd.orderserial is null"		' 품절문자 발송이 있는경우, 출고지연 문자는 보내지 않는다.
		' 		'sqlStr = sqlStr + " and m.sitename not in ('10x10_cs')"
		' 		sqlStr = sqlStr + " and isnull(i.reserveItemTp,'') not in ('1')"	' 단독구매상품,예약구매상품 제낌.
		' 		sqlStr = sqlStr + " and isnull(i.itemdiv,'')<>'75'"		' 정기구독상품 제외
		'         sqlStr = sqlStr + " order by d.itemid, d.itemoption "

		' 		''response.write sqlStr
		' 		rsget.CursorLocation = adUseClient
		' 		rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

		' 		if Not rsget.Eof then
		' 			itemCnt = 0
		' 			itemName = ""
		' 			orderdate = ""
        '             sitename = ""
		' 			mibeasongidxArr = ""
		' 			buyhp = ""
		' 			do until rsget.eof
		' 				itemCnt = itemCnt + 1
		' 				mibeasongidxArr = mibeasongidxArr & rsget("mibeasongidx") & ","
		' 				if itemName = "" then
		' 					buyhp = rsget("buyhp")
		' 					itemName = db2html(rsget("itemname"))
		' 					orderdate = Left(rsget("regdate"), 10)

        '                     sitename = rsget("sitename")
        '                     select case sitename
        '                         case "interpark"
        '                             sitename = "인터파크"
        '                             mastertelno = "1588-1555"
        '                         case "coupang"
        '                             sitename = "쿠팡"
        '                             mastertelno = "1577-7011"
        '                         case "11st1010"
        '                             sitename = "11번가"
        '                             mastertelno = "1599-0110"
        '                         case "auction1010"
        '                             sitename = "옥션"
        '                             mastertelno = "1588-0184"
        '                         case "gmarket1010"
        '                             sitename = "지마켓"
        '                             mastertelno = "1566-5701"
        '                         case "nvstorefarm"
        '                             sitename = "네이버 스토어팜"
        '                             mastertelno = "1588-3819"
        '                         case "lfmall"
        '                             sitename = "lfmall"
        '                             mastertelno = "1544-5114"
        '                        case "lotteon"
        '                            sitename = "롯데온"
        '                            mastertelno = "1899-7000"
        '                         case else
        '                             sitename = ""
        '                             mastertelno = ""
        '                     end select
		' 				end if
		' 				rsget.MoveNext
    	' 			loop
		' 			rsget.close

		' 			if Right(mibeasongidxArr,1)="," then mibeasongidxArr=Left(mibeasongidxArr,Len(mibeasongidxArr)-1)

		' 			if (itemCnt > 0) and (itemName <> "") and (buyhp <> "") then
        '                 '// 제휴몰 메일발송 스킵
        '                 if (sitename = "") then
		' 				    Call sendmaildelayAlarm(orderserial)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[MAIL] 발송지연 안내 메일이 발송되었습니다.")
        '                 end if

		' 				if (itemCnt > 1) then
		' 					itemName = itemName & " 외 " & (itemCnt - 1) & "종"
		' 				end if

        '                 if (sitename = "") then
		' 				    smstext = ""
		' 				    smstext = smstext + "[텐바이텐]죄송합니다. 고객님" + vbCrLf
		' 				    smstext = smstext + "주문하신 상품의 재고를 확보중에 있으나 예상보다 지연되어 문자드립니다." + vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "주문상품: " & itemName & vbCrLf
		' 				    smstext = smstext + "주문일자: " & orderdate & vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "빠른 시일내에 발송할 수 있도록 노력 중이나" + vbCrLf
		' 				    smstext = smstext + "혹 배송이 늦어져 수령을 원치 않으시는 고객님 께서는 주문취소 접수 부탁드립니다." + vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "취소하기: http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx=" & orderserial & vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "다시 한번 이용에 불편 드린 점 진심으로 사과 말씀 드리며," + vbCrLf
		' 				    smstext = smstext + "앞으로 더 나은 서비스로 보답하고자 노력하겠습니다." + vbCrLf
		' 				    smstext = smstext + "감사합니다."

		' 				    Call SendNormalLMSTimeFix(buyhp, "[텐바이텐] 주문하신 상품 발송지연 안내드립니다.", CNORMALCALLBAKC, smstext)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [텐바이텐] 주문하신 상품 발송지연 안내드립니다.")
        '                 else
		' 				    smstext = ""
		' 				    smstext = smstext + "[텐바이텐] " & sitename & " 통해 주문하신 " & itemName & " 상품이 발송지연 되어 문자안내드립니다." + vbCrLf
		' 				    smstext = smstext + "번거로우시겠지만 구매진행하신 " & sitename & " 통해 취소접수 부탁드립니다." + vbCrLf
        '                     smstext = smstext + "쇼핑몰 이용시 불편드려 대단히 죄송합니다. [" & sitename & " : " & mastertelno & "]" + vbCrLf

		' 				    Call SendNormalLMSTimeFix(buyhp, "[텐바이텐] 주문하신 상품 발송지연 안내드립니다.", mastertelno, smstext)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [텐바이텐] 주문하신 상품 발송지연 안내드립니다.")
        '                 end if

		' 				sqlStr = " update db_temp.dbo.tbl_mibeasong_list "
		' 				sqlStr = sqlStr + " set isSendSMS = 'Y', isSendEmail = 'Y', sendCount = sendCount + 1, state = 4 "
		' 				sqlStr = sqlStr + " , isSendSMSdate=getdate(), isSendEmaildate=getdate()" & vbcrlf
		' 				sqlStr = sqlStr + " where idx in (" & mibeasongidxArr & ") and isSendSMS <> 'Y' and state < '4' "

		' 				'response.write sqlStr
		' 				dbget.Execute sqlStr

		' 				successCnt = successCnt + 1
		' 			end if
		' 		else
		' 			rsget.close
		' 			continue
		' 		end if
		' 	next
		' end if

		' response.write "<br>출고지연 " & successCnt & "건 전송완료.<br>"

	case else
		dbget.Close()
		response.end
end select

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67","192.168.1.67","192.168.1.73", "121.78.103.60")
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
