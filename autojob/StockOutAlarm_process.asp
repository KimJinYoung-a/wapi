<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ǰ ǰ����ǰ/������� �ȳ�
' History : �̻� ����
'           2020.10.27 �ѿ�� ����
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
		response.write "���� IP�� �ƴմϴ�.[" & ref & "]"
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
	'// ǰ�� �˶�(����,SMS)
    Case "soalarm"
		' D+3�ʰ� �������Էº� ����������� �ڵ�����	' 2020.10.27 �ѿ��
		'sqlStr="exec db_cs.dbo.usp_Ten_CS_michulgo_itemsoldout"

		'response.write sqlStr
		''dbget.Execute sqlStr		'// �ڵ��Է� ����, skyer9, 2020-12-21

		'////////////////// ǰ����ǰ
        sqlStr = " select distinct top 100 l.orderserial "
		sqlStr = sqlStr + " from db_temp.dbo.tbl_mibeasong_list l with (nolock)"
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_detail] d with (nolock) on d.idx = l.detailidx "
		sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_order_master] m with (nolock) on d.orderserial = m.orderserial "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item i with (nolock)"
		sqlStr = sqlStr + " 		on d.itemid = i.itemid"
		sqlStr = sqlStr + " where l.code = '05' "
		sqlStr = sqlStr + " 	and l.state <= '4' "						'// ������� => ǰ�� ��Ͻ� ���°� ������ϰ� ����, skyer9, 2020-12-21
		sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		sqlStr = sqlStr + " 	and l.sendCount = 0 "
		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����. cs���񿬴� ��û �߼� �ش޶����.
		sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
		        sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		        sqlStr = sqlStr + " 	and ( "
		        sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		        sqlStr = sqlStr + " 		or "
		        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		        sqlStr = sqlStr + " 	) "
				'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
				'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����. cs���񿬴� ��û �߼� �ش޶����.
				sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
                                    sitename = "������ũ"
                                    mastertelno = "1588-1555"
                                case "coupang"
                                    sitename = "����"
                                    mastertelno = "1577-7011"
                                case "11st1010"
                                    sitename = "11����"
                                    mastertelno = "1599-0110"
                                case "auction1010"
                                    sitename = "����"
                                    mastertelno = "1588-0184"
                                case "gmarket1010"
                                    sitename = "������"
                                    mastertelno = "1566-5701"
                                case "nvstorefarm"
                                    sitename = "���̹� �������"
                                    mastertelno = "1588-3819"
                                case "lfmall"
                                    sitename = "lfmall"
                                    mastertelno = "1544-5114"
                                case "lotteon"
                                    sitename = "�Ե���"
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
                        '// ���޸� ���Ϲ߼� ��ŵ
                        if (sitename = "") then
						    Call sendmailStockOutAlarm(orderserial)
						    Call AddCsMemo(orderserial,"1","", "system","[MAIL] ǰ���ȳ� ������ �߼۵Ǿ����ϴ�.")
                        end if

						if (itemCnt > 1) then
							itemName = itemName & " �� " & (itemCnt - 1) & "��"
						end if

                        if (sitename = "") then
						    ' smstext = ""
						    ' smstext = smstext + "[�ٹ�����]�˼��մϴ�. ����" + vbCrLf
						    ' smstext = smstext + "�ֹ��Ͻ� ��ǰ�� ��� Ȯ���ϱ� ���� ����Ͽ�����" + vbCrLf
						    ' smstext = smstext + "��Ÿ���Ե� ��� �������� ǰ���Ǿ� �ȳ��帳�ϴ�." + vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "ǰ�� ��ǰ�� ���� �ڵ� ��� �� ȯ���ص帱�����Դϴ�." + vbCrLf
						    ' smstext = smstext + "(�������ܺ� ȯ�� �ҿ��� ����)" + vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "�ֹ���ǰ: " & itemName & vbCrLf
						    ' smstext = smstext + "�ֹ�����: " & orderdate & vbCrLf
						    ' smstext = smstext + "����ϱ�: http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx=" & orderserial & vbCrLf
						    ' smstext = smstext + vbCrLf
						    ' smstext = smstext + "���Ŀ��� ���� ö���� ��� ������ ���Բ� ����帮�� �ʵ��� ����ϰڽ��ϴ�."
						    ' Call SendNormalLMSTimeFix(buyhp, "[�ٹ�����] �ֹ��Ͻ� ��ǰ ǰ�� �ȳ��帳�ϴ�.", CNORMALCALLBAKC, smstext)
							smstitlestr = "[�ٹ�����]�ֹ��Ͻ� ��ǰ ǰ�� �ȳ��帳�ϴ�."
							smsmsgstr = "[10x10] ǰ�� �ȳ�" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "�˼��մϴ�. ����" & vbCrLf
							smsmsgstr = smsmsgstr & "�ֹ��Ͻ� ��ǰ�� ��� Ȯ���ϱ� ���� ����Ͽ�����" & vbCrLf
							smsmsgstr = smsmsgstr & "��Ÿ���Ե� ��� �������� ǰ���Ǿ� �ȳ��帳�ϴ�." & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "ǰ�� ��ǰ�� ���� �ڵ� ��� �� ȯ���ص帱�����Դϴ�." & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							smsmsgstr = smsmsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "���Ŀ��� ���� ö���� ��� ������ ���Բ� ����帮�� �ʵ��� ����ϰڽ��ϴ�." & vbCrLf
							smsmsgstr = smsmsgstr & "�����մϴ�." & vbCrLf
							smsmsgstr = smsmsgstr & "����ϱ� : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							btnJson = "{""button"":[{""name"":""����ϱ�"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							kakaomsgstr = "[10x10] ǰ�� �ȳ�" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "�˼��մϴ�. ����" & vbCrLf
							kakaomsgstr = kakaomsgstr & "�ֹ��Ͻ� ��ǰ�� ��� Ȯ���ϱ� ���� ����Ͽ�����" & vbCrLf
							kakaomsgstr = kakaomsgstr & "��Ÿ���Ե� ��� �������� ǰ���Ǿ� �ȳ��帳�ϴ�." & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "ǰ�� ��ǰ�� ���� �ڵ� ��� �� ȯ���ص帱�����Դϴ�." & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							kakaomsgstr = kakaomsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "���Ŀ��� ���� ö���� ��� ������ ���Բ� ����帮�� �ʵ��� ����ϰڽ��ϴ�." & vbCrLf
							kakaomsgstr = kakaomsgstr & "�����մϴ�."
							Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0018",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")

						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [�ٹ�����] �ֹ��Ͻ� ��ǰ ǰ�� �ȳ��帳�ϴ�.")
                        else
						    smstext = ""
						    smstext = smstext + "[�ٹ�����] " & sitename & " ���� �ֹ��Ͻ� " & itemName & " ��ǰ�� ǰ���Ǿ� ���ھȳ��帳�ϴ�." + vbCrLf
						    smstext = smstext + "���ŷο�ð����� ���������Ͻ� " & sitename & " ���� ������� ��Ź�帳�ϴ�." + vbCrLf
                            smstext = smstext + "���θ� �̿�� ������ ����� �˼��մϴ�. [" & sitename & " : " & mastertelno & "]" + vbCrLf

						    Call SendNormalLMSTimeFix(buyhp, "[�ٹ�����] �ֹ��Ͻ� ��ǰ ǰ�� �ȳ��帳�ϴ�.", mastertelno, smstext)
						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [�ٹ�����] �ֹ��Ͻ� ��ǰ ǰ�� �ȳ��帳�ϴ�.")
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

		response.write "��ǰǰ�� " & successCnt & "�� ���ۿϷ�.<br>"

		'////////////////// �ù��ľ�	' 2022.01.17 �ѿ�� ����
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
		sqlStr = sqlStr + " 	and l.state <= '4' "						'// ������� => ǰ�� ��Ͻ� ���°� ������ϰ� ����, skyer9, 2020-12-21
		sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		sqlStr = sqlStr + " 	and l.sendCount = 0 "
		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		sqlStr = sqlStr + " 	and ( "
		sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		sqlStr = sqlStr + " 		or "
		sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		sqlStr = sqlStr + " 	) "
		'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
		'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����. cs���񿬴� ��û �߼� �ش޶����.
		sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
		        sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		        sqlStr = sqlStr + " 	and ( "
		        sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		        sqlStr = sqlStr + " 		or "
		        sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		        sqlStr = sqlStr + " 	) "
				'sqlStr = sqlStr + " 	and m.sitename not in ('10x10_cs')"
				'sqlStr = sqlStr + " 	and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����. cs���񿬴� ��û �߼� �ش޶����.
				sqlStr = sqlStr + " 	and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
                                    sitename = "������ũ"
                                    mastertelno = "1588-1555"
                                case "coupang"
                                    sitename = "����"
                                    mastertelno = "1577-7011"
                                case "11st1010"
                                    sitename = "11����"
                                    mastertelno = "1599-0110"
                                case "auction1010"
                                    sitename = "����"
                                    mastertelno = "1588-0184"
                                case "gmarket1010"
                                    sitename = "������"
                                    mastertelno = "1566-5701"
                                case "nvstorefarm"
                                    sitename = "���̹� �������"
                                    mastertelno = "1588-3819"
                                case "lfmall"
                                    sitename = "lfmall"
                                    mastertelno = "1544-5114"
                                case "lotteon"
                                    sitename = "�Ե���"
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
                        '// ���޸� ���Ϲ߼� ��ŵ
                        if (sitename = "") then
						    Call sendmailDeliverystrikeAlarm(orderserial)
						    Call AddCsMemo(orderserial,"1","", "system","[MAIL] �ù��ľ��ȳ� ������ �߼۵Ǿ����ϴ�.")
                        end if

						if (itemCnt > 1) then
							itemName = itemName & " �� " & (itemCnt - 1) & "��"
						end if

                        if (sitename = "") then
						    smstitlestr = "[�ٹ�����]�ֹ��Ͻ� ��ǰ �ù��ľ� ��ۺҰ� �ȳ�"
							smsmsgstr = "[10x10] �ù��ľ� ��ۺҰ��ȳ�" & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "�˼��մϴ�. ����" & vbCrLf
							smsmsgstr = smsmsgstr & "�ù��ľ����� ���� ������ ������� �ù�߼��� ��ưԵǾ� �ȳ��帳�ϴ�." & vbCrLf
							smsmsgstr = smsmsgstr & "���� ����簳 ���� ������ �˼����� ��Ȳ����" & vbCrLf
							smsmsgstr = smsmsgstr & "��Ÿ������ �ֹ���� �ȳ��帮�� �� ���غ�Ź�帳�ϴ�." & vbCrLf
							smsmsgstr = smsmsgstr & "�ֹ���ǰ�� ���� �ڵ� ��� �� ȯ�ҿ����Դϴ�." & vbCrLf & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							smsmsgstr = smsmsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							smsmsgstr = smsmsgstr & "�����մϴ�." & vbCrLf
							smsmsgstr = smsmsgstr & "����ϱ� : http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &""
							btnJson = "{""button"":[{""name"":""����ϱ�"",""type"":""WL"", ""url_mobile"":""http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx="& orderserial &"""}]}"
							kakaomsgstr = "[10x10] �ù��ľ� ��ۺҰ��ȳ�" & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "�˼��մϴ�. ����" & vbCrLf
							kakaomsgstr = kakaomsgstr & "�ù��ľ����� ���� ������ ������� �ù�߼��� ��ưԵǾ� �ȳ��帳�ϴ�." & vbCrLf
							kakaomsgstr = kakaomsgstr & "���� ����簳 ���� ������ �˼����� ��Ȳ����" & vbCrLf
							kakaomsgstr = kakaomsgstr & "��Ÿ������ �ֹ���� �ȳ��帮�� �� ���غ�Ź�帳�ϴ�." & vbCrLf
							kakaomsgstr = kakaomsgstr & "�ֹ���ǰ�� ���� �ڵ� ��� �� ȯ�ҿ����Դϴ�." & vbCrLf & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "�� �ֹ���ȣ : "& orderserial &"" & vbCrLf
							kakaomsgstr = kakaomsgstr & "�� ��ǰ�� : "& itemName &"" & vbCrLf & vbCrLf & vbCrLf
							kakaomsgstr = kakaomsgstr & "�����մϴ�."
							Call SendKakaoCSMsg_LINK("",buyhp,"1644-6030","KC-0025",kakaomsgstr,"LMS", smstitlestr, smsmsgstr,btnJson,"","")

						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [�ٹ�����]�ֹ��Ͻ� ��ǰ �ù��ľ� ��ۺҰ� �ȳ��帳�ϴ�.")
                        else
							smstext = "[" & sitename & "]����Ʈ ���� �ֹ��Ͻ� [" & itemName & "]��ǰ�� �ù��ľ����� �߼� ��ưԵǾ� �ȳ��帳�ϴ�." & vbCrLf
							smstext = smstext & "���ŷο�ð����� ���������Ͻ� [" & sitename & "]����Ʈ ���� ������� ��Ź�帳�ϴ�." & vbCrLf
							smstext = smstext & "���θ� �̿�� ������ ����� �˼��մϴ�." & vbCrLf
							smstext = smstext & "[" & sitename & " : " & mastertelno & "]"
						    Call SendNormalLMSTimeFix(buyhp, "[10x10] �ù��ľ� ��ۺҰ��ȳ�", mastertelno, smstext)
						    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [10x10] �ù��ľ� ��ۺҰ��ȳ�")
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

		response.write "�ù��ľ� " & successCnt & "�� ���ۿϷ�.<br>"

		' '////////////////// �������	' 2020.10.27 �ѿ��
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
		' sqlStr = sqlStr + " 		and dd.isupchebeasong='N'"		' �ٹ踸üũ
		' sqlStr = sqlStr + " where l.code = '03' "
		' sqlStr = sqlStr + " 	and l.state < '4' "
		' sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		' sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		' sqlStr = sqlStr + " 	and l.sendCount = 0 "
		' sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		' sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		' sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		' sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		' sqlStr = sqlStr + " 	and ( "
		' sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		' sqlStr = sqlStr + " 		or "
		' sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		' sqlStr = sqlStr + " 	) "
		' sqlStr = sqlStr + " and d.isupchebeasong='N'"	' �ٹ踸
		' sqlStr = sqlStr + " and dd.orderserial is null"		' ǰ������ �߼��� �ִ°��, ������� ���ڴ� ������ �ʴ´�.
		' 'sqlStr = sqlStr + " and m.sitename not in ('10x10_cs')"
		' sqlStr = sqlStr + " and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����.
		' sqlStr = sqlStr + " and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
		' 		sqlStr = sqlStr + " 		and dd.isupchebeasong='N'"		' �ٹ踸üũ
		'         sqlStr = sqlStr + " where l.orderserial = '" & orderserial & "' "
		'         sqlStr = sqlStr + " 	and l.code = '03' "
		'         sqlStr = sqlStr + " 	and l.state < '4' "
		'         sqlStr = sqlStr + " 	and l.isSendSMS = 'N' "
		'         sqlStr = sqlStr + " 	and l.isSendEmail = 'N' "
		'         sqlStr = sqlStr + " 	and l.sendCount = 0 "
		' 		sqlStr = sqlStr + " 	and l.itemno>0 and l.itemlackno>0 "
		'         sqlStr = sqlStr + " 	and d.cancelyn <> 'Y' "
		'         sqlStr = sqlStr + " 	and m.cancelyn = 'N' "
		'         sqlStr = sqlStr + " 	and IsNull(d.currstate, '') <> '7' "		'// ���Ϸ� ����
		'         sqlStr = sqlStr + " 	and ( "
		'         sqlStr = sqlStr + " 		(m.jumundiv <> '5') "
		'         sqlStr = sqlStr + " 		or "
		'         sqlStr = sqlStr + " 		(m.jumundiv = '5' and m.sitename in ('interpark', 'coupang', '11st1010', 'auction1010', 'gmarket1010', 'nvstorefarm', 'lfmall','lotteon','10x10_cs')) "		'// ������ũ,LF,������,����,11����,����,���̹� ��������� �츮�� ���� SMS �߼��� ���� �����ؾ� ������� ����
		'         sqlStr = sqlStr + " 	) "
		' 		sqlStr = sqlStr + " and d.isupchebeasong='N'"	' �ٹ踸
		' 		sqlStr = sqlStr + " and dd.orderserial is null"		' ǰ������ �߼��� �ִ°��, ������� ���ڴ� ������ �ʴ´�.
		' 		'sqlStr = sqlStr + " and m.sitename not in ('10x10_cs')"
		' 		sqlStr = sqlStr + " and isnull(i.reserveItemTp,'') not in ('1')"	' �ܵ����Ż�ǰ,���౸�Ż�ǰ ����.
		' 		sqlStr = sqlStr + " and isnull(i.itemdiv,'')<>'75'"		' ���ⱸ����ǰ ����
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
        '                             sitename = "������ũ"
        '                             mastertelno = "1588-1555"
        '                         case "coupang"
        '                             sitename = "����"
        '                             mastertelno = "1577-7011"
        '                         case "11st1010"
        '                             sitename = "11����"
        '                             mastertelno = "1599-0110"
        '                         case "auction1010"
        '                             sitename = "����"
        '                             mastertelno = "1588-0184"
        '                         case "gmarket1010"
        '                             sitename = "������"
        '                             mastertelno = "1566-5701"
        '                         case "nvstorefarm"
        '                             sitename = "���̹� �������"
        '                             mastertelno = "1588-3819"
        '                         case "lfmall"
        '                             sitename = "lfmall"
        '                             mastertelno = "1544-5114"
        '                        case "lotteon"
        '                            sitename = "�Ե���"
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
        '                 '// ���޸� ���Ϲ߼� ��ŵ
        '                 if (sitename = "") then
		' 				    Call sendmaildelayAlarm(orderserial)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[MAIL] �߼����� �ȳ� ������ �߼۵Ǿ����ϴ�.")
        '                 end if

		' 				if (itemCnt > 1) then
		' 					itemName = itemName & " �� " & (itemCnt - 1) & "��"
		' 				end if

        '                 if (sitename = "") then
		' 				    smstext = ""
		' 				    smstext = smstext + "[�ٹ�����]�˼��մϴ�. ����" + vbCrLf
		' 				    smstext = smstext + "�ֹ��Ͻ� ��ǰ�� ��� Ȯ���߿� ������ ���󺸴� �����Ǿ� ���ڵ帳�ϴ�." + vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "�ֹ���ǰ: " & itemName & vbCrLf
		' 				    smstext = smstext + "�ֹ�����: " & orderdate & vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "���� ���ϳ��� �߼��� �� �ֵ��� ��� ���̳�" + vbCrLf
		' 				    smstext = smstext + "Ȥ ����� �ʾ��� ������ ��ġ �����ô� ���� ������ �ֹ���� ���� ��Ź�帳�ϴ�." + vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "����ϱ�: http://m.10x10.co.kr/my10x10/order/order_cancel_detail.asp?mode=so&idx=" & orderserial & vbCrLf
		' 				    smstext = smstext + vbCrLf
		' 				    smstext = smstext + "�ٽ� �ѹ� �̿뿡 ���� �帰 �� �������� ��� ���� �帮��," + vbCrLf
		' 				    smstext = smstext + "������ �� ���� ���񽺷� �����ϰ��� ����ϰڽ��ϴ�." + vbCrLf
		' 				    smstext = smstext + "�����մϴ�."

		' 				    Call SendNormalLMSTimeFix(buyhp, "[�ٹ�����] �ֹ��Ͻ� ��ǰ �߼����� �ȳ��帳�ϴ�.", CNORMALCALLBAKC, smstext)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [�ٹ�����] �ֹ��Ͻ� ��ǰ �߼����� �ȳ��帳�ϴ�.")
        '                 else
		' 				    smstext = ""
		' 				    smstext = smstext + "[�ٹ�����] " & sitename & " ���� �ֹ��Ͻ� " & itemName & " ��ǰ�� �߼����� �Ǿ� ���ھȳ��帳�ϴ�." + vbCrLf
		' 				    smstext = smstext + "���ŷο�ð����� ���������Ͻ� " & sitename & " ���� ������� ��Ź�帳�ϴ�." + vbCrLf
        '                     smstext = smstext + "���θ� �̿�� ������ ����� �˼��մϴ�. [" & sitename & " : " & mastertelno & "]" + vbCrLf

		' 				    Call SendNormalLMSTimeFix(buyhp, "[�ٹ�����] �ֹ��Ͻ� ��ǰ �߼����� �ȳ��帳�ϴ�.", mastertelno, smstext)
		' 				    Call AddCsMemo(orderserial,"1","", "system","[SMS "+ buyhp + "] [�ٹ�����] �ֹ��Ͻ� ��ǰ �߼����� �ȳ��帳�ϴ�.")
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

		' response.write "<br>������� " & successCnt & "�� ���ۿϷ�.<br>"

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
