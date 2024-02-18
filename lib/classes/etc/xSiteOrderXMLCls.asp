<%

'' !!!! 아래 파일이 인클루드 되어 있어야 한다.
''/admin/etc/lotte/inc_dailyAuthCheck.asp
''/lib/classes/etc/lotteitemcls.asp
''/admin/etc/incOutMallCommonFunction.asp

Class CxSiteOrderXML
    public FItemList()
	public FOneItem
	public FResultCount
	public FTotalCount

	public FRectSellSite
	public FRectYYYYMMDD
	public FRectStartYYYYMMDD
	public FRectEndYYYYMMDD

	public FRectGubun

	public FRectAPIURL
	public FRectAuthNo

	public ErrMsg
	private objXML
	private xmlDOM

	private xmlURL
	private objData

	public function GetCheckStatus(byRef LastCheckDate, byRef isSuccess)
		dim strSql

		''db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] 에 해당 사이트 데이타 없으면 넣어주어야 한다.
		''insert into db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp](sellsite, lastcheckdate, issuccess)
		''values('gseshop', '2014-08-20', 'N')

        ''2013/11/20 어제 데이터는 한번더 가져 오는걸로
        strSql = " IF Exists("
        strSql = strSql + " 	select LastcheckDate"
        strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
        strSql = strSql + " 	where dateDiff(d,LastUpdate,getdate())>0"
        strSql = strSql + " 	and sellsite='" + CStr(FRectSellSite) + "'"
        strSql = strSql + " )"
        strSql = strSql + " BEGIN"
        strSql = strSql + " 	Update T"
        strSql = strSql + " 	set LastcheckDate=dateadd(d,-1, LastcheckDate)"
        strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] T"
        strSql = strSql + " 	where sellsite='" + CStr(FRectSellSite) + "'"
        strSql = strSql + " END"
        dbget.Execute strSql

		strSql = " select LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "

		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			LastCheckDate = rsget("LastCheckDate")
			isSuccess = rsget("isSuccess")
		rsget.Close

	end function

	public function SetCheckStatusStarting(LastCheckDate)
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set LastCheckDate = '" + CStr(LastCheckDate) + "', isSuccess = 'N' "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	dbget.execute strSql

	end function

	public function SetCheckStatusEnded()
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set isSuccess = 'Y' "
		strSql = strSql + " ,LastUpdate=getdate()"
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	dbget.execute strSql

	end function

	public function SetCheckDate(LastCheckDate)
		dim strSql

		strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
		strSql = strSql + " set LastCheckDate = '" + CStr(LastCheckDate) + "' "
		strSql = strSql + " where sellsite = '" + CStr(FRectSellSite) + "' "
    	dbget.execute strSql

	end function

	public function SavexSiteOrderListtoDB()
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			rw xmlURL
'response.end
			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "등록되지 않은 제휴몰입니다.[0]"
			end if
		end if
'dbget.close() : response.end
		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "통신중에 오류가 발생했습니다."
			end if
		end if
'dbget.close() : response.end
		if (ErrMsg = "") then
			Call ActSavexSiteOrderListtoDB()
		end if

    end function

	public function RequestxSiteOrderListOnly()
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			rw xmlURL
'response.end
			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "등록되지 않은 제휴몰입니다.[0]"
			end if
		end if
'dbget.close() : response.end
		if (ErrMsg = "") then
			Call RequestXmlFromWeb()
		end if

    end function

	function ActSavexSiteOrderListtoDB()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim SellSite, OutMallOrderSerial, SellDate, PayType, Paydate, partnerItemID, partnerItemName, partnerOption, partnerOptionName, OrderUserID, OrderName, OrderEmail, OrderTelNo, OrderHpNo
		dim ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey, deliverymemo, requireDetail, orderDlvPay, etc1, countryCode
		dim matchItemID,matchItemOption, orderCsGbn, outMallGoodsNo
		dim strSql
		dim errCode, errStr

		dim tmpStr, isCSOrder
        dim retVal, succCNT,failCNT

        ''인터파크용
        Dim ORD_NO,ORDER_DT,ORD_NM,CLM_NO,CLM_SEQ
        Dim DEL_AMT
        Dim ENTR_DC_COUPON_AMT,DC_COUPON_AMT ,ENTR_PRD_NO ,SALE_UNITCOST ,ORD_AMT  ,SMONEY_EXPN_AMT ,SALE_FEE
        Dim ORD_QTY ,OPT_PRD_NO
        Dim COMP_AMT,EXPN_AMT ,PRE_USE_UNITCOST ,ORD_SEQ,IPOINT_PAY_AMT ,SUPPLY_CTRT_SEQ,IPOINT_SAVE_AMT
        Dim OLD_SALE_UNITCOST,SHOP_PAY_AMT ,PRD_NO ,OPT_NO ,ON_INTEREST_FEE,SETL_DT,PRE_USE_AMT
        Dim OPT_NM,SEL_OPT_NM,IN_OPT_NM

'response.write objdata
'response.end

        succCNT=0
        failCNT=0
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		if (FRectSellSite <> "gseshop") then
			xmlDOM.LoadXML replace(objData,"&","＆")
		else
			xmlDOM.LoadXML(Request)
		end if

		if (FRectSellSite = "lotteimall") then
			'// lotteimall
           '' rw "<textarea cols=20 rwos=5>"&objData&"</textarea>" ''' for XML data parsing

			'// 카운트
			masterCnt = (xmlDOM.selectNodes("/Response/Result/OrderCount").item(0).text * 1)

			response.write "(" & masterCnt & "건)"

			if (masterCnt > 0) then
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
				masterCnt = objMasterListXML.length

				for i = 0 to masterCnt - 1
					set objMasterOneXML = objMasterListXML.item(i)

					SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					SellDate			= objMasterOneXML.selectSingleNode("TrdDate").text
					PayType				= "50"
					Paydate				= objMasterOneXML.selectSingleNode("TrdDate").text
					OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("OrderName").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("OrderTelNo").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("OrderHpNo").text

					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/recvName").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text
					ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text
					if Len(ReceiveZipCode) = 6 then
						'
					end if
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text
					ReceiveAddr2		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text

					deliverymemo		= objMasterOneXML.selectSingleNode("DlvMemoCont").text
					if (deliverymemo = "null") then
						deliverymemo = ""
					end if

					etc1				= ""
					countryCode			= "KR"

					'// 디테일
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					for j = 0 to detailCnt - 1
						set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey		= objDetailOneXML.selectSingleNode("ProdSeq").text

						SellPrice			= objDetailOneXML.selectSingleNode("ordPrice").text

						''RealSellPrice		= objDetailOneXML.selectSingleNode("buyRealPrice").text
						RealSellPrice		= objDetailOneXML.selectSingleNode("ordPrice").text ''2013/07/01 수정

						ItemOrderCount		= objDetailOneXML.selectSingleNode("ordQty").text

						requireDetail		= objDetailOneXML.selectSingleNode("GoodsChocDesc").text
						if (requireDetail = "null") then
							requireDetail = ""
						end if

						orderDlvPay			= 0

						tmpStr = Split(objDetailOneXML.selectSingleNode("CorpItemNo").text, "_")
						matchItemID = ""        ''초기화 2013/08/19 추가
						matchItemOption = ""    ''초기화 2013/08/19 추가
						IF IsArray(tmpStr) then

						    if (Ubound(tmpStr)>=0) then
        						matchItemID			= tmpStr(0)
        					end if

        					if (Ubound(tmpStr)>=1) then
    						    matchItemOption		= tmpStr(1)
    						end if
                            'rw Ubound(tmpStr)
                            'rw matchItemID
                            'rw matchItemOption
    						if (matchItemID="null") then matchItemID=""
                        end if

                        if (matchItemID="") then
                            matchItemID =  objDetailOneXML.selectSingleNode("EntrProdNo").text

                        end if


						partnerItemName		= objDetailOneXML.selectSingleNode("ProdName").text
						if (matchItemID="791471") then
						    partnerItemName     = "[빠띠라인] 미끄럼 방지 튜브매트 특대 100 x 120cm 6종 택1"'replace(partnerItemName,"&nbsp;","")
						elseif (matchItemID="635526") then
						    partnerItemName = replace(partnerItemName,"&nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"＆nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
                            partnerItemName = replace(partnerItemName,"  "," ")

					    end if

						partnerOptionName	= Trim(objDetailOneXML.selectSingleNode("prodOption").text)

						outMallGoodsNo		= objDetailOneXML.selectSingleNode("ProdCode").text

						if (partnerOptionName = "null") then
							partnerOptionName = ""
						end if

						partnerItemID		= objDetailOneXML.selectSingleNode("ProdCode").text
						partnerOption		= ""

						if (matchItemOption="") and (partnerOptionName="") then ''2013/07/30 추가
						    matchItemOption="0000"
						end if

						sqlStr = ""
						sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
						sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
						sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '"&matchItemID&"' "
						sqlStr = sqlStr & " and mallid = 'lotteimall' "
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						If (Not rsget.EOF) Then
							matchItemID = rsget("itemid")
							matchItemOption = rsget("itemoption")
						End If
						rsget.Close

                        if (matchItemOption="") then
                            rw "<font color=red>["&matchItemID&"]옵션매칭 실패:["&partnerOptionName&"]</font>"
                            matchItemOption="0000"
                        end if
'                rw "partnerItemID"&partnerItemID
'                rw "partnerOption"&partnerOption
'                rw "matchItemID"&matchItemID
'                rw "matchItemOption"&matchItemOption

						'// CS출고인지
						tmpStr = objDetailOneXML.selectSingleNode("Exchange").text
						isCSOrder = (tmpStr <> "일반")

                        orderCsGbn=""               ''초기화 2013/08/19 추가
						if (Not isCSOrder) then
							orderCsGbn = "0"
						else
							orderCsGbn = "3"
						end if

						''response.write "<br>" & matchItemID & partnerItemName
						retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
								, OrderName, OrderTelNo, OrderHpNo _
								, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
								, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
								, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
								, errCode, errStr )
                        if (retVal) then
                            succCNT=succCNT+1
                        else
                            failCNT=failCNT+1
                        end if
						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			set objMasterListXML = Nothing

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite and c.sellsite='"&FRectSellSite&"'"
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			''rsget.Open strSql, dbget, 1
		ElseIf (FRectSellSite = "lotteCom") then

		    ''rw objData
		    ''dbget.close : response.end

			'// 카운트
			on Error resume next
			masterCnt = (xmlDOM.selectNodes("/Response/Result/OrderCount").item(0).text * 1)
			on Error Goto 0
			If masterCnt = "" Then
				rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				rw "(" & masterCnt & "건)"
			End If

			If (masterCnt > 0) Then
				set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
				masterCnt = objMasterListXML.length

				For i = 0 to masterCnt - 1
					Set objMasterOneXML = objMasterListXML.item(i)
					SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					SellDate			= objMasterOneXML.selectSingleNode("TrdDate").text
					PayType				= "50"
					Paydate				= objMasterOneXML.selectSingleNode("TrdDate").text
					OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("OrderName").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("OrderTelNo").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("OrderHpNo").text
					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/recvName").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text
					ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text
					If Len(ReceiveZipCode) = 6 Then
						'
					End If
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text
					ReceiveAddr2		= objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text
					deliverymemo		= objMasterOneXML.selectSingleNode("DlvMemoCont").text
					If (deliverymemo = "null") Then
						deliverymemo = ""
					End If

					etc1				= ""
					countryCode			= "KR"

                    ''201407263910626 CASE 수령인명에 수령인 주소가 들어가 있음 //2014/07/28
                    ''if (ReceiveName=ReceiveAddr1&ReceiveAddr2) then
                    if Len(ReceiveName)>32 then
                        ReceiveName = OrderName
                    end if

					'// 디테일
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					For j = 0 to detailCnt - 1
						Set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey		= objDetailOneXML.selectSingleNode("ProdSeq").text
						SellPrice			= objDetailOneXML.selectSingleNode("ordPrice").text
						''RealSellPrice		= objDetailOneXML.selectSingleNode("buyRealPrice").text
						RealSellPrice		= objDetailOneXML.selectSingleNode("ordPrice").text
						ItemOrderCount		= objDetailOneXML.selectSingleNode("ordQty").text
						requireDetail		= objDetailOneXML.selectSingleNode("GoodsChocDesc").text
						If (requireDetail = "null") Then
							requireDetail = ""
						End If

						orderDlvPay			= 0
						tmpStr = Split(objDetailOneXML.selectSingleNode("EntrProdNo").text, "_")
						matchItemID = ""
						matchItemOption = ""

                        If (matchItemID="") Then
                            matchItemID =  objDetailOneXML.selectSingleNode("EntrProdNo").text
                            matchItemOption = getOptionCodByOptionNameLotte(objDetailOneXML.selectSingleNode("EntrProdNo").text, Trim(objDetailOneXML.selectSingleNode("prodOption").text))
                        End If

						partnerItemName		= objDetailOneXML.selectSingleNode("ProdName").text
						if (matchItemID="791471") then
						    partnerItemName     = "[빠띠라인] 미끄럼 방지 튜브매트 특대 100 x 120cm 6종 택1"'replace(partnerItemName,"&nbsp;","")
						elseif (matchItemID="635526") then
						    partnerItemName = replace(partnerItemName,"&nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"＆nbsp;"," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
						    partnerItemName = replace(partnerItemName,"   "," ")
                            partnerItemName = replace(partnerItemName,"  "," ")
					    end if

						partnerOptionName	= Trim(objDetailOneXML.selectSingleNode("prodOption").text)
						outMallGoodsNo		= objDetailOneXML.selectSingleNode("ProdCode").text

						if (partnerOptionName = "null") then
							partnerOptionName = ""
						end if

						partnerItemID		= objDetailOneXML.selectSingleNode("ProdCode").text
						partnerOption		= ""

						if (matchItemOption="") and (partnerOptionName="") then
						    matchItemOption="0000"
						end if

                        if (matchItemOption="") then
                            rw "<font color=red>["&matchItemID&"]옵션매칭 실패:["&partnerOptionName&"]</font>"
                            matchItemOption="0000"
                        end if
'                rw "partnerItemID"&partnerItemID
'                rw "partnerOption"&partnerOption
'                rw "matchItemID"&matchItemID
'                rw "matchItemOption"&matchItemOption

						'// CS출고인지
						tmpStr = objDetailOneXML.selectSingleNode("Exchange").text
						isCSOrder = (tmpStr <> "일반")

                        orderCsGbn=""
						if (Not isCSOrder) then
							orderCsGbn = "0"
						else
							orderCsGbn = "3"
						end if

						''response.write "<br>" & matchItemID & partnerItemName
                        if (orderCsGbn<>"3") then ''조건 추가 2014/01/15
    						retVal = saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
    								, OrderName, OrderTelNo, OrderHpNo _
    								, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
    								, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
    								, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
    								, errCode, errStr )

                            if (retVal) then
                                succCNT=succCNT+1
                            else
                                failCNT=failCNT+1
                            end if
                        end if

						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			set objMasterListXML = Nothing

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite and c.sellsite='"&FRectSellSite&"'"
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			''rsget.Open strSql, dbget, 1
		ElseIf (FRectSellSite = "interpark") and (FRectGubun="js") then
		    'on Error resume next
			masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
			'on Error Goto 0
			If masterCnt = "" Then
				rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				rw "(" & masterCnt & "건)"
			End If


			if (masterCnt > 0) then
                set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
                masterCnt = objMasterListXML.length

                for i = 0 to masterCnt - 1
                    set objMasterOneXML = objMasterListXML.item(i)
                    SellSite 			= FRectSellSite
					ORD_NO	            = objMasterOneXML.selectSingleNode("ORD_NO").text
					ORDER_DT			= objMasterOneXML.selectSingleNode("ORDER_DT").text
                    PayType				= "50"
                    ''Paydate				= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
                    OrderUserID			= ""
					ORD_NM			= objMasterOneXML.selectSingleNode("ORD_NM").text
					OrderEmail			= ""
					'OrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
					'OrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
					'ReceiveName			= objMasterOneXML.selectSingleNode("RCVR_NM").text
					'ReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
					'ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					'ReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text

                    CLM_NO  = objMasterOneXML.selectSingleNode("CLM_NO").text
                    CLM_SEQ = objMasterOneXML.selectSingleNode("CLM_SEQ").text
                    DEL_AMT = objMasterOneXML.selectNodes("DELIVERY/DELV").item(0).selectSingleNode("DEL_AMT").text
                    if (CLM_SEQ="") then CLM_SEQ="0"
					set objDetailListXML = objMasterOneXML.selectNodes("PRODUCT/PRD")
					detailCnt = objDetailListXML.length
					For j = 0 to detailCnt - 1
						Set objDetailOneXML = objDetailListXML.item(j)
						ORD_SEQ             = objDetailOneXML.selectSingleNode("ORD_SEQ").text          ''주문순번
						SETL_DT             = objDetailOneXML.selectSingleNode("SETL_DT").text           ''매출 정산일자(YYYYMMDD)
						ENTR_PRD_NO         = objDetailOneXML.selectSingleNode("ENTR_PRD_NO").text          ''제휴업체 상품 코드
						OPT_NO              = objDetailOneXML.selectSingleNode("OPT_NO").text               ''제휴업체 옵션코드

						PRD_NO              = objDetailOneXML.selectSingleNode("PRD_NO").text               ''인터파크상품코드
                        OPT_PRD_NO          = objDetailOneXML.selectSingleNode("OPT_PRD_NO").text           ''옵션코드

						OLD_SALE_UNITCOST   = objDetailOneXML.selectSingleNode("OLD_SALE_UNITCOST").text    ''원판매단가
                        SALE_UNITCOST       = objDetailOneXML.selectSingleNode("SALE_UNITCOST").text        ''실판매단가

						ORD_QTY             = objDetailOneXML.selectSingleNode("ORD_QTY").text          ''상품별 주문수량
						ORD_AMT             = objDetailOneXML.selectSingleNode("ORD_AMT").text          ''상품별 주문금액(판매단가 * 수량)
						ENTR_DC_COUPON_AMT  = objDetailOneXML.selectSingleNode("ENTR_DC_COUPON_AMT").text   ''업체발행쿠폰
						DC_COUPON_AMT       = objDetailOneXML.selectSingleNode("DC_COUPON_AMT").text        ''인터파크발행쿠폰
						SALE_FEE            = objDetailOneXML.selectSingleNode("SALE_FEE").text         ''판매수수료

						PRE_USE_UNITCOST    = objDetailOneXML.selectSingleNode("PRE_USE_UNITCOST").text ''상품별 선할인단가
						PRE_USE_AMT         = objDetailOneXML.selectSingleNode("PRE_USE_AMT").text      ''상품별 선할인금액(설할인단가 * 수량)

						SMONEY_EXPN_AMT     = objDetailOneXML.selectSingleNode("SMONEY_EXPN_AMT").text  ''추가SMoney
						COMP_AMT            = objDetailOneXML.selectSingleNode("COMP_AMT").text         ''고객보상금액
                        EXPN_AMT            = objDetailOneXML.selectSingleNode("EXPN_AMT").text         ''추가적립금


                        IPOINT_PAY_AMT      = objDetailOneXML.selectSingleNode("IPOINT_PAY_AMT").text   ''I-Point몰 판매수수료
                        IPOINT_SAVE_AMT     = objDetailOneXML.selectSingleNode("IPOINT_SAVE_AMT").text  ''I-Point 적립금
                        SHOP_PAY_AMT        = objDetailOneXML.selectSingleNode("SHOP_PAY_AMT").text     ''샵플러스유지수수료
                        ON_INTEREST_FEE     = objDetailOneXML.selectSingleNode("ON_INTEREST_FEE").text   ''무이자수수료


                         ''SUPPLY_CTRT_SEQ     = objDetailOneXML.selectSingleNode("SUPPLY_CTRT_SEQ").text
                         ''OPT_PARENT_SEQ      = objDetailOneXML.selectSingleNode("OPT_PARENT_SEQ").text ''부모상품순번(옵션)
                         ''OPT_PRD_TP          = objDetailOneXML.selectSingleNode("OPT_PRD_TP").text ''옵션유형

                        sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanData"
        				sqlStr = sqlStr & "(sellsite,ORD_NO,ORDER_DT,ORD_NM,CLM_NO,CLM_SEQ,DEL_AMT,ORD_SEQ ,SETL_DT ,ENTR_PRD_NO ,OPT_NO,PRD_NO,OPT_PRD_NO,OLD_SALE_UNITCOST ,SALE_UNITCOST ,ORD_QTY ,ORD_AMT "
                        sqlStr = sqlStr & ",ENTR_DC_COUPON_AMT,DC_COUPON_AMT ,SALE_FEE,PRE_USE_UNITCOST,PRE_USE_AMT ,SMONEY_EXPN_AMT ,COMP_AMT,EXPN_AMT,IPOINT_PAY_AMT,IPOINT_SAVE_AMT ,SHOP_PAY_AMT,ON_INTEREST_FEE )"
                        sqlStr = sqlStr & " values('"&sellsite&"'"
                        sqlStr = sqlStr & " ,'"&ORD_NO&"'"
                        sqlStr = sqlStr & " ,'"&ORDER_DT&"'"
                        sqlStr = sqlStr & " ,'"&ORD_NM&"'"
                        sqlStr = sqlStr & " ,'"&CLM_NO&"'"
                        sqlStr = sqlStr & " ,'"&CLM_SEQ&"'"
                        sqlStr = sqlStr & " ,'"&DEL_AMT&"'"
                        sqlStr = sqlStr & " ,'"&ORD_SEQ&"'"
                        sqlStr = sqlStr & " ,'"&SETL_DT&"'"
                        sqlStr = sqlStr & " ,'"&ENTR_PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OPT_NO&"'"
                        sqlStr = sqlStr & " ,'"&PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OPT_PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OLD_SALE_UNITCOST&"'"
                        sqlStr = sqlStr & " ,'"&SALE_UNITCOST&"'"
                        sqlStr = sqlStr & " ,'"&ORD_QTY&"'"
                        sqlStr = sqlStr & " ,'"&ORD_AMT&"'"

                        sqlStr = sqlStr & " ,'"&ENTR_DC_COUPON_AMT&"'"
                        sqlStr = sqlStr & " ,'"&DC_COUPON_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SALE_FEE&"'"
        				sqlStr = sqlStr & " ,'"&PRE_USE_UNITCOST&"'"
        				sqlStr = sqlStr & " ,'"&PRE_USE_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SMONEY_EXPN_AMT&"'"
        				sqlStr = sqlStr & " ,'"&COMP_AMT&"'"
        				sqlStr = sqlStr & " ,'"&EXPN_AMT&"'"
        				sqlStr = sqlStr & " ,'"&IPOINT_PAY_AMT&"'"
        				sqlStr = sqlStr & " ,'"&IPOINT_SAVE_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SHOP_PAY_AMT&"'"
        				sqlStr = sqlStr & " ,'"&ON_INTEREST_FEE&"'"
        				sqlStr = sqlStr & ")"

        				if CLM_SEQ="0" then
        				    dbget.Execute sqlStr
                        end if

						set objDetailOneXML = Nothing
					next
					set objDetailListXML = Nothing

					set objMasterOneXML = Nothing
                next
				set objMasterListXML = Nothing


            end if
        ElseIf (FRectSellSite = "gseshop") then

		    ''rw objData
		    ''dbget.close : response.end

			'// 카운트
			on Error resume next
			masterCnt = 1
			on Error Goto 0
			If masterCnt = "" Then
				''rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				''rw "(" & masterCnt & "건)"
			End If

			If (masterCnt > 0) Then
				Set objMasterOneXML = xmlDoc.documentElement.selectSingleNode("PurchaseOrder_V01_00")

				SellSite 			= FRectSellSite
				OutMallOrderSerial	= objMasterOneXML.selectSingleNode("MessageBody/OrdNo").text

				'// 1시간에 1회 출고지시한다.
				SellDate			= objMasterOneXML.selectSingleNode("MessageBody/shipDirDt").text + " " + objMasterOneXML.selectSingleNode("MessageBody/shipDirDt").text + ":00:00"

				PayType				= "50"
				Paydate				= objMasterOneXML.selectSingleNode("MessageBody/shipDirDt").text + " " + objMasterOneXML.selectSingleNode("MessageBody/shipDirDt").text + ":00:00"
				OrderUserID			= ""
				OrderName			= objMasterOneXML.selectSingleNode("MessageBody/rlOrdPrsnNm").text
				OrderEmail			= ""
				OrderTelNo			= objMasterOneXML.selectSingleNode("MessageBody/rlOrdPrsnHomTel").text
				OrderHpNo			= objMasterOneXML.selectSingleNode("MessageBody/rlOrdPrsnCelTel").text
				ReceiveName			= objMasterOneXML.selectSingleNode("MessageBody/custPrsnNm").text
				ReceiveTelNo		= objMasterOneXML.selectSingleNode("MessageBody/custPrsnHomTel").text
				ReceiveHpNo			= objMasterOneXML.selectSingleNode("MessageBody/custPrsnCelTel").text
				ReceiveZipCode		= objMasterOneXML.selectSingleNode("MessageBody/delivZip").text
				ReceiveAddr1		= objMasterOneXML.selectSingleNode("MessageBody/delivAddr1").text
				ReceiveAddr2		= objMasterOneXML.selectSingleNode("MessageBody/delivAddr2").text
				deliverymemo		= objMasterOneXML.selectSingleNode("MessageBody/delivMsg").text

				etc1				= ""
				countryCode			= "KR"


				OrgDetailKey		= Left(objMasterOneXML.selectSingleNode("MessageBody/ordItemNo").text, 5) * 1
				SellPrice			= objMasterOneXML.selectSingleNode("MessageBody/stdUprc").text
				RealSellPrice		= objMasterOneXML.selectSingleNode("MessageBody/salePrc").text
				ItemOrderCount		= objMasterOneXML.selectSingleNode("MessageBody/ordQty").text
				requireDetail		= ""

				orderDlvPay			= 0				'// 없다??
				matchItemID = objMasterOneXML.selectSingleNode("MessageBody/supPrdCd").text
				matchItemOption = ""

				partnerItemName		= objMasterOneXML.selectSingleNode("MessageBody/prdNm").text
				partnerOptionName	= objMasterOneXML.selectSingleNode("MessageBody/attrTypNm1").text
				outMallGoodsNo		= objMasterOneXML.selectSingleNode("MessageBody/prdCd").text

				partnerItemID		= objMasterOneXML.selectSingleNode("MessageBody/supPrdCd").text
				partnerOption		= ""

				if (matchItemOption="") and (partnerOptionName="") then
					matchItemOption="0000"
				end if

                if (matchItemOption="") then
                    ''rw "<font color=red>["&matchItemID&"]옵션매칭 실패:["&partnerOptionName&"]</font>"
                    matchItemOption="0000"
                end if

				'// CS출고인지
				tmpStr = objMasterOneXML.selectSingleNode("MessageBody/ordTypeCd").text
				isCSOrder = (tmpStr <> "O")			'// (주문:O, 반품:R, 교환:X, 취소:C)

                orderCsGbn=""
				if (Not isCSOrder) then
					orderCsGbn = "0"
				else
					orderCsGbn = "3"
				end if

                if (orderCsGbn<>"3") then ''조건 추가 2014/01/15
    				retVal = saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
    				, OrderName, OrderTelNo, OrderHpNo _
    				, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
    				, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
    				, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
    				, errCode, errStr )

                    if (retVal) then
                        succCNT=succCNT+1

						response.write "<PurchaseOrder_V01_00>" + vbCrLf
						response.write "<MessageHeader>" + vbCrLf
						response.write "        <Sender>TENBYTEN</Sender>" + vbCrLf
						response.write "        <Receiver>GS SHOP</Receiver>" + vbCrLf
						response.write "        <MessageID>" + objMasterOneXML.selectSingleNode("MessageHeader/MessageID").text + "</MessageID>" + vbCrLf
						response.write "        <DateTime>" + objMasterOneXML.selectSingleNode("MessageHeader/DateTime").text + "</DateTime>" + vbCrLf
						response.write "        <ProcessType>S</ProcessType>" + vbCrLf
						response.write "        <DocumentID>" + objMasterOneXML.selectSingleNode("MessageHeader/DocumentID").text + "</DocumentID>" + vbCrLf
						response.write "        <UniqueID>" + objMasterOneXML.selectSingleNode("MessageHeader/UniqueID").text + "</UniqueID>" + vbCrLf
						response.write "        <ErrorOccur></ErrorOccur>" + vbCrLf
						response.write "        <ErrorMessage></ErrorMessage>" + vbCrLf
						response.write "</MessageHeader>" + vbCrLf
						response.write "<MessageBody>" + vbCrLf
						response.write "        <PurchaseOrders>" + vbCrLf
						response.write "                <ordItemNo>" + objMasterOneXML.selectSingleNode("MessageBody/ordItemNo").text + "</ordItemNo>" + vbCrLf
						response.write "                <ordNo>" + objMasterOneXML.selectSingleNode("MessageBody/ordNo").text + "</ordNo>" + vbCrLf
						response.write "                <OrderGenerationDate>" + Left(Now(), 10) + "</OrderGenerationDate>" + vbCrLf
						response.write "                <ProductLineItem>" + vbCrLf
						response.write "                        <ConfirmedDeliveryDate>" + Left(Now(), 10) + "</ConfirmedDeliveryDate>" + vbCrLf
						response.write "                        <sendFg>S</sendFg>" + vbCrLf
						response.write "                </ProductLineItem>" + vbCrLf
						response.write "        </PurchaseOrders>" + vbCrLf
						response.write "</MessageBody>" + vbCrLf
						response.write "</PurchaseOrder_V01_00>" + vbCrLf
                    else
                        failCNT=failCNT+1
                    end if
                end if

			end if

			strSql = " update c "
			strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			strSql = strSql + " from "
			strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			strSql = strSql + " on "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.SellSite = o.SellSite and c.sellsite='"&FRectSellSite&"'"
			strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			strSql = strSql + " where "
			strSql = strSql + " 	1 = 1 "
			strSql = strSql + " 	and c.orderserial is NULL "
			strSql = strSql + " 	and o.orderserial is not NULL "
			''rw strSql
			''rsget.Open strSql, dbget, 1
		ElseIf (FRectSellSite = "interpark") and (FRectGubun="js") then
		    'on Error resume next
			masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
			'on Error Goto 0
			If masterCnt = "" Then
				rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				rw "(" & masterCnt & "건)"
			End If


			if (masterCnt > 0) then
                set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
                masterCnt = objMasterListXML.length

                for i = 0 to masterCnt - 1
                    set objMasterOneXML = objMasterListXML.item(i)
                    SellSite 			= FRectSellSite
					ORD_NO	            = objMasterOneXML.selectSingleNode("ORD_NO").text
					ORDER_DT			= objMasterOneXML.selectSingleNode("ORDER_DT").text
                    PayType				= "50"
                    ''Paydate				= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
                    OrderUserID			= ""
					ORD_NM			= objMasterOneXML.selectSingleNode("ORD_NM").text
					OrderEmail			= ""
					'OrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
					'OrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
					'ReceiveName			= objMasterOneXML.selectSingleNode("RCVR_NM").text
					'ReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
					'ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					'ReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text

                    CLM_NO  = objMasterOneXML.selectSingleNode("CLM_NO").text
                    CLM_SEQ = objMasterOneXML.selectSingleNode("CLM_SEQ").text
                    DEL_AMT = objMasterOneXML.selectNodes("DELIVERY/DELV").item(0).selectSingleNode("DEL_AMT").text
                    if (CLM_SEQ="") then CLM_SEQ="0"
					set objDetailListXML = objMasterOneXML.selectNodes("PRODUCT/PRD")
					detailCnt = objDetailListXML.length
					For j = 0 to detailCnt - 1
						Set objDetailOneXML = objDetailListXML.item(j)
						ORD_SEQ             = objDetailOneXML.selectSingleNode("ORD_SEQ").text          ''주문순번
						SETL_DT             = objDetailOneXML.selectSingleNode("SETL_DT").text           ''매출 정산일자(YYYYMMDD)
						ENTR_PRD_NO         = objDetailOneXML.selectSingleNode("ENTR_PRD_NO").text          ''제휴업체 상품 코드
						OPT_NO              = objDetailOneXML.selectSingleNode("OPT_NO").text               ''제휴업체 옵션코드

						PRD_NO              = objDetailOneXML.selectSingleNode("PRD_NO").text               ''인터파크상품코드
                        OPT_PRD_NO          = objDetailOneXML.selectSingleNode("OPT_PRD_NO").text           ''옵션코드

						OLD_SALE_UNITCOST   = objDetailOneXML.selectSingleNode("OLD_SALE_UNITCOST").text    ''원판매단가
                        SALE_UNITCOST       = objDetailOneXML.selectSingleNode("SALE_UNITCOST").text        ''실판매단가

						ORD_QTY             = objDetailOneXML.selectSingleNode("ORD_QTY").text          ''상품별 주문수량
						ORD_AMT             = objDetailOneXML.selectSingleNode("ORD_AMT").text          ''상품별 주문금액(판매단가 * 수량)
						ENTR_DC_COUPON_AMT  = objDetailOneXML.selectSingleNode("ENTR_DC_COUPON_AMT").text   ''업체발행쿠폰
						DC_COUPON_AMT       = objDetailOneXML.selectSingleNode("DC_COUPON_AMT").text        ''인터파크발행쿠폰
						SALE_FEE            = objDetailOneXML.selectSingleNode("SALE_FEE").text         ''판매수수료

						PRE_USE_UNITCOST    = objDetailOneXML.selectSingleNode("PRE_USE_UNITCOST").text ''상품별 선할인단가
						PRE_USE_AMT         = objDetailOneXML.selectSingleNode("PRE_USE_AMT").text      ''상품별 선할인금액(설할인단가 * 수량)

						SMONEY_EXPN_AMT     = objDetailOneXML.selectSingleNode("SMONEY_EXPN_AMT").text  ''추가SMoney
						COMP_AMT            = objDetailOneXML.selectSingleNode("COMP_AMT").text         ''고객보상금액
                        EXPN_AMT            = objDetailOneXML.selectSingleNode("EXPN_AMT").text         ''추가적립금


                        IPOINT_PAY_AMT      = objDetailOneXML.selectSingleNode("IPOINT_PAY_AMT").text   ''I-Point몰 판매수수료
                        IPOINT_SAVE_AMT     = objDetailOneXML.selectSingleNode("IPOINT_SAVE_AMT").text  ''I-Point 적립금
                        SHOP_PAY_AMT        = objDetailOneXML.selectSingleNode("SHOP_PAY_AMT").text     ''샵플러스유지수수료
                        ON_INTEREST_FEE     = objDetailOneXML.selectSingleNode("ON_INTEREST_FEE").text   ''무이자수수료


                         ''SUPPLY_CTRT_SEQ     = objDetailOneXML.selectSingleNode("SUPPLY_CTRT_SEQ").text
                         ''OPT_PARENT_SEQ      = objDetailOneXML.selectSingleNode("OPT_PARENT_SEQ").text ''부모상품순번(옵션)
                         ''OPT_PRD_TP          = objDetailOneXML.selectSingleNode("OPT_PRD_TP").text ''옵션유형

                        sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanData"
        				sqlStr = sqlStr & "(sellsite,ORD_NO,ORDER_DT,ORD_NM,CLM_NO,CLM_SEQ,DEL_AMT,ORD_SEQ ,SETL_DT ,ENTR_PRD_NO ,OPT_NO,PRD_NO,OPT_PRD_NO,OLD_SALE_UNITCOST ,SALE_UNITCOST ,ORD_QTY ,ORD_AMT "
                        sqlStr = sqlStr & ",ENTR_DC_COUPON_AMT,DC_COUPON_AMT ,SALE_FEE,PRE_USE_UNITCOST,PRE_USE_AMT ,SMONEY_EXPN_AMT ,COMP_AMT,EXPN_AMT,IPOINT_PAY_AMT,IPOINT_SAVE_AMT ,SHOP_PAY_AMT,ON_INTEREST_FEE )"
                        sqlStr = sqlStr & " values('"&sellsite&"'"
                        sqlStr = sqlStr & " ,'"&ORD_NO&"'"
                        sqlStr = sqlStr & " ,'"&ORDER_DT&"'"
                        sqlStr = sqlStr & " ,'"&ORD_NM&"'"
                        sqlStr = sqlStr & " ,'"&CLM_NO&"'"
                        sqlStr = sqlStr & " ,'"&CLM_SEQ&"'"
                        sqlStr = sqlStr & " ,'"&DEL_AMT&"'"
                        sqlStr = sqlStr & " ,'"&ORD_SEQ&"'"
                        sqlStr = sqlStr & " ,'"&SETL_DT&"'"
                        sqlStr = sqlStr & " ,'"&ENTR_PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OPT_NO&"'"
                        sqlStr = sqlStr & " ,'"&PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OPT_PRD_NO&"'"
                        sqlStr = sqlStr & " ,'"&OLD_SALE_UNITCOST&"'"
                        sqlStr = sqlStr & " ,'"&SALE_UNITCOST&"'"
                        sqlStr = sqlStr & " ,'"&ORD_QTY&"'"
                        sqlStr = sqlStr & " ,'"&ORD_AMT&"'"

                        sqlStr = sqlStr & " ,'"&ENTR_DC_COUPON_AMT&"'"
                        sqlStr = sqlStr & " ,'"&DC_COUPON_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SALE_FEE&"'"
        				sqlStr = sqlStr & " ,'"&PRE_USE_UNITCOST&"'"
        				sqlStr = sqlStr & " ,'"&PRE_USE_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SMONEY_EXPN_AMT&"'"
        				sqlStr = sqlStr & " ,'"&COMP_AMT&"'"
        				sqlStr = sqlStr & " ,'"&EXPN_AMT&"'"
        				sqlStr = sqlStr & " ,'"&IPOINT_PAY_AMT&"'"
        				sqlStr = sqlStr & " ,'"&IPOINT_SAVE_AMT&"'"
        				sqlStr = sqlStr & " ,'"&SHOP_PAY_AMT&"'"
        				sqlStr = sqlStr & " ,'"&ON_INTEREST_FEE&"'"
        				sqlStr = sqlStr & ")"

        				if CLM_SEQ="0" then
        				    dbget.Execute sqlStr
                        end if

						set objDetailOneXML = Nothing
					next
					set objDetailListXML = Nothing

					set objMasterOneXML = Nothing
                next
				set objMasterListXML = Nothing


            end if
        ElseIf (FRectSellSite = "interpark") then
            'on Error resume next
			masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
			'on Error Goto 0
			If masterCnt = "" Then
				rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				rw "(" & masterCnt & "건)"
			End If

			if (masterCnt > 0) then
                set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
                masterCnt = objMasterListXML.length

                for i = 0 to masterCnt - 1
                    set objMasterOneXML = objMasterListXML.item(i)
                    SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("ORD_NO").text
					SellDate			= objMasterOneXML.selectSingleNode("ORDER_DT").text
                    PayType				= "50"
                    Paydate				= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
                    OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("ORD_NM").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
					ReceiveName			= objMasterOneXML.selectSingleNode("RCVR_NM").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
					ReceiveHpNo			= objMasterOneXML.selectSingleNode("DELI_MOBILE").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DELI_ADDR1").text
					ReceiveAddr2		= objMasterOneXML.selectSingleNode("DELI_ADDR2").text
					deliverymemo		= objMasterOneXML.selectSingleNode("DELI_COMMENT").text
''----------
                    DEL_AMT = objMasterOneXML.selectNodes("DELIVERY/DELV").item(0).selectSingleNode("DEL_AMT").text

					set objDetailListXML = objMasterOneXML.selectNodes("PRODUCT/PRD")
					detailCnt = objDetailListXML.length
					For j = 0 to detailCnt - 1
						Set objDetailOneXML = objDetailListXML.item(j)
						ORD_SEQ             = objDetailOneXML.selectSingleNode("ORD_SEQ").text          ''주문순번
						ENTR_PRD_NO         = objDetailOneXML.selectSingleNode("ENTR_PRD_NO").text      ''제휴업체 상품 코드
						OPT_NO              = objDetailOneXML.selectSingleNode("OPT_NO").text           ''제휴업체 옵션코드

						PRD_NO              = objDetailOneXML.selectSingleNode("PRD_NO").text           ''인터파크상품코드
                        OPT_PRD_NO          = objDetailOneXML.selectSingleNode("OPT_PRD_NO").text       ''옵션코드

                        OPT_NM              = objDetailOneXML.selectSingleNode("OPT_NM").text           ''옵션명(선택형+입력형)
                        SEL_OPT_NM          = objDetailOneXML.selectSingleNode("SEL_OPT_NM").text       ''옵션명(선택형)
                        IN_OPT_NM           = objDetailOneXML.selectSingleNode("IN_OPT_NM").text        ''옵션명(입력형)


						''OLD_SALE_UNITCOST   = objDetailOneXML.selectSingleNode("OLD_SALE_UNITCOST").text    ''원판매단가
                        SALE_UNITCOST       = objDetailOneXML.selectSingleNode("SALE_UNITCOST").text        ''실판매단가

						ORD_QTY             = objDetailOneXML.selectSingleNode("ORD_QTY").text          ''상품별 주문수량
						ORD_AMT             = objDetailOneXML.selectSingleNode("ORD_AMT").text          ''상품별 주문금액(판매단가 * 수량)
						ENTR_DC_COUPON_AMT  = objDetailOneXML.selectSingleNode("ENTR_DC_COUPON_AMT").text   ''업체(TEN)발행쿠폰
						DC_COUPON_AMT       = objDetailOneXML.selectSingleNode("DC_COUPON_AMT").text        ''인터파크발행쿠폰

						PRE_USE_UNITCOST    = objDetailOneXML.selectSingleNode("PRE_USE_UNITCOST").text ''상품별 선할인단가

                        OLD_SALE_UNITCOST   = SALE_UNITCOST+ENTR_DC_COUPON_AMT+DC_COUPON_AMT ''원판매가May

                         ''SUPPLY_CTRT_SEQ     = objDetailOneXML.selectSingleNode("SUPPLY_CTRT_SEQ").text
                         ''OPT_PARENT_SEQ      = objDetailOneXML.selectSingleNode("OPT_PARENT_SEQ").text ''부모상품순번(옵션)
                         ''OPT_PRD_TP          = objDetailOneXML.selectSingleNode("OPT_PRD_TP").text ''옵션유형

'                        sqlStr = " Update T"
'                        sqlStr = sqlStr & " SET realSellPrice="&SALE_UNITCOST&""
'                        sqlStr = sqlStr & " ,PRE_USE_UNITCOST="&PRE_USE_UNITCOST&""
'                        sqlStr = sqlStr & " ,tenCpnUint="&ENTR_DC_COUPON_AMT&""
'                        sqlStr = sqlStr & " ,mallCpnUnit="&DC_COUPON_AMT&""
'                        sqlStr = sqlStr & " From db_temp.dbo.tbl_xSite_tmporder_Back T"
'        				sqlStr = sqlStr & " where sellsite='"&sellsite&"'"
'                        sqlStr = sqlStr & " and outmallorderserial='"&OutMallOrderSerial&"'"
'                        sqlStr = sqlStr & " and OrgDetailKey='"&ORD_SEQ&"'"
'                        sqlStr = sqlStr & " and orderitemid='"&ENTR_PRD_NO&"'"
'
'        				dbget.Execute sqlStr

                        sqlStr = " Update T"
                        sqlStr = sqlStr & " SET realSellPrice=(CASE WHEN SellPrice<>realSellPrice THEN realSellPrice ELSE "&SALE_UNITCOST&" END)"
                        sqlStr = sqlStr & " ,PRE_USE_UNITCOST="&PRE_USE_UNITCOST&""
                        sqlStr = sqlStr & " ,tenCpnUint="&ENTR_DC_COUPON_AMT&""
                        sqlStr = sqlStr & " ,mallCpnUnit="&DC_COUPON_AMT&""
                        sqlStr = sqlStr & " From db_temp.dbo.tbl_xSite_tmporder T"
        				sqlStr = sqlStr & " where sellsite='"&sellsite&"'"
                        sqlStr = sqlStr & " and outmallorderserial='"&OutMallOrderSerial&"'"
                        sqlStr = sqlStr & " and OrgDetailKey='"&ORD_SEQ&"'"
                        sqlStr = sqlStr & " and orderitemid='"&ENTR_PRD_NO&"'"
                        sqlStr = sqlStr & " and mallCpnUnit is NULL" ''2014/02/01
''rw sqlStr
        				dbget.Execute sqlStr

						set objDetailOneXML = Nothing
					next
					set objDetailListXML = Nothing

					set objMasterOneXML = Nothing
                next
				set objMasterListXML = Nothing


            end if
		ElseIf (FRectSellSite = "interpark") then
		    'rw objData

		    'on Error resume next
			masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
			'on Error Goto 0
			If masterCnt = "" Then
				rw "해당기간에 맞는 데이터가 없습니다"
				Exit Function
			Else
				rw "(" & masterCnt & "건)"
			End If

            if (masterCnt > 0) then
                set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
                masterCnt = objMasterListXML.length

                for i = 0 to masterCnt - 1
					set objMasterOneXML = objMasterListXML.item(i)
					SellSite 			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("ORD_NO").text
					SellDate			= objMasterOneXML.selectSingleNode("ORDER_DT").text
                    PayType				= "50"
                    Paydate				= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
                    OrderUserID			= ""
					OrderName			= objMasterOneXML.selectSingleNode("ORD_NM").text
					OrderEmail			= ""
					OrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
					OrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
					ReceiveName			= objMasterOneXML.selectSingleNode("RCVR_NM").text
					ReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
					'ReceiveHpNo			= objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text
					ReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text

					rw OutMallOrderSerial
					rw SellDate
				next
				set objMasterListXML = Nothing
            end if
		else
			ErrMsg = "파싱에 실패했습니다."
		end if
		Set xmlDOM = Nothing

		if (failCNT<>0) then
		    rw "["&failCNT&"] 건 실패"
		end if

		if (succCNT<>0) then
		    rw "["&succCNT&"] 건 성공"
		end if

	end function

	public function GetXmlFromWeb()
		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000  ''2013/08/01 추가
		objXML.Send()

		if objXML.Status = "200" then
			if (sellsite <> "gseshop") then
				objData = BinaryToText(objXML.ResponseBody, "euc-kr")
			end if
		end if

		Set objXML  = Nothing
	end function

	public function RequestXmlFromWeb()
		objData = ""

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000  ''2013/08/01 추가
		objXML.Send()

		if objXML.Status <> "200" then
			ErrMsg = "SERVER ERROR"
		end if

		Set objXML  = Nothing
	end function

	public function GetXMLURL()
		dim tmp

		tmp = GetxSiteDateFormat(FRectStartYYYYMMDD)

		if (tmp = "") then
			GetXMLURL = ""
			ErrMsg = "날자형식이 지정되지 않았습니다."
			exit function
		end if

		if (sellsite = "gseshop") then
			if (FRectGubun = "new") then
				GetXMLURL = FRectAPIURL
			elseif (FRectGubun = "all") then
				GetXMLURL = FRectAPIURL
			else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			end if
		Elseif (sellsite = "lotteimall") then
			if (FRectGubun = "new") then
				'// 신규(01)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=01"
			elseif (FRectGubun = "all") then
				'// 전체(02)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=02"
			else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			end if
		Elseif (sellsite = "lotteCom") Then
			If (FRectGubun = "new") Then
				'// 신규(01)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=01"
			ElseIf (FRectGubun = "all") Then
				'// 전체(02)
				GetXMLURL = FRectAPIURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + CStr(FRectAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&SelOption=02"
			Else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			End If
		Elseif (sellsite = "interpark") Then
		    If (FRectGubun = "new") Then
				'// 신규(01)
				GetXMLURL = FRectAPIURL + "/order/OrderClmAPI.do?_method=orderListForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "000000" + "&sc.endDate=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "235959"
			ElseIf (FRectGubun = "all") Then
				'// 전체(02)
				GetXMLURL = FRectAPIURL + "/order/OrderClmAPI.do?_method=orderListDelvForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "000000" + "&sc.endDate=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "235959"
			ElseIf (FRectGubun = "js") Then
			    GetXMLURL = FRectAPIURL + "/order/OrderClmAPI.do?_method=OMSettlementListForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.setlDate="+CStr(GetxSiteDateFormat(FRectStartYYYYMMDD))
			Else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			End If
		else
			GetXMLURL = ""
			ErrMsg = "등록되지 않은 제휴몰입니다.[2]"
		end if
	end function

	public function GetxSiteDateFormat(dt)
		if (FRectSellSite = "lotteimall") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "lotteCom") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "interpark") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "gseshop") then
			GetxSiteDateFormat = dt
		else
			GetxSiteDateFormat = ""
		end if
	end function



	public function ResetXML()
		Set objXML = Nothing
		Set xmlDOM = Nothing
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FResultCount = 0
		FTotalCount = 0

		Call ResetXML()
	End Sub

	Private Sub Class_Terminate()
		Call ResetXML()
	End Sub

End Class

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr )
    dim paramInfo, retParamInfo
    dim PayType  : PayType  = "50"
    dim sqlStr
	dim countryCode

	if countryCode="" then countryCode="KR"

    saveOrderOneToTmpTable =false

    OrderTelNo = replace(OrderTelNo,")","-")
    OrderHpNo = replace(OrderHpNo,")","-")
    ReceiveTelNo = replace(ReceiveTelNo,")","-")
    ReceiveHpNo = replace(ReceiveHpNo,")","-")

    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
		,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, OutMallOrderSerial)	_
		,Array("@SellDate"	,adDate, adParamInput,, SellDate) _
		,Array("@PayType"	,adVarchar, adParamInput,32, PayType) _
		,Array("@Paydate"	,adDate, adParamInput,, SellDate) _
		,Array("@matchItemID"	,adInteger, adParamInput,, matchItemID) _
		,Array("@matchItemOption"	,adVarchar, adParamInput,4, matchItemOption) _
		,Array("@partnerItemID"	,adVarchar, adParamInput,32, matchItemID) _
		,Array("@partnerItemName"	,adVarchar, adParamInput,128, partnerItemName) _
		,Array("@partnerOption"	,adVarchar, adParamInput,128, matchItemOption) _
		,Array("@partnerOptionName"	,adVarchar, adParamInput,128, partnerOptionName) _
		,Array("@outMallGoodsNo"	,adVarchar, adParamInput,16, outMallGoodsNo) _
		,Array("@OrderUserID"	,adVarchar, adParamInput,32, "") _
		,Array("@OrderName"	,adVarchar, adParamInput,32, OrderName) _
		,Array("@OrderEmail"	,adVarchar, adParamInput,100, "") _
		,Array("@OrderTelNo"	,adVarchar, adParamInput,16, OrderTelNo) _
		,Array("@OrderHpNo"	,adVarchar, adParamInput,16, OrderHpNo) _
		,Array("@ReceiveName"	,adVarchar, adParamInput,32, ReceiveName) _
		,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, ReceiveTelNo) _
		,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, ReceiveHpNo) _
		,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, ReceiveZipCode) _
		,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, ReceiveAddr1) _
		,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, ReceiveAddr2) _
		,Array("@SellPrice"	,adCurrency, adParamInput,, SellPrice) _
		,Array("@RealSellPrice"	,adCurrency, adParamInput,, RealSellPrice) _
		,Array("@ItemOrderCount"	,adInteger, adParamInput,, ItemOrderCount) _
		,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, OrgDetailKey) _
		,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
		,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
		,Array("@deliverymemo"	,adVarchar, adParamInput,400, deliverymemo) _
		,Array("@requireDetail"	,adVarchar, adParamInput,400, requireDetail) _
		,Array("@orderDlvPay"	,adCurrency, adParamInput,, orderDlvPay) _
		,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    	,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_FromXML"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러메세지
    else
        ierrCode = -999
        ierrStr = "상품코드 또는 옵션코드  매칭 실패" & OrgDetailKey & " 상품코드 =" & matchItemID&" 옵션명 = "&partnerOptionName
        rw "["&ierrCode&"]"&ierrStr
        dbget.close() : response.end
    end if

    saveOrderOneToTmpTable = (ierrCode=0)
    if (ierrCode<>0) then
        rw "["&ierrCode&"]"&ierrStr
    end if
end function

public function getChrCount(orgStr, delim)
    dim retCNT : retCNT = 0
    dim buf
    buf = split(orgStr,delim)

    if IsArray(buf) then
        retCNT = UBound(buf)
    end if
    getChrCount = retCNT
end function

public function getOptionCodByOptionNameLotte(iitemid,ioptionname)
    dim retStr, sqlStr : retStr=""
    dim IsDoubleOption, IsTreepleOption
    IF (getChrCount(ioptionname,":")=2) THEN
        IF (getChrCount(ioptionname,",")=1) THEN
            IsDoubleOption = TRUE
        END IF
    ELSEIF (getChrCount(ioptionname,":")=3) THEN  '''디자인:c21,폰트선택:폰트2,리필잉크추가 선택:추가안함
        IF (getChrCount(ioptionname,",")=2) THEN
            IsTreepleOption = TRUE
        END IF
    ENd IF

    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&"'"
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and replace(optionname,'*','')='"&SplitValue(SplitValue(ioptionname,",",0),":",1)&","&SplitValue(SplitValue(ioptionname,",",1),":",1)&","&SplitValue(SplitValue(ioptionname,",",2),":",1)&"'"
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        ''sqlStr = sqlStr & " and optionTypename='"&SplitValue(ioptionname,":",0)&"'"
        sqlStr = sqlStr & " and Replace(Replace(replace(optionname,'*',''),',',''),'#','')=Replace('"&SplitValue(ioptionname,":",1)&"','#','')"
    END IF

	'response.write sqlstr & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>0) THEN retStr = "0000"
    	end if
        rsget.Close

    END IF
    getOptionCodByOptionNameLotte = retStr

'	if retStr="" then
'	    rw sqlStr
'	end if
end function
%>
