<%

'' !!!! 아래 파일이 인클루드 되어 있어야 한다.
''/admin/etc/lotte/inc_dailyAuthCheck.asp
''/lib/classes/etc/lotteitemcls.asp
''/admin/etc/incOutMallCommonFunction.asp


Class CxSiteTmpCSItem

    public Fidx
	public Fdivcd
	public Fdivname
	public Fgubunname
    public FSellSite
    public FOutMallOrderSerial
    public FOrgDetailKey
	public FCSDetailKey
	public FOrderSerial
	public FItemID
	public Fitemoption
	public Fitemno
	public FOutMallItemName
	public FOutMallItemOptionName
    public FOrderName
    public FOrderEmail
    public FOrderTelNo
    public FOrderHpNo
    public FReceiveName
    public FReceiveTelNo
    public FReceiveHpNo
    public FReceiveZipCode
    public FReceiveAddr1
    public FReceiveAddr2
    public Fdeliverymemo
	public Fcurrstate
	public Fdeleteyn
    public FOutMallRegDate
    public FRegDate
	public Ftencsdivname
	public Ftencscnt

	public ForgOutMallOrderSerial				'// 원주문번호(기출고 주문일 경우)

	public function GetCurrStateName
		if (Fcurrstate = "B001") then
			GetCurrStateName = "등록이전"
		elseif (Fcurrstate = "B007") then
			GetCurrStateName = "등록완료"
		else
			GetCurrStateName = Fcurrstate
		end if
	end function

	public function GetCurrStateColor
		if (Fcurrstate = "B001") then
			GetCurrStateColor = "black"
		elseif (Fcurrstate = "B007") then
			GetCurrStateColor = "blue"
		else
			GetCurrStateColor = "red"
		end if
	end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class


Class CxSiteCSOrder
    public FItemList()
	public FOneItem
	public FCurrPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FTotalCount
	public FTotalPage
	public FRectSellSite
	public FRectOrderSerial
	public FRectOutMallOrderSerial
	public FRectCurrState

	public function getCSMasterList()
	    dim i,sqlStr
	    sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" & FPageSize & ") as totPg"
	    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPCS m"
	    sqlStr = sqlStr & " where 1=1"

	    if (FRectSellSite<>"") then
    	    sqlStr = sqlStr & " and SellSite='"&FRectSellSite&"'"
    	end if

	    if (FRectOutMallOrderSerial<>"") then
    	    sqlStr = sqlStr & " and OutMallOrderSerial = '"&FRectOutMallOrderSerial&"'"
    	end if

	    if (FRectOrderSerial<>"") then
    	    sqlStr = sqlStr & " and OrderSerial='"&FRectOrderSerial&"'"
    	end if

	    if (FRectCurrState<>"") then
    	    sqlStr = sqlStr & " and currstate='"&FRectCurrState&"'"
    	end if

		'response.write sqlstr & "<Br>"
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			FTotalCount = rsget("cnt")
			FTotalPage = rsget("totPg")
		rsget.Close

		'지정페이지가 전체 페이지보다 클 때 함수종료
		if Cint(FCurrPage)>Cint(FTotalPage) then
			FResultCount = 0
			exit function
		end if

		sqlStr = "select top " + CStr(FPageSize*FCurrPage) + " m.*, C1.comm_name as divname "
		sqlStr = sqlStr & " , (select IsNull(max(C1.comm_name), '') "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 "
		sqlStr = sqlStr & " on "
		sqlStr = sqlStr & " 	t.divcd=C1.comm_cd "
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " 	t.orderserial = m.orderserial) as tencsdivname "
		sqlStr = sqlStr & " , (select count(*) "
		sqlStr = sqlStr & " from "
		sqlStr = sqlStr & " db_cs.dbo.tbl_new_as_list t "
		sqlStr = sqlStr & " LEFT JOIN [db_cs].[dbo].tbl_cs_comm_code C1 "
		sqlStr = sqlStr & " on "
		sqlStr = sqlStr & " 	t.divcd=C1.comm_cd "
		sqlStr = sqlStr & " where "
		sqlStr = sqlStr & " 	t.orderserial = m.orderserial) as tencscnt "
	    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPCS m"
		sqlStr = sqlStr & " left join [db_cs].[dbo].tbl_cs_comm_code C1 "
		sqlStr = sqlStr & " on "
		sqlStr = sqlStr & " 	m.divcd=C1.comm_cd "
	    sqlStr = sqlStr & " where 1=1"

	    if (FRectSellSite<>"") then
    	    sqlStr = sqlStr & " and m.SellSite='"&FRectSellSite&"'"
    	end if

	    if (FRectOutMallOrderSerial<>"") then
    	    sqlStr = sqlStr & " and OutMallOrderSerial = '"&FRectOutMallOrderSerial&"'"
    	end if

	    if (FRectOrderSerial<>"") then
    	    sqlStr = sqlStr & " and m.OrderSerial='"&FRectOrderSerial&"'"
    	end if

	    if (FRectCurrState<>"") then
    	    sqlStr = sqlStr & " and currstate='"&FRectCurrState&"'"
    	end if

    	sqlStr = sqlStr & " order by m.CSDetailKey desc"

		'response.write sqlStr & "<Br>"
	    rsget.pagesize = FPageSize
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

		FResultCount = rsget.RecordCount-(FPageSize*(FCurrPage-1))

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
			rsget.absolutepage = FCurrPage
			do until rsget.eof
				set FItemList(i) = new CxSiteTmpCSItem

				FItemList(i).Fidx             		= rsget("idx")
				FItemList(i).FSellSite             	= rsget("SellSite")
				FItemList(i).FOutMallOrderSerial    = rsget("OutMallOrderSerial")
				FItemList(i).FOrgDetailKey          = rsget("OrgDetailKey")
				FItemList(i).FCSDetailKey          	= rsget("CSDetailKey")
				FItemList(i).FOrderSerial           = rsget("OrderSerial")
				FItemList(i).Fdivcd             	= rsget("divcd")

				FItemList(i).Fdivname             	= rsget("divname")
				if ((FItemList(i).FSellSite = "lotteCom" or FItemList(i).FSellSite = "lotteimall") and FItemList(i).Fdivcd = "A004") then
					'// 롯데닷컴은 반품/교환 건을 구분할 수 없다.
					FItemList(i).Fdivname             	= "반품/교환"
				elseif (FItemList(i).FSellSite = "cjmall" and FItemList(i).Fdivcd = "A004") then
					'// 반품 : 업배반품인지 텐배반품인지는 상품을 보고 결정한다.
					FItemList(i).Fdivname             	= "반품"
				elseif (FItemList(i).FSellSite = "cjmall" and FItemList(i).Fdivcd = "A011") then
					'// 교환 : 업배교환인지 텐배교환인지는 상품을 보고 결정한다.
					FItemList(i).Fdivname             	= "교환"
				else
					'//
				end if

				FItemList(i).Fgubunname             = rsget("gubunname")
				FItemList(i).FOrderName             = db2html(rsget("OrderName"))
				FItemList(i).FOrderEmail            = rsget("OrderEmail")
				FItemList(i).FOrderTelNo            = rsget("OrderTelNo")
				FItemList(i).FOrderHpNo             = rsget("OrderHpNo")
				FItemList(i).FReceiveName           = db2html(rsget("ReceiveName"))
				FItemList(i).FReceiveTelNo          = rsget("ReceiveTelNo")
				FItemList(i).FReceiveHpNo           = rsget("ReceiveHpNo")
				FItemList(i).FReceiveZipCode        = rsget("ReceiveZipCode")
				FItemList(i).FReceiveAddr1          = rsget("ReceiveAddr1")
				FItemList(i).FReceiveAddr2          = rsget("ReceiveAddr2")
				FItemList(i).Fdeliverymemo          = rsget("deliverymemo")
				FItemList(i).FOutMallRegDate        = rsget("OutMallRegDate")
				FItemList(i).FRegDate             	= rsget("RegDate")

				FItemList(i).FItemID             	= rsget("ItemID")
				FItemList(i).Fitemoption            = rsget("itemoption")

				FItemList(i).Fitemno             	= rsget("itemno")
				if ((FItemList(i).FSellSite = "lotteCom" or FItemList(i).FSellSite = "lotteimall") and FItemList(i).Fdivcd = "A004") then
					'// 롯데닷컴은 반품/교환 건은 접수수량이 없다.
					FItemList(i).Fitemno             	= ""
				end if

				FItemList(i).FOutMallItemName       = db2html(rsget("OutMallItemName"))
				FItemList(i).FOutMallItemOptionName = db2html(rsget("OutMallItemOptionName"))

				FItemList(i).Ftencsdivname          = rsget("tencsdivname")
				FItemList(i).Ftencscnt             	= rsget("tencscnt")

				FItemList(i).Fcurrstate             = rsget("currstate")

				FItemList(i).ForgOutMallOrderSerial	= rsget("orgOutMallOrderSerial")

				i=i+1
				rsget.moveNext
			loop
		end if
		rsget.Close
    end function

    Private Sub Class_Initialize()
		redim  FItemList(0)

		FCurrPage =1
		FPageSize = 20
		FResultCount = 0
		FScrollCount = 10
		FTotalCount =0
		FTotalPage =0
	End Sub

	Private Sub Class_Terminate()
	End Sub

	public Function HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	end Function

	public Function HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	end Function

	public Function StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	end Function

End Class

Class CxSiteCSOrderXML
    public FItemList()
	public FOneItem
	public FResultCount
	public FTotalCount
	public FRectDivCD
	public FRectSellSite
	public FRectOutMallOrderSerial
	public FRectYYYYMMDD
	public FRectStartYYYYMMDD
	public FRectEndYYYYMMDD

	public ErrMsg
	private objXML
	private xmlDOM

	private xmlURL
	private objData

	public function SavexSiteCSOrderListtoDB()
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			''response.write xmlURL

			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "등록되지 않은 제휴몰입니다.[0]"
			end if
		end if

        ''response.write xmlURL
		''response.write "<br>테스트 중입니다.<br>"
		''dbget.close()
		''response.end

		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "롯데닷컴과 통신중에 오류가 발생했습니다."
			end if
		end if

		if (ErrMsg = "") then
			Call ActSavexSiteCSOrderListtoDB()
		end if

    end function

	public function SendxSiteSongjangNo(ord_no, ord_dtl_sn, sendQnt, sendDate, outmallGoodsID, hdc_cd, inv_no)
		ErrMsg = ""

		if (ErrMsg = "") then
			xmlURL = GetXMLURL()
			'response.write xmlURL

			if (xmlURL = "") and (ErrMsg = "") then
				ErrMsg = "등록되지 않은 제휴몰입니다.[0]"
			end if
		end if

		if (ErrMsg = "") then
			Call GetXmlFromWeb()

			if (objData = "") and (ErrMsg = "") then
				ErrMsg = "롯데닷컴과 통신중에 오류가 발생했습니다."
			end if
		end if

		if (ErrMsg = "") then
			Call GetSongjangSendResult()
		end if

    end function

	function ActSavexSiteCSOrderListtoDB()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo, OutMallRegDate
		dim OrgDetailKey, CSDetailKey, itemno, OutMallFinishDate, cnt, jupsucnt
		dim strSql, iAssignedRow

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML Replace(replace(objData,"&","＆"), "", "")

		If session("ssBctID")="kjy8517" Then
			response.write "<textarea cols=100 rows=30>"&replace(objData,"&","＆")&"</textarea>"
		End If

		if (FRectSellSite = "lotteCom") or (FRectSellSite = "lotteimall") then
			'// 롯데닷컴, 롯데i몰

			'// 마스터
			set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")

			masterCnt = objMasterListXML.length

			if (masterCnt > 0) then
				for i = 0 to masterCnt - 1
					set objMasterOneXML = objMasterListXML.item(i)

					divcd				= FRectDivCD
					gubunname			= objMasterOneXML.selectSingleNode("DelvInfo/ClmCausNm").text
					SellSite			= FRectSellSite
					OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
					OrderName			= objMasterOneXML.selectSingleNode("DelvInfo/OrdManNm").text
					OrderEmail			= ""
					OrderTelNo			= ""
					OrderHpNo			= ""
					ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/RmitNm").text
					ReceiveTelNo		= ""
					ReceiveHpNo			= ""
					ReceiveZipCode		= ""
					ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/Addr").text
					ReceiveAddr2		= ""
					deliverymemo		= ""

					if (FRectDivCD = "A008") then
						OutMallRegDate		= objMasterOneXML.selectSingleNode("DelvInfo/CnclDtime").text
					elseif (FRectDivCD = "A004_B007") then
						OutMallRegDate		= Left(now, 10)
						OutMallFinishDate	= objMasterOneXML.selectSingleNode("DelvInfo/SndFinDtime").text
						divcd 				= Left(divcd, 4)
					else
						OutMallRegDate		= Left(now, 10)
					end if

					'// 디테일
					set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
					detailCnt = objDetailListXML.length
					for j = 0 to detailCnt - 1
						set objDetailOneXML = objDetailListXML.item(j)

						OrgDetailKey	= objDetailOneXML.selectSingleNode("OrgOrdDtlSn").text
						CSDetailKey		= objDetailOneXML.selectSingleNode("OrdDtlSn").text

						if (FRectDivCD = "A008") then
							itemno			= objDetailOneXML.selectSingleNode("CnclQty").text
						else
							itemno			= objDetailOneXML.selectSingleNode("OrdQty").text
						end if

						strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
						strSql = strSql + " begin "
						strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
						strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
						strSql = strSql + " values('" + CStr(divcd) + "', '" + html2db(CStr(gubunname)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '" + html2db(CStr(OrderName)) + "', '" + html2db(CStr(OrderEmail)) + "', '" + html2db(CStr(OrderTelNo)) + "', '" + html2db(CStr(OrderHpNo)) + "', '" + html2db(CStr(ReceiveName)) + "', '" + html2db(CStr(ReceiveTelNo)) + "', '" + html2db(CStr(ReceiveHpNo)) + "', '" + html2db(CStr(ReceiveZipCode)) + "', '" + html2db(CStr(ReceiveAddr1)) + "', '" + html2db(CStr(ReceiveAddr2)) + "', '" + html2db(CStr(deliverymemo)) + "' "
						strSql = strSql + " , '" + html2db(CStr(OutMallRegDate)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(itemno) + ") "
						strSql = strSql + " end "
						''rw strSql
						dbget.execute strSql

						if (FRectDivCD = "A004_B007") then
							strSql = " update "
							strSql = strSql + " [db_temp].[dbo].[tbl_xSite_TMPCS] "
							strSql = strSql + " set OutMallFinishDate = '" & OutMallFinishDate & "', OutMallCurrState = 'B007' "
							strSql = strSql + " where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' and OutMallFinishDate is NULL and orderserial is not NULL "
							''rw strSql
							dbget.Execute strSql, iAssignedRow

							if (iAssignedRow > 0) then
								strSql = " select count(*) as cnt, IsNull(sum(case when a.currstate < 'B007' then 1 else 0 end), 0) as jupsucnt "
								strSql = strSql + " from "
								strSql = strSql + " 	[db_temp].[dbo].[tbl_xSite_TMPCS] c "
								strSql = strSql + " 	join db_cs.dbo.tbl_new_as_list a "
								strSql = strSql + " 	on "
								strSql = strSql + " 		1 = 1 "
								strSql = strSql + " 		and a.orderserial = c.OrderSerial "
								strSql = strSql + " 		and a.deleteyn = 'N' "
								strSql = strSql + " where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
								'response.write strSql & "<Br>"
								rsget.Open strSql,dbget,1

								if not rsget.EOF then
									cnt = rsget("cnt")
									jupsucnt = rsget("jupsucnt")
								end if

								rsget.Close

								if (cnt = 0) or (jupsucnt > 0) then
									'// CS완료이전 내역이 있으면 등록이전 상태로 전환
									strSql = " update "
									strSql = strSql + " [db_temp].[dbo].[tbl_xSite_TMPCS] "
									strSql = strSql + " set currstate = 'B001' "
									strSql = strSql + " where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "' "
									''rw strSql
									rsget.Open strSql, dbget, 1
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
						strSql = strSql + " 	and c.SellSite = o.SellSite "
						strSql = strSql + " 	and c.OutMallOrderSerial = o.OutMallOrderSerial "
						strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
						strSql = strSql + " where "
						strSql = strSql + " 	1 = 1 "
						strSql = strSql + " 	and c.orderserial is NULL "
						strSql = strSql + " 	and o.orderserial is not NULL "
						strSql = strSql + " 	and c.SellSite = '" + CStr(SellSite) + "' and c.OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and c.OrgDetailKey = '" + CStr(OrgDetailKey) + "' "
						''rw strSql
						dbget.execute strSql, iAssignedRow

						If iAssignedRow = 0 Then
							strSql = " update c "
							strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
							strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
							strSql = strSql + " from "
							strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
							strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
							strSql = strSql + " on "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.SellSite = o.SellSite "
							strSql = strSql + " 	and c.OutMallOrderSerial = o.ref_OutMallOrderSerial "
							strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
							strSql = strSql + " where "
							strSql = strSql + " 	1 = 1 "
							strSql = strSql + " 	and c.orderserial is NULL "
							strSql = strSql + " 	and o.orderserial is not NULL "
							strSql = strSql + " 	and c.SellSite = '" + CStr(SellSite) + "' and c.OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and c.OrgDetailKey = '" + CStr(OrgDetailKey) + "' "
							''rw strSql
							dbget.execute strSql
						End If

						set objDetailOneXML = Nothing
					next

					set objDetailListXML = Nothing
					set objMasterOneXML = Nothing
				next
			end if

			if Hour(Now()) = 9 then
				'// 주문입력 전에 CS내역이 먼저 오는 경우가 있음.
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " from "
				strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
				strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
				strSql = strSql + " on "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.SellSite = o.SellSite "
				strSql = strSql + "		and c.OutMallOrderSerial = o.OutMallOrderSerial "
				strSql = strSql + "		and c.OrgDetailKey = o.OrgDetailKey "
				strSql = strSql + " where "
				strSql = strSql + "		1 = 1 "
				strSql = strSql + "		and c.orderserial is NULL "
				strSql = strSql + "		and o.orderserial is not NULL "
				strSql = strSql + "		and c.SellSite = '" + CStr(SellSite) + "' "
				''rw strSql
				rsget.Open strSql, dbget, 1
			end if

			set objMasterListXML = Nothing

		else
			ErrMsg = "파싱에 실패했습니다."
		end if
		Set xmlDOM = Nothing
	end function

	function GetSongjangSendResult()
		dim i, j
		dim objMasterListXML, objMasterOneXML, objDetailListXML, objDetailOneXML
		dim masterCnt, detailCnt
		dim divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo, OutMallRegDate
		dim OrgDetailKey, CSDetailKey, itemno
		dim strSql
		dim IsSuccess

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","＆")

		ErrMsg = ""

		if (FRectSellSite = "lotteimall") then
			'// 롯데닷i몰

			IsSuccess = False
			set objMasterListXML = xmlDOM.selectNodes("/Response/Result")
			if (objMasterListXML.length > 0) then
				IsSuccess = True
			end if

			if IsSuccess then
				'// 성공
				strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
				strSql = strSql & " set sendstate=1"
				strSql = strSql & " ,sendreqCnt=IsNULL(sendreqCnt,0)+1"
				strSql = strSql & " where outmallorderserial='"&ord_no&"'"
				strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
				strSql = strSql & " and IsNULL(sendstate,0)=0"
				strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and ordercsgbn = 0"
				'rw strSql
				dbget.Execute strSql

				ErrMsg = "OK"
			else
				'// 에러
				set objMasterListXML = xmlDOM.selectNodes("/Response/Errors")
				set objMasterOneXML = objMasterListXML.item(0)

				ErrMsg = objMasterOneXML.selectSingleNode("Error/Message").text

				strSql = " update db_temp.dbo.tbl_xSite_TMPOrder"
				strSql = strSql & " set sendreqCnt=IsNULL(sendreqCnt,0)+1"
				strSql = strSql & " where outmallorderserial='"&ord_no&"'"
				strSql = strSql & " and orgdetailkey='"&ord_dtl_sn&"'"
				strSql = strSql & " and IsNULL(sendstate,0)=0"
				strSql = strSql & " and IsNULL(matchstate,'') <> 'D' and IsNULL(ordercsgbn, 0) = 0"
				''response.write strSql
				dbget.Execute strSql

				'// 에러 3회 이상이면 수기처리
				Dim errCount
				strSql = ""
				strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
				strSql = strSql & "	where OutMallOrderSerial='"&ord_no&"'"
				strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
				strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If Not rsget.Eof Then
					errCount = rsget("cnt")
				End If
				rsget.Close

				If errCount > 0 Then
					response.write  "<select name='updateSendState' id=""updateSendState"">" &_
									"	<option value=''>선택</option>" &_
									"	<option value='901'>발송처리누락 수기등록건</option>" &_
									"	<option value='902'>취소후 제결제건</option>" &_
									"	<option value='903'>반품처리건</option>" &_
									"</select>&nbsp;&nbsp;"
					response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)""><br>"
					response.write "<script language='javascript'>"&VbCRLF
					response.write "function finCancelOrd2(ord_no,ord_dtl_sn,selectValue){"&VbCRLF
					response.write "    if(selectValue == ''){"&VbCRLF
					response.write "    	alert('선택해주세요');"&VbCRLF
					response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
					response.write "    	return;"&VbCRLF
					response.write "    }"&VbCRLF
					response.write "    var uri = 'xSiteCSOrder_lotteimall_Process.asp?mode=updateSendState&ord_no='+ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
					response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
					response.write "    popwin.focus()"&VbCRLF
					response.write "}"&VbCRLF
					response.write "</script>"&VbCRLF
				End If
			end if

			'' '// 마스터
			'' set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")

			'' masterCnt = objMasterListXML.length

			'' if (masterCnt > 0) then
			'' 	for i = 0 to masterCnt - 1
			'' 		set objMasterOneXML = objMasterListXML.item(i)

			'' 		divcd				= FRectDivCD
			'' 		gubunname			= objMasterOneXML.selectSingleNode("DelvInfo/ClmCausNm").text
			'' 		SellSite			= FRectSellSite
			'' 		OutMallOrderSerial	= objMasterOneXML.selectSingleNode("OrdNo").text
			'' 		OrderName			= objMasterOneXML.selectSingleNode("DelvInfo/OrdManNm").text
			'' 		OrderEmail			= ""
			'' 		OrderTelNo			= ""
			'' 		OrderHpNo			= ""
			'' 		ReceiveName			= objMasterOneXML.selectSingleNode("DelvInfo/RmitNm").text
			'' 		ReceiveTelNo		= ""
			'' 		ReceiveHpNo			= ""
			'' 		ReceiveZipCode		= ""
			'' 		ReceiveAddr1		= objMasterOneXML.selectSingleNode("DelvInfo/Addr").text
			'' 		ReceiveAddr2		= ""
			'' 		deliverymemo		= ""

			'' 		if (FRectDivCD = "A008") then
			'' 			OutMallRegDate		= objMasterOneXML.selectSingleNode("DelvInfo/CnclDtime").text
			'' 		else
			'' 			OutMallRegDate		= Left(now, 10)
			'' 		end if

			'' 		'// 디테일
			'' 		set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
			'' 		detailCnt = objDetailListXML.length
			'' 		for j = 0 to detailCnt - 1
			'' 			set objDetailOneXML = objDetailListXML.item(j)

			'' 			OrgDetailKey	= objDetailOneXML.selectSingleNode("OrgOrdDtlSn").text
			'' 			CSDetailKey		= objDetailOneXML.selectSingleNode("OrdDtlSn").text

			'' 			if (FRectDivCD = "A008") then
			'' 				itemno			= objDetailOneXML.selectSingleNode("CnclQty").text
			'' 			else
			'' 				itemno			= objDetailOneXML.selectSingleNode("OrdQty").text
			'' 			end if

			'' 			strSql = " if not exists (select idx from db_temp.dbo.tbl_xSite_TMPCS where SellSite = '" + CStr(SellSite) + "' and OutMallOrderSerial = '" + CStr(OutMallOrderSerial) + "' and OrgDetailKey = '" + CStr(OrgDetailKey) + "' and CSDetailKey = '" + CStr(CSDetailKey) + "') "
			'' 			strSql = strSql + " begin "
			'' 			strSql = strSql + " insert into db_temp.dbo.tbl_xSite_TMPCS(divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			'' 			strSql = strSql + " , OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) "
			'' 			strSql = strSql + " values('" + CStr(divcd) + "', '" + html2db(CStr(gubunname)) + "', '" + html2db(CStr(SellSite)) + "', '" + html2db(CStr(OutMallOrderSerial)) + "', '" + html2db(CStr(OrderName)) + "', '" + html2db(CStr(OrderEmail)) + "', '" + html2db(CStr(OrderTelNo)) + "', '" + html2db(CStr(OrderHpNo)) + "', '" + html2db(CStr(ReceiveName)) + "', '" + html2db(CStr(ReceiveTelNo)) + "', '" + html2db(CStr(ReceiveHpNo)) + "', '" + html2db(CStr(ReceiveZipCode)) + "', '" + html2db(CStr(ReceiveAddr1)) + "', '" + html2db(CStr(ReceiveAddr2)) + "', '" + html2db(CStr(deliverymemo)) + "' "
			'' 			strSql = strSql + " , '" + html2db(CStr(OutMallRegDate)) + "', '" + html2db(CStr(OrgDetailKey)) + "', '" + html2db(CStr(CSDetailKey)) + "', " + CStr(itemno) + ") "
			'' 			strSql = strSql + " end "
			'' 			''rw strSql
			'' 			rsget.Open strSql, dbget, 1

			'' 			set objDetailOneXML = Nothing
			'' 		next

			'' 		set objDetailListXML = Nothing
			'' 		set objMasterOneXML = Nothing
			'' 	next
			'' end if

			'' set objMasterListXML = Nothing

			'' strSql = " update c "
			'' strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
			'' strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
			'' strSql = strSql + " from "
			'' strSql = strSql + " db_temp.dbo.tbl_xSite_TMPCS c "
			'' strSql = strSql + " join db_temp.dbo.tbl_xSite_TMPOrder o "
			'' strSql = strSql + " on "
			'' strSql = strSql + " 	1 = 1 "
			'' strSql = strSql + " 	and c.SellSite = o.SellSite "
			'' strSql = strSql + " 	and c.OutMallOrderSerial = Replace(o.OutMallOrderSerial, '-', '') "
			'' strSql = strSql + " 	and c.OrgDetailKey = o.OrgDetailKey "
			'' strSql = strSql + " where "
			'' strSql = strSql + " 	1 = 1 "
			'' strSql = strSql + " 	and c.orderserial is NULL "
			'' strSql = strSql + " 	and o.orderserial is not NULL "
			'' ''rw strSql
			'' rsget.Open strSql, dbget, 1
		else
			ErrMsg = "파싱에 실패했습니다."
		end if
		Set xmlDOM = Nothing
	end function

	'// 사용않함
	function GetxSiteCSOrderCount_XXX()
		dim objNode, objNodes

		FResultCount = 0

		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML replace(objData,"&","＆")

		if (FRectSellSite = "lotteCom") then
			'// 롯데닷컴
			FResultCount = xmlDOM.selectNodes("/Response/Result/OrderInfo").length
		else
			FResultCount = 0
		end if
		Set xmlDOM = Nothing
	end function

	public function GetXmlFromWeb()
		objData = ""
		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,60000,60000,90000  ''2013/08/01 추가
		objXML.Send()

		if objXML.Status = "200" then
			objData = BinaryToText(objXML.ResponseBody, "euc-kr")
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

		if (sellsite = "lotteCom") then
			if (FRectDivCD = "A008") then
				'// 취소
				GetXMLURL = lotteAPIURL + "/openapi/searchCnclList.lotte?subscriptionId=" + CStr(lotteAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD))
			elseif (FRectDivCD = "A004") then
				'// 반품
				GetXMLURL = lotteAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(lotteAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=20"
			elseif (FRectDivCD = "A004_B007") then
				'// 반품
				GetXMLURL = lotteAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(lotteAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=21"
			else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			end if
		elseif (sellsite = "lotteimall") then
			if (FRectDivCD = "sendsongjang") then
				'// 송장전송

				if application("Svr_Info")="Dev" then
					'// aaaaaaaaaaaaaaaaaaaaaaa
					'// 개발서버
					ltiMallAuthNo = "fe790295f0fec596ebc6a8a13a55513208e7830182501abf7b70d1fbc4e09ffde03afe430407f378238bccd00eda50718c4695904037247c5da9451d4f75dddc"
					ltiMallAPIURL = "http://openapi.lotteimall.com"
				end if

				'// sfin : 출고완료
				'// dfin : 배송완료(사용하지 말것)
				''GetXMLURL = ltiMallAPIURL + "/openapi/registDeliver.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&ord_no=" + CStr(ord_no) + "&ord_dtl_sn=" + CStr(ord_dtl_sn) + "&proc_gubun=dfin&hdc_cd=" + CStr(hdc_cd) + "&inv_no=" + CStr(inv_no) + "&dlv_fin_dtime=" + CStr(GetxSiteDateFormat(sendDate))
				GetXMLURL = ltiMallAPIURL + "/openapi/registDeliver.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&ord_no=" + CStr(ord_no) + "&ord_dtl_sn=" + CStr(ord_dtl_sn) + "&proc_gubun=sfin&hdc_cd=" + CStr(hdc_cd) + "&inv_no=" + CStr(inv_no) + "&dlv_fin_dtime=" + CStr(GetxSiteDateFormat(sendDate))
			elseif (FRectDivCD = "A008") then
				'// 취소
				GetXMLURL = ltiMallAPIURL + "/openapi/searchCnclList.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD))
			elseif (FRectDivCD = "A004") then
				'// 반품
				GetXMLURL = ltiMallAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=20"
			elseif (FRectDivCD = "A004_B007") then
				'// 반품
				GetXMLURL = ltiMallAPIURL + "/openapi/searchReturnList.lotte?subscriptionId=" + CStr(ltiMallAuthNo) + "&start_date=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "&end_date=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "&ord_dtl_stat_cd=21"
			else
				GetXMLURL = ""
				ErrMsg = "등록되지 않은 제휴몰입니다.[1]"
			end if
		else
			GetXMLURL = ""
			ErrMsg = "등록되지 않은 제휴몰입니다.[2]"
		end if
	end function

	public function GetxSiteDateFormat(dt)
		if (FRectSellSite = "lotteCom") then
			GetxSiteDateFormat = Replace(dt, "-", "")
		elseif (FRectSellSite = "lotteimall") then
			GetxSiteDateFormat = Replace(dt, "-", "")
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

%>
