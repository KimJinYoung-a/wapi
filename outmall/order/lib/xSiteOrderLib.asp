<%

Class OrderItem
	public FSellSite
	public FOutMallOrderSerial
End Class

class COrderMasterItem
	public FSellSite
	public FOutMallOrderSerial
	public FSellDate
	public FPayType
	public FPaydate
	public FOrderUserID
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
	public FdeliverPay

	public FUserID
	public ForderCsGbn
	public FcountryCode
	public Fshoplinkermallname
	public FshoplinkerOrderID
	public FshoplinkerMallID
	public FoverseasPrice
	public FoverseasDeliveryPrice
	public FoverseasRealPrice
	public Freserve01
	public FbeasongNum11st

	Private Sub Class_Initialize()
		ForderCsGbn = 0
		FcountryCode = "KR"
		''FoverseasPrice = 0
		''FoverseasDeliveryPrice = 0
		''FoverseasRealPrice = 0
	End Sub
end class

class COrderDetail
	public FdetailSeq
	public FItemID
	public FItemOption
	public FOutMallItemID
	public FOutMallItemName
	public FOutMallItemOption
	public FOutMallItemOptionName
	public Fitemcost
	public FReducedPrice
	public FItemNo
	public FOutMallCouponPrice
	public FTenCouponPrice
	public FrequireDetail

	public FshoplinkerPrdCode
end class

function GetOrderFromExtSite_example(example, selldate)
	dim orderObjArr(0)
	dim tmpItem

	'// 중간 처리과정
	'//  - 제휴몰에 접속
	'//  - 데이타 수신
	'//  - 클래스에 생성해서 주문정보 입력
	'//  - 배열에 클래스 객체 추가
	'// - 클래스배열 리턴

	set tmpItem = new OrderItem
	tmpItem.FSellSite = "example"
	tmpItem.FOutMallOrderSerial = "123412341234"
	set orderObjArr(0) = tmpItem

	GetOrderFromExtSite_example = orderObjArr
end function

function GetOrderFromExtSite(sellsite, selldate)
	select case sellsite
		case "interpark"
			Call GetOrderFrom_interpark(selldate)
		case "auction1010"
			Call GetOrderFrom_auction1010(selldate)
		case "gseshop"
			Call GetOrderFrom_gseshop(selldate)
		case "gseshopNew"
			Call GetOrderFrom_gseshopNew(selldate)
		case "sabangnet"
			Call GetOrderFrom_sabangnet(selldate)
		case "nvstorefarm", "nvstoremoonbangu", "nvstoregift", "Mylittlewhoopee"
			Call GetOrderFrom_nvstorefarm(sellsite, selldate)
		case "ezwel"
			Call GetOrderFrom_ezwel(selldate)
		case "lotteCom"
			Call GetOrderFrom_lotteCom(selldate)
		case "lotteon"
			Call GetOrderFrom_lotteon(selldate)
		case "shintvshopping"
			Call GetOrderFrom_shintvshopping(selldate)
		case "skstoa"
			Call GetOrderFrom_skstoa(selldate)
		case "lfmall"
			Call GetOrderFrom_lfmall(selldate)
		case "wetoo1300k"
			Call GetOrderFrom_wetoo1300k(selldate)
		case else
			response.write "잘못된 접근입니다."
		dbget.close : response.end
	end select
end function

function GetOrderFromExtSiteConfirmlist(sellsite, selldate)
	select case sellsite
		case "lotteCom"
			Call GetOrderFrom_lotteComConfirmList(selldate)
		case else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
	end select
end function

function GetOrderFrom_interpark(selldate)
	dim sellsite : sellsite = "interpark"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim sellFromDate

	GetOrderFrom_interpark = False

	sellFromDate = selldate
	if (sellFromDate = Left(Now(), 10)) then
		'// 오늘 신규주문 가져올 경우, 최근7일 주문내역에서 확인(주문일/결제일 상이한 케이스 등)
		sellFromDate = Left(DateAdd("d", -7, CDate(sellFromDate)), 10)
	end if

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-10-21"
	xmlSelldate = Replace(selldate, "-", "")
	sellFromDate = Replace(sellFromDate, "-", "")

	'// API URL(기간동안의 신규주문 가져오기)
	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=orderListForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + xmlSelldate + "000000" + "&sc.endDate=" + xmlSelldate + "235959"
	''주문확인 한 내역.. 만약 발송준비중 상태면 아래 주석풀고 실행..2021-08-02 김진영
	'''주문제작문구 긴 케이스가 있음.."OPT_NM" 292LINE 주석풀기

'	xmlURL = xmlURL + "/order/OrderClmAPI.do?_method=orderListDelvForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + xmlSelldate + "000000" + "&sc.endDate=" + xmlSelldate + "235959"
	''response.write xmlURL


	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status = "200" then
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
		''response.write objData
	else
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(objData,"&","＆")

	Set obj = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	if obj is Nothing then
		if IsAutoScript then
			''response.write "내역없음 : 종료"
		end if

		GetOrderFrom_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	masterCnt = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
	''response.write masterCnt

	if masterCnt = 0 then
		if IsAutoScript then
			''response.write "내역없음<br />"
		end if

		GetOrderFrom_interpark = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
	masterCnt = objMasterListXML.length

	''if IsAutoScript then
		response.write "건수(" & masterCnt & ") " & "<br />"
	''end if

	for i = 0 to masterCnt - 1
		set objMasterOneXML = objMasterListXML.item(i)
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite = sellsite
		oMaster.FOutMallOrderSerial = objMasterOneXML.selectSingleNode("ORD_NO").text
		oMaster.FSellDate			= objMasterOneXML.selectSingleNode("ORDER_DT").text
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= Left(objMasterOneXML.selectSingleNode("PAY_DTS").text,8)
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= Left(html2db(objMasterOneXML.selectSingleNode("ORD_NM").text), 28)
		oMaster.FOrderEmail			= ""
		oMaster.FOrderTelNo			= objMasterOneXML.selectSingleNode("TEL").text
		oMaster.FOrderHpNo			= objMasterOneXML.selectSingleNode("MOBILE_TEL").text
		oMaster.FReceiveName		= Left(html2db(objMasterOneXML.selectSingleNode("RCVR_NM").text), 28)
		oMaster.FReceiveTelNo		= objMasterOneXML.selectSingleNode("DELI_TEL").text
		oMaster.FReceiveHpNo		= objMasterOneXML.selectSingleNode("DELI_MOBILE").text
		oMaster.FReceiveZipCode		= objMasterOneXML.selectSingleNode("DEL_ZIP").text
		oMaster.FReceiveAddr1		= html2db(objMasterOneXML.selectSingleNode("DELI_ADDR1").text)
		oMaster.FReceiveAddr2		= html2db(objMasterOneXML.selectSingleNode("DELI_ADDR2").text)
		oMaster.FdeliverPay			= objMasterOneXML.selectNodes("DELIVERY/DELV").item(0).selectSingleNode("DEL_AMT").text
		If oMaster.FOutMallOrderSerial = "20221228130606838259" Then
			oMaster.Fdeliverymemo		= "축하카드문구 여보! 생일축하해! 항상 고맙고 또 사랑해♡ 수령요청일1월2일"
		Else
			'oMaster.Fdeliverymemo		= html2db(objMasterOneXML.selectSingleNode("DELI_COMMENT").text)
			oMaster.Fdeliverymemo		= objMasterOneXML.selectSingleNode("DELI_COMMENT").text	'2021-12-20 김진영 html2db 삭제.
		End If

		'// 날짜 수정
		oMaster.FSellDate = Left(oMaster.FSellDate,4) & "-" & Mid(oMaster.FSellDate,5,2) & "-" & Mid(oMaster.FSellDate,7,2)
		oMaster.FPaydate = Left(oMaster.FPaydate,4) & "-" & Mid(oMaster.FPaydate,5,2) & "-" & Mid(oMaster.FPaydate,7,2)

		'// 우편번호 수정
		if Len(oMaster.FReceiveZipCode) = 4 then
			oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
		end if
		if Len(oMaster.FReceiveZipCode) > 4 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		set objDetailListXML = objMasterOneXML.selectNodes("PRODUCT/PRD")
		detailCnt = objDetailListXML.length
		redim oDetailArr(detailCnt - 1)
		For j = 0 to detailCnt - 1
			Set objDetailOneXML = objDetailListXML.item(j)
			Set oDetailArr(j) = new COrderDetail

			oDetailArr(j).FdetailSeq = objDetailOneXML.selectSingleNode("ORD_SEQ").text
			oDetailArr(j).FItemID = objDetailOneXML.selectSingleNode("ENTR_PRD_NO").text
			oDetailArr(j).FItemOption = objDetailOneXML.selectSingleNode("OPT_NO").text
			oDetailArr(j).FOutMallItemID = objDetailOneXML.selectSingleNode("PRD_NO").text
			oDetailArr(j).FOutMallItemOption = objDetailOneXML.selectSingleNode("OPT_PRD_NO").text
			oDetailArr(j).FOutMallItemName = html2db(objDetailOneXML.selectSingleNode("PRD_NM").text)
			oDetailArr(j).FOutMallItemOptionName = html2db(objDetailOneXML.selectSingleNode("OPT_NM").text)
			'주문제작문구 긴 케이스가 있음..그럴 때 하단 주석풀자
			'oDetailArr(j).FOutMallItemOptionName = LEFT(html2db(objDetailOneXML.selectSingleNode("OPT_NM").text), 50)

			'2018-09-13 10:52 김진영 주석처리..REAL_SALE_UNITCOST(실상품판매총액) 가 새로 생긴듯.
			'oDetailArr(j).FReducedPrice = objDetailOneXML.selectSingleNode("SALE_UNITCOST").text
			oDetailArr(j).FReducedPrice = objDetailOneXML.selectSingleNode("REAL_SALE_UNITCOST").text
			oDetailArr(j).FOutMallCouponPrice = objDetailOneXML.selectSingleNode("DC_COUPON_AMT").text
			oDetailArr(j).FTenCouponPrice = objDetailOneXML.selectSingleNode("ENTR_DC_COUPON_AMT").text
			'2018-09-13 10:52 김진영 주석처리..REAL_SALE_UNITCOST(실상품판매총액) 가 새로 생김에 따라 itemcost구하는 것도 변경
			'2018-09-19 16:51 김진영 추가..itemcost에 하나를 더 더해야함..포인트즉시할인액(PRE_USE_UNITCOST) 이란 게 또 있음
			'oDetailArr(j).Fitemcost = CLng(oDetailArr(j).FReducedPrice) + CLng(oDetailArr(j).FOutMallCouponPrice) + CLng(oDetailArr(j).FTenCouponPrice)
			'oDetailArr(j).Fitemcost = CLng(oDetailArr(j).FReducedPrice) + CLng(oDetailArr(j).FOutMallCouponPrice) + CLng(oDetailArr(j).FTenCouponPrice) + CLng(objDetailOneXML.selectSingleNode("IPOINT_DC_UNITCOST").text)
			oDetailArr(j).Fitemcost = CLng(oDetailArr(j).FReducedPrice) + CLng(oDetailArr(j).FOutMallCouponPrice) + CLng(oDetailArr(j).FTenCouponPrice) + CLng(objDetailOneXML.selectSingleNode("IPOINT_DC_UNITCOST").text) + CLng(objDetailOneXML.selectSingleNode("PRE_USE_UNITCOST").text)

			oDetailArr(j).FItemNo = objDetailOneXML.selectSingleNode("ORD_QTY").text

			if (oDetailArr(j).FItemID = oDetailArr(j).FItemOption) then
				oDetailArr(j).FItemOption = "0000"
			end if

			if IsNull(oDetailArr(j).FItemOption) then
				oDetailArr(j).FItemOption = ""
			end if

			if (oDetailArr(j).FItemOption = "") then
				oDetailArr(j).FItemOption = "0000"
			end if

			'// 주문제작문구 수정
			if InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구 / ") <> 0 then
				oDetailArr(j).FrequireDetail = Mid(oDetailArr(j).FOutMallItemOptionName, InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구 / ") + Len("주문제작문구 / "), 1000)
			elseif InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구/") <> 0 then
				oDetailArr(j).FrequireDetail = Mid(oDetailArr(j).FOutMallItemOptionName, InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구/") + Len("주문제작문구/"), 1000)
			elseif InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구 | ") <> 0 then
				oDetailArr(j).FrequireDetail = Mid(oDetailArr(j).FOutMallItemOptionName, InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구 | ") + Len("주문제작문구 | "), 1000)
			elseif InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구") <> 0 then
				oDetailArr(j).FrequireDetail = Mid(oDetailArr(j).FOutMallItemOptionName, InStr(oDetailArr(j).FOutMallItemOptionName, "주문제작문구") + Len("주문제작문구"), 1000)
			end if
		next

		if (SaveOrderToDB(oMaster, oDetailArr) = True) then
			successCnt = successCnt + 1
		end if

		Set oMaster = Nothing
	next

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")"
	''end if

	Set xmlDOM = Nothing
	Set objXML = Nothing

	GetOrderFrom_interpark = True
end function

function SaveOrderToDB(oMaster, oDetailArr)
	dim sqlStr
	dim i, j, k
	dim paramInfo, retParamInfo, RetErr, retErrStr
	dim orderDlvPay
	dim tmpStr

	SaveOrderToDB = False

	if NOT isNULL(oMaster.FReceiveZipCode) then
		if (LEN(replace(Trim(oMaster.FReceiveZipCode),"-",""))=5) then  ''5자리 우편번호이면
			oMaster.FReceiveZipCode = replace(Trim(oMaster.FReceiveZipCode),"-","")
		end if
	end if

	for i = 0 to UBound(oDetailArr)
		if (i = 0) then
			orderDlvPay = oMaster.FdeliverPay
		else
			orderDlvPay = 0
		end if

		tmpStr = " exec db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_TEST "
		tmpStr = tmpStr + "'" & oMaster.FSellSite & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOutMallOrderSerial & "'"
		tmpStr = tmpStr + ", '" & oMaster.FSellDate & "'"
		tmpStr = tmpStr + ", '" & oMaster.FPayType & "'"
		tmpStr = tmpStr + ", '" & oMaster.FPaydate & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemID & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemOption & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemID & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemName & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemOption & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemOptionName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FUserID & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderEmail & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderTelNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderHpNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveName & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveTelNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FOrderHpNo & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveZipCode & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveAddr1 & "'"
		tmpStr = tmpStr + ", '" & oMaster.FReceiveAddr2 & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).Fitemcost & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FReducedPrice & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FItemNo & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FdetailSeq & "'"
		tmpStr = tmpStr + ", '" & 0 & "'"
		tmpStr = tmpStr + ", '" & 0 & "'"
		tmpStr = tmpStr + ", '" & oMaster.Fdeliverymemo & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FrequireDetail & "'"
		tmpStr = tmpStr + ", '" & orderDlvPay & "'"
		tmpStr = tmpStr + ", '" & oMaster.ForderCsGbn & "'"
		tmpStr = tmpStr + ", '" & oMaster.FcountryCode & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FOutMallItemID & "'"
		tmpStr = tmpStr + ", '" & oMaster.Fshoplinkermallname & "'"
		tmpStr = tmpStr + ", '" & oDetailArr(i).FshoplinkerPrdCode & "'"
		tmpStr = tmpStr + ", '" & oMaster.FshoplinkerOrderID & "'"
		tmpStr = tmpStr + ", '" & oMaster.FshoplinkerMallID & "'"
		tmpStr = tmpStr + ", ''"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasDeliveryPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.FoverseasRealPrice & "'"
		tmpStr = tmpStr + ", '" & oMaster.Freserve01 & "'"
		tmpStr = tmpStr + ", '" & oMaster.FbeasongNum11st & "'"

		tmpStr = Replace(tmpStr, "'", "^")
		sqlStr = "insert into db_temp.dbo.tbl_tmp_gsOrder"
		sqlStr = sqlStr&" (regdate,refip,xmlData)"
		sqlStr = sqlStr&" values(getdate(),'XXX','KKK-" & tmpStr & "')"
		''dbget.Execute sqlStr

		paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
        	,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, Trim(oMaster.FSellSite))	_
			,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(oMaster.FOutMallOrderSerial)) _
			,Array("@SellDate"				, adDate		, adParamInput		,	  , Trim(oMaster.FSellDate)) _
			,Array("@PayType"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FPayType)) _
			,Array("@Paydate"				, adDate		, adParamInput		,     , Trim(oMaster.FPaydate)) _
			,Array("@matchItemID"			, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemID)) _
			,Array("@matchItemOption"		, adVarchar		, adParamInput		,    4, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerItemID"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FItemID)) _
			,Array("@partnerItemName"		, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FOutMallItemName)) _
			,Array("@partnerOption"			, adVarchar		, adParamInput		,  128, Trim(oDetailArr(i).FItemOption)) _
			,Array("@partnerOptionName"		, adVarchar		, adParamInput		, 1024, Trim(oDetailArr(i).FOutMallItemOptionName)) _
			,Array("@OrderUserID"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FUserID)) _
			,Array("@OrderName"				, adVarchar		, adParamInput		,   32, Trim(oMaster.FOrderName)) _
			,Array("@OrderEmail"			, adVarchar		, adParamInput		,  100, Trim(oMaster.FOrderEmail)) _
			,Array("@OrderTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderTelNo)) _
			,Array("@OrderHpNo"				, adVarchar		, adParamInput		,   16, Trim(oMaster.FOrderHpNo)) _
			,Array("@ReceiveName"			, adVarchar		, adParamInput		,   32, Trim(oMaster.FReceiveName)) _
			,Array("@ReceiveTelNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveTelNo)) _
			,Array("@ReceiveHpNo"			, adVarchar		, adParamInput		,   16, Trim(oMaster.FReceiveHpNo)) _
			,Array("@ReceiveZipCode"		, adVarchar		, adParamInput		,   20, Trim(oMaster.FReceiveZipCode)) _
			,Array("@ReceiveAddr1"			, adVarchar		, adParamInput		,  128, Trim(oMaster.FReceiveAddr1)) _
			,Array("@ReceiveAddr2"			, adVarchar		, adParamInput		,  512, Trim(oMaster.FReceiveAddr2)) _
			,Array("@SellPrice"				, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).Fitemcost)) _
			,Array("@RealSellPrice"			, adCurrency	, adParamInput		,     , Trim(oDetailArr(i).FReducedPrice)) _
			,Array("@ItemOrderCount"		, adInteger		, adParamInput		,     , Trim(oDetailArr(i).FItemNo)) _
			,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FdetailSeq)) _
			,Array("@DeliveryType"			, adInteger		, adParamInput		,     , 0) _
			,Array("@deliveryprice"			, adCurrency	, adParamInput		,     , 0) _
			,Array("@deliverymemo"			, adVarchar		, adParamInput		,  400, Trim(oMaster.Fdeliverymemo)) _
			,Array("@requireDetail"			, adVarchar		, adParamInput		, 400, Trim(oDetailArr(i).FrequireDetail)) _
			,Array("@orderDlvPay"			, adCurrency	, adParamInput		,     , orderDlvPay) _
			,Array("@orderCsGbn"			, adInteger		, adParamInput		,     , oMaster.ForderCsGbn) _
			,Array("@countryCode"			, adVarchar		, adParamInput		,    2, oMaster.FcountryCode) _
            ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(oDetailArr(i).FOutMallItemID)) _
			,Array("@shoplinkerMallName" 	, adVarchar		, adParamInput		,   64, oMaster.Fshoplinkermallname) _
			,Array("@shoplinkerPrdCode"		, adVarchar		, adParamInput		,   16, oDetailArr(i).FshoplinkerPrdCode) _
			,Array("@shoplinkerOrderID"		, adVarchar		, adParamInput		,   16, oMaster.FshoplinkerOrderID) _
			,Array("@shoplinkerMallID"		, adVarchar		, adParamInput		,   32, oMaster.FshoplinkerMallID) _
			,Array("@retErrStr"				, adVarchar		, adParamOutput		,  100, "") _
			,Array("@overseasPrice"			, adCurrency	, adParamInput		,     , oMaster.FoverseasPrice) _
			,Array("@overseasDeliveryPrice"	, adCurrency	, adParamInput		,     , oMaster.FoverseasDeliveryPrice) _
			,Array("@overseasRealPrice"		, adCurrency	, adParamInput		,     , oMaster.FoverseasRealPrice) _
			,Array("@reserve01"				, adVarchar		, adParamInput		,   32, oMaster.Freserve01) _
			,Array("@beasongNum11st"		, adVarchar		, adParamInput		,   16, oMaster.FbeasongNum11st) _
			,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(oDetailArr(i).FOutMallItemOption)) _
    	)

		if (IS_TEST_MODE = True) then
			sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_TEST"
'			response.write oMaster.FSellSite & "<br />"
'			response.write oMaster.FOutMallOrderSerial & "<br />"
'			response.write oMaster.FSellDate & "<br />"
'			response.write oMaster.FPayType & "<br />"
'			response.write oMaster.FPaydate & "<br />"
'			response.write oDetailArr(i).FItemID & "<br />"
'			response.write oDetailArr(i).FItemOption & "<br />"
'			response.write oDetailArr(i).FItemID & "<br />"
'			response.write oDetailArr(i).FOutMallItemName & "<br />"
'			response.write oDetailArr(i).FItemOption & "<br />"
'			response.write oDetailArr(i).FOutMallItemOptionName & "<br />"
'			response.write oMaster.FUserID & "<br />"
'			response.write oMaster.FOrderName & "<br />"
'			response.write oMaster.FOrderEmail & "<br />"
'			response.write oMaster.FOrderTelNo & "<br />"
'			response.write oMaster.FOrderHpNo & "<br />"
'			response.write oMaster.FReceiveName & "<br />"
'			response.write oMaster.FReceiveTelNo & "<br />"
'			response.write oMaster.FReceiveHpNo & "<br />"
'			response.write oMaster.FReceiveZipCode & "<br />"
'			response.write oMaster.FReceiveAddr1 & "<br />"
'			response.write oMaster.FReceiveAddr2 & "<br />"
'			response.write oDetailArr(i).Fitemcost & "<br />"
'			response.write oDetailArr(i).FReducedPrice & "<br />"
'			response.write oDetailArr(i).FItemNo & "<br />"
'			response.write oDetailArr(i).FdetailSeq & "<br />"
'			response.write oMaster.Fdeliverymemo & "<br />"
'			response.write oDetailArr(i).FrequireDetail & "<br />"
'			response.write oMaster.FdeliverPay & "<br />"
'			response.write oMaster.ForderCsGbn & "<br />"
'			response.write oMaster.FcountryCode & "<br />"
'			response.write oDetailArr(i).FOutMallItemID & "<br />"
'			response.write oMaster.Fshoplinkermallname & "<br />"
'			response.write oDetailArr(i).FshoplinkerPrdCode & "<br />"
'			response.write oMaster.FshoplinkerOrderID & "<br />"
'			response.write oMaster.FshoplinkerMallID & "<br />"
'			response.write oMaster.FoverseasPrice & "<br />"
'			response.write oMaster.FoverseasDeliveryPrice & "<br />"
'			response.write oMaster.FoverseasRealPrice & "<br />"
'			response.write oMaster.Freserve01 & "<br />"
'			response.write oMaster.FbeasongNum11st & "<br />"

			''dbget.rollbackTrans
			''dbget.close() : response.end
		else
			sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
		end if

			' If session("ssBctID")="kjy8517" Then
			' 	On Error Resume Next
			' 	dbget.BeginTrans
			' 	retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
			' 	If Err.Number <> 0 Then
			' 		tmpStr = " exec db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert "
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FSellSite) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOutMallOrderSerial) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FSellDate) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FPayType) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FPaydate) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemOption) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemOption) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemOptionName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FUserID) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderEmail) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderTelNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FOrderHpNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveName) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveTelNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveHpNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveZipCode) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveAddr1) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.FReceiveAddr2) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).Fitemcost) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FReducedPrice) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FItemNo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FdetailSeq) & "',"
			' 		tmpStr = tmpStr + "'0',"
			' 		tmpStr = tmpStr + "'0',"
			' 		tmpStr = tmpStr + "'" & Trim(oMaster.Fdeliverymemo) & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FrequireDetail) & "',"
			' 		tmpStr = tmpStr + "'" & orderDlvPay & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.ForderCsGbn & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FcountryCode & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemID) & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.Fshoplinkermallname & "',"
			' 		tmpStr = tmpStr + "'" & oDetailArr(i).FshoplinkerPrdCode & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FshoplinkerOrderID & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FshoplinkerMallID & "',"
			' 		tmpStr = tmpStr + "'',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasDeliveryPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FoverseasRealPrice & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.Freserve01 & "',"
			' 		tmpStr = tmpStr + "'" & oMaster.FbeasongNum11st & "',"
			' 		tmpStr = tmpStr + "'" & Trim(oDetailArr(i).FOutMallItemOption) & "'"
			' 		rw tmpStr
			' 		rw "-----------------------------"
			' 	End If


			' 	RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
			' 	retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

			' 	if (RetErr<0) and (RetErr<>-1) then ''Break
			' 		'// 에러코드 -1 은 중복입력
			' 		dbget.rollbackTrans
			' 		if IsAutoScript then
			' 			response.write "ERROR["&retErr&"]"& retErrStr
			' 		else
			' 			response.write "ERROR["&retErr&"]"& retErrStr
			' 			response.write "<script>alert('오류가 발생했습니다.');</script>"
			' 		end if

			' 		dbget.close() : response.end
			' 	elseif (RetErr <> -1) then
			' 		SaveOrderToDB = True
			' 	end if

			' 	dbget.CommitTrans
			' 	On Error Goto 0
			' Else
			' 		dbget.BeginTrans

			' 		retParamInfo = fnExecSPOutput(sqlStr, paramInfo)

			' 		RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
			' 		retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

			' 		if (RetErr<0) and (RetErr<>-1) then ''Break
			' 			'// 에러코드 -1 은 중복입력
			' 			dbget.rollbackTrans
			' 			if IsAutoScript then
			' 				response.write "ERROR["&retErr&"]"& retErrStr
			' 			else
			' 				response.write "ERROR["&retErr&"]"& retErrStr
			' 				response.write "<script>alert('오류가 발생했습니다.');</script>"
			' 			end if

			' 			dbget.close() : response.end
			' 		elseif (RetErr <> -1) then
			' 			SaveOrderToDB = True
			' 		end if

			' 		dbget.CommitTrans
			' End If

		dbget.BeginTrans

		retParamInfo = fnExecSPOutput(sqlStr, paramInfo)

		RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
		retErrStr  = GetValue(retParamInfo, "@retErrStr") ' 오류명

		if (RetErr<0) and (RetErr<>-1) then ''Break
			'// 에러코드 -1 은 중복입력
			dbget.rollbackTrans
			if IsAutoScript then
				response.write "ERROR["&retErr&"]"& retErrStr
			else
				response.write "ERROR["&retErr&"]"& retErrStr
				response.write "<script>alert('오류가 발생했습니다.');</script>"
			end if

			dbget.close() : response.end
		elseif (RetErr <> -1) then
			SaveOrderToDB = True
		end if

		dbget.CommitTrans
	next
end function

function GetOrderFrom_lotteCom(selldate)
	dim sellsite : sellsite = "lotteCom"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim tmpStr, pos
	dim i, j, k
	dim found, successCnt

	GetOrderFrom_lotteCom = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 전체내역 가져오기)
	xmlURL = "https://openapi.lotte.com"
	xmlURL = xmlURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + GetLotteAuthNo() + "&start_date=" + xmlSelldate + "&end_date=" + xmlSelldate + "&SelOption=01"

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	else
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML (objData)

	if xmlDOM.getElementsByTagName("Response/Result/OrderInfo").length < 1 then
		''if IsAutoScript then
			response.write "내역없음(0)<br />"
		''end if

		GetOrderFrom_lotteCom = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	masterCnt = xmlDOM.getElementsByTagName("Response/Result/OrderInfo").length
	''if IsAutoScript then
		response.write "건수(" & masterCnt & ") " & "<br />"
	''end if

	set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
	for each objMasterOneXML in objMasterListXML
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite 			= sellsite
		oMaster.FOutMallOrderSerial = Replace(objMasterOneXML.selectSingleNode("OrdNo").text, "-", "")
		oMaster.FSellDate 			= Left(Now(), 10)
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= oMaster.FSellDate
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("CardMemoSndrName").text))
		oMaster.FOrderTelNo			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("OrderTelNo").text))
		oMaster.FOrderHpNo			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("OrderHpNo").text))
		oMaster.FOrderEmail			= ""
		oMaster.FReceiveName		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvName").text))
		if Len(oMaster.FReceiveName) > 32 then
			oMaster.FReceiveName = oMaster.FOrderName
		end if

		oMaster.FReceiveTelNo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text))
		oMaster.FReceiveHpNo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text))

		oMaster.Fdeliverymemo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DlvMemoCont").text))
		if (oMaster.Fdeliverymemo = "null") then
			oMaster.Fdeliverymemo = ""
		end if

		'// 배송비 안넘어옴
		oMaster.FdeliverPay			= 0

		oMaster.FReceiveZipCode		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text))

		'// 우편번호 수정
		if Len(oMaster.FReceiveZipCode) = 4 then
			oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
		end if

		oMaster.FReceiveAddr1 = html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text))
		oMaster.FReceiveAddr2 = html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text))

		if InStr(oMaster.FReceiveZipCode, "-") = 0 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
		oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)


		'// FROM
		'// 견종선택:푸들,사이즈:S (23-30cm),끈타입:타입3 (믹스)
		'// TO
		'// 푸들,S (23-30cm),타입3 (믹스)
		dim regEx
		set regEx = New RegExp
		With regEx
			.Pattern = ",[^:]+:"
			.IgnoreCase = True
			.Global = True
		end with


		set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
		for each objDetailOneXML in objDetailListXML
			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objDetailOneXML.selectSingleNode("ProdSeq").text
			oDetailArr(0).FItemID = ""
			oDetailArr(0).FItemOption = "0000"
			oDetailArr(0).FOutMallItemID = objDetailOneXML.selectSingleNode("ProdCode").text
			oDetailArr(0).FOutMallItemOption = "0000"
			oDetailArr(0).FOutMallItemName = html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("ProdName").text))
			oDetailArr(0).FOutMallItemOptionName = html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("prodOption").text))
			if (oDetailArr(0).FOutMallItemOptionName = "null") then
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			'// 롯데닷컴은 업체상품코드/옵션코드 모두 안준다.
			found = False
			sqlStr = ""
			sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
			sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddOption_regItem as r on m.idx = r.midx "
			sqlStr = sqlStr & " WHERE IsNULL(r.LotteGoodNo, r.LotteTmpGoodNo)= '"& oDetailArr(0).FOutMallItemID &"' "
			sqlStr = sqlStr & " and m.mallid = 'lotteCom' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If (Not rsget.EOF) Then
				found = True
				oDetailArr(0).FItemID = rsget("itemid")
				oDetailArr(0).FItemOption = rsget("itemoption")
				oDetailArr(0).FOutMallItemOption = rsget("itemoption")
			End If
			rsget.Close

			if found = False then
		        sqlStr = " select top 2 itemid from db_item.dbo.tbl_lotte_regItem "
		        sqlStr = sqlStr & " where IsNULL(LotteGoodNo,LotteTmpGoodNo)='"& oDetailArr(0).FOutMallItemID &"'"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

				if (rsget.RecordCount = 1) and (Not rsget.EOF) then
					oDetailArr(0).FItemID = rsget("itemid")
				elseif (oDetailArr(0).FOutMallItemID = "19710092") then
					oDetailArr(0).FItemID = 481915
				end if
				rsget.Close

				if oDetailArr(0).FItemID <> "" then
					if oDetailArr(0).FOutMallItemOptionName <> "" then
						oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, mid(regEx.replace("," & oDetailArr(0).FOutMallItemOptionName, ","), 2, 1000))
					else
						oDetailArr(0).FItemOption = "0000"
					end if
					oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
				else
					'2018-10-02 15:43 김진영 추가..TimeOut등으로 lotte_regItem테이블에 값을 못 넣는 경우 발생
					'oDetailArr(0).FItemID = -1
					sqlStr = " select top 1 itemid from db_item.dbo.tbl_item "
					sqlStr = sqlStr & " where itemname = '"& html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("ProdName").text)) &"'"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					If (Not rsget.EOF) Then
						oDetailArr(0).FItemID = rsget("itemid")
					End If
					rsget.Close

					if oDetailArr(0).FItemID <> "" then
						if oDetailArr(0).FOutMallItemOptionName <> "" then
							oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, mid(regEx.replace("," & oDetailArr(0).FOutMallItemOptionName, ","), 2, 1000))
						else
							oDetailArr(0).FItemOption = "0000"
						end if
						oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
					Else
						oDetailArr(0).FItemID = -1
					End If
				end if
			end if

			oDetailArr(0).FItemNo = CLng(objDetailOneXML.selectSingleNode("ordQty").text)

			oDetailArr(0).Fitemcost = objDetailOneXML.selectSingleNode("ordPrice").text
			oDetailArr(0).FReducedPrice = objDetailOneXML.selectSingleNode("ordPrice").text
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0

			oDetailArr(0).FrequireDetail = objDetailOneXML.selectSingleNode("GoodsChocDesc").text
			if (oDetailArr(0).FrequireDetail = "null") then
				oDetailArr(0).FrequireDetail = ""
			end if

			oMaster.ForderCsGbn = "0"
			if (objDetailOneXML.selectSingleNode("Exchange").text <> "일반") then
				oMaster.ForderCsGbn = "3"
			end if

			if (SaveOrderToDB(oMaster, oDetailArr) = True) then
					successCnt = successCnt + 1
			end if

			Set oDetailArr = Nothing
		next
		Set oMaster = Nothing
	next

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

Function GetOrderFrom_skstoa_Gubun1(selldate)
	dim sellsite : sellsite = "skstoa"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0
	Dim orderGb, orderNo, orderGSeq, orderDSeq, orderWSeq, goodsCode, goodsdtCode
	GetOrderFrom_skstoa_Gubun1 = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode			'#연결코드 | SKB에서 부여한 연결코드
	addParam = addParam & "&entpCode=" & skstoaentpCode			'#업체코드 | SKB에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & skstoaentpId				'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & skstoaentpPass			'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&doFlag=25"							'진행단계 | 25:출하지시대상(default), 30:출고대상, 40:배송완료대상
	addParam = addParam & "&orderGb=10"							'배송 구분 값 | 00:전체(default), 10:주문, 40:교환
	addParam = addParam & "&bDate="& xmlSelldate				'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate				'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="							'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/delivery/order-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				If returnCode = "200" Then
					Set orderList = strObj.orderList
						If orderList.length > 0 Then
							For i=0 to orderList.length-1
								orderGb = orderList.get(i).orderGb
								orderNo = orderList.get(i).orderNo
								orderGSeq = orderList.get(i).orderGSeq
								orderDSeq = orderList.get(i).orderDSeq
								orderWSeq = orderList.get(i).orderWSeq
								goodsCode = orderList.get(i).goodsCode
								goodsdtCode = orderList.get(i).goodsdtCode

								strSql = ""
								strSql = strSql & " IF NOT EXISTS (SELECT idx FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] "
								strSql = strSql & " WHERE sellsite = '" & sellsite & "' "
								strSql = strSql & " and OutMallOrderSerial = '" & orderNo & "' "
								strSql = strSql & " and orderGSeq = '" & orderGSeq & "' "
								strSql = strSql & " and orderDSeq = '" & orderDSeq & "' "
								strSql = strSql & " and orderWSeq = '" & orderWSeq & "' "
								strSql = strSql & " and shintvshoppingGoodNo = '"& goodsCode &"' "
								strSql = strSql & " and outmallOptCode = '" & goodsdtCode & "' "
								strSql = strSql & " ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] ([sellsite], [orderGb], [outmallorderserial], [orderGSeq], [orderDSeq], [orderWSeq], [shintvshoppingGoodNo], [outmallOptCode], [regdate]) "
								strSql = strSql & " 	VALUES ('"& sellsite &"', '"& orderGb &"', '"& orderNo &"', '"& orderGSeq &"', '"& orderDSeq &"', '"& orderWSeq &"', '"& goodsCode &"', '"& goodsdtCode &"', GETDATE()) "
								strSql = strSql & " END "
								dbget.Execute strSql
							Next
						End If
					Set orderList = nothing
				Else
					response.write "내역없음(0)<br />"
					GetOrderFrom_skstoa_Gubun1 = True
					Set strObj = Nothing
					rw "------------"
					Exit Function
				End If
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
			GetOrderFrom_skstoa_Gubun1 = True
			Set strObj = Nothing
			rw "------------"
			Exit Function
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_skstoa_Gubun2(outmallorderserial, orderGSeq, orderDSeq, orderWSeq, skstoaGoodNo, outmallOptCode)
	dim sellsite : sellsite = "skstoa"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0

	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode			'#연결코드 | SKB에서 부여한 연결코드
	addParam = addParam & "&entpCode=" & skstoaentpCode			'#업체코드 | SKB에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & skstoaentpId				'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & skstoaentpPass			'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&orderGb=10"							'#배송 구분 값 | 10:주문, 40:교환
	addParam = addParam & "&orderNo=" & outmallorderserial		'#주문번호
	addParam = addParam & "&orderGSeq=" & orderGSeq				'#상품순번 | 주문번호 G시퀀스 코드
	addParam = addParam & "&orderDSeq=" & orderDSeq				'#세트순번 | 주문번호 D시퀀스 코드
	addParam = addParam & "&orderWSeq=" & orderWSeq				'#처리순번 | 주문번호 W시퀀스 코드
	addParam = addParam & "&goodsCode=" & skstoaGoodNo			'#판매상품코드
	addParam = addParam & "&goodsdtCode=" & outmallOptCode		'#판매단품코드

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", skstoaAPIURL & "/partner/delivery/delivery-order", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(addParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					rw "orderNo : " & outmallorderserial & " |||| goodsCode : " & skstoaGoodNo & " [Deliver Ready Complete] "
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_skstoa(selldate)
	dim sellsite : sellsite = "skstoa"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0
	Dim orderGb, orderNo, orderGSeq, orderDSeq, orderWSeq, goodsCode, goodsdtCode
	dim oMaster, oDetail, oDetailArr, tmpItemid, tmpItemOption
	dim tmpOptionSeq : tmpOptionSeq = 0
	Dim POS1, POS2, POS3
	GetOrderFrom_skstoa = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & skstoalinkCode			'#연결코드 | SKB에서 부여한 연결코드
	addParam = addParam & "&entpCode=" & skstoaentpCode			'#업체코드 | SKB에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & skstoaentpId				'#업체사용자ID | SKB에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & skstoaentpPass			'#업체PASSWORD | SKB에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&doFlag=30"							'진행단계 | 25:출하지시대상(default), 30:출고대상, 40:배송완료대상
	addParam = addParam & "&orderGb=10"							'배송 구분 값 | 00:전체(default), 10:주문, 40:교환
	addParam = addParam & "&bDate="& xmlSelldate				'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate				'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="							'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'rw addParam
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", skstoaAPIURL & "/partner/delivery/order-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				If returnCode = "200" Then
					Set orderList = strObj.orderList
						If orderList.length > 0 Then
							For i=0 to orderList.length-1
								Set oMaster = new COrderMasterItem
									oMaster.FSellSite 			= sellsite
									oMaster.FOutMallOrderSerial = orderList.get(i).orderNo
									oMaster.FSellDate 			= Left(orderList.get(i).procDate, 8)
									oMaster.FSellDate			= Left(oMaster.FSellDate, 4) & "-" & Right(Left(oMaster.FSellDate,6), 2) & "-" & Right(oMaster.FSellDate, 2)
									oMaster.FPayType			= "50"
									oMaster.FPaydate			= oMaster.FSellDate
									oMaster.FOrderUserID		= ""
									oMaster.FOrderName			= LEFT(html2db(orderList.get(i).custName), 28)
									On Error Resume Next
										oMaster.FOrderTelNo		= LEFT(html2db(orderList.get(i).custTel), 16)
										If Err.number <> 0 Then
											oMaster.FOrderTelNo = ""
										End If
									On Error Goto 0

									On Error Resume Next
										oMaster.FOrderHpNo			= LEFT(html2db(orderList.get(i).custHp), 16)
										If Err.number <> 0 Then
											oMaster.FOrderHpNo = ""
										End If
									On Error Goto 0

									if Len(CStr(oMaster.FOrderTelNo)) <= 3 then
										oMaster.FOrderTelNo = oMaster.FOrderHpNo
									end if

									oMaster.FOrderEmail			= ""
									oMaster.FReceiveName		= LEFT(html2db(orderList.get(i).receiverName), 28)
									On Error Resume Next
										oMaster.FReceiveTelNo	= LEFT(html2db(orderList.get(i).receiverTel), 16)
										If Err.number <> 0 Then
											oMaster.FReceiveTelNo = ""
										End If
									On Error Goto 0

									On Error Resume Next
										oMaster.FReceiveHpNo		= LEFT(html2db(orderList.get(i).receiverHp), 16)
										If Err.number <> 0 Then
											oMaster.FReceiveHpNo = ""
										End If
									On Error Goto 0
									
									if Len(CStr(oMaster.FReceiveTelNo)) <= 3 then
										oMaster.FReceiveTelNo = oMaster.FReceiveHpNo
									end if
									
									On Error Resume Next
										oMaster.Fdeliverymemo		= html2db(orderList.get(i).msg)
										If Err.number <> 0 Then
											oMaster.Fdeliverymemo	= ""
										End If
									On Error Goto 0
									oMaster.FdeliverPay 		= orderList.get(i).shipCost
									
									On Error Resume Next
										oMaster.FReceiveZipCode		= html2db(orderList.get(i).receiverPost)
										If Err.number <> 0 Then
											oMaster.FReceiveZipCode		= ""
										End If
									On Error Goto 0

									oMaster.FReceiveAddr1		= html2db(orderList.get(i).receiverAddr)

									'''주소와 상세주소가 같은경우 3번째 Blank에서 끊음.
									POS1 = 0
									POS2 = 0
									POS3 = 0
									POS1 = InStr(oMaster.FReceiveAddr1," ")
									''rw "POS1="&POS1
									IF (POS1>0) then
										POS2 = InStr(MID(oMaster.FReceiveAddr1,POS1+1,512)," ")
										''rw "POS2="&POS2
										IF POS2>0 then
											POS3 = InStr(MID(oMaster.FReceiveAddr1,POS1+POS2+1,512)," ")
											IF POS3>0 then
												oMaster.FReceiveAddr2=MID(oMaster.FReceiveAddr1,POS1+POS2+POS3+1,512)
												oMaster.FReceiveAddr1=LEFT(oMaster.FReceiveAddr1, POS1 + POS2 + POS3 - 1)
											END IF
										END IF
									END IF

									Redim oDetailArr(0)
									Set oDetailArr(0) = new COrderDetail
										oDetailArr(0).FdetailSeq			= orderList.get(i).orderGSeq & "-" & orderList.get(i).orderDSeq & "-" & orderList.get(i).orderWSeq
										tmpItemid = ""
										tmpItemOption = ""
										strsql = ""
										strsql = strsql & " SELECT TOP 1 itemid "
										strsql = strsql & " FROM db_etcmall.dbo.tbl_skstoa_regitem "
										strsql = strsql & " WHERE skstoaGoodno = '"& orderList.get(i).goodsCode &"' "
										rsget.CursorLocation = adUseClient
										rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
										If (Not rsget.EOF) Then
											tmpItemid = rsget("itemid")
										End If
										rsget.Close

										strsql = ""
										strsql = strsql & " SELECT TOP 1 itemoption "
										strsql = strsql & " FROM db_item.dbo.tbl_OutMall_regedoption "
										strsql = strsql & " WHERE itemid = '"& tmpItemid &"' "
										strsql = strsql & " and outmallOptCode = '"& orderList.get(i).goodsdtCode &"' "
										strsql = strsql & " and mallid = 'skstoa' "
										rsget.CursorLocation = adUseClient
										rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
										If (Not rsget.EOF) Then
											tmpItemOption = rsget("itemoption")
										Else
											tmpOptionSeq = tmpOptionSeq + 1
											tmpItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
										End If
										rsget.Close

										If tmpItemid = "" Then
											tmpItemid = "99999999"
										End If

										oDetailArr(0).FItemID				= tmpItemid
										oDetailArr(0).FItemOption			= tmpItemOption
										oDetailArr(0).FOutMallItemID		= orderList.get(i).goodsCode
										oDetailArr(0).FOutMallItemOption	= orderList.get(i).goodsdtCode
										oDetailArr(0).FOutMallItemName		= html2db(orderList.get(i).goodsName)
										oDetailArr(0).FOutMallItemOptionName = html2db(orderList.get(i).goodsdtInfo)
										oDetailArr(0).FItemNo				= orderList.get(i).syslast
										oDetailArr(0).Fitemcost				= Clng(orderList.get(i).salePrice)
										oDetailArr(0).FReducedPrice			= oDetailArr(0).Fitemcost
										oDetailArr(0).FOutMallCouponPrice	= 0
										oDetailArr(0).FTenCouponPrice		= 0
										oDetailArr(0).FrequireDetail 		= ""

										If (SaveOrderToDB(oMaster, oDetailArr) = True) Then
											successCnt = successCnt + 1
										End If
									Set oDetailArr = Nothing
								Set oMaster = nothing
							Next
							response.write "주문입력(" & successCnt & ")" & "<br />"
'							response.end
						End If
					Set orderList = nothing
				Else
					response.write "내역없음(0)<br />"
					GetOrderFrom_skstoa = True
					Set strObj = Nothing
					rw "------------"
					Exit Function
				End If
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
			GetOrderFrom_skstoa = True
			Set strObj = Nothing
			rw "------------"
			Exit Function
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_shintvshopping_Gubun1(selldate)
	dim sellsite : sellsite = "shintvshopping"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0
	Dim orderGb, orderNo, orderGSeq, orderDSeq, orderWSeq, goodsCode, goodsdtCode
	GetOrderFrom_shintvshopping_Gubun1 = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & entpId				'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & entpPass			'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&doFlag=25"						'주문 진행 단계 | 25:출하지시대상(default), 30:출고대상, 40:배송완료대상
	addParam = addParam & "&orderGb=10"						'배송 구분 값 | 00:전체(default), 10:주문, 40:교환
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'	addParam = addParam & "&slipINo="						'운송장식별번호 | doFlag >= 30 일 경우에만 해당 매개변수로 조회 가능
'rw addParam
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/delivery/order-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,800000,800000,800000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				If returnCode = "200" Then
					Set orderList = strObj.orderList
						If orderList.length > 0 Then
							For i=0 to orderList.length-1
								orderGb = orderList.get(i).orderGb
								orderNo = orderList.get(i).orderNo
								orderGSeq = orderList.get(i).orderGSeq
								orderDSeq = orderList.get(i).orderDSeq
								orderWSeq = orderList.get(i).orderWSeq
								goodsCode = orderList.get(i).goodsCode
								goodsdtCode = orderList.get(i).goodsdtCode

								strSql = ""
								strSql = strSql & " IF NOT EXISTS (SELECT idx FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] "
								strSql = strSql & " WHERE sellsite = '" & sellsite & "' "
								strSql = strSql & " and OutMallOrderSerial = '" & orderNo & "' "
								strSql = strSql & " and orderGSeq = '" & orderGSeq & "' "
								strSql = strSql & " and orderDSeq = '" & orderDSeq & "' "
								strSql = strSql & " and orderWSeq = '" & orderWSeq & "' "
								strSql = strSql & " and shintvshoppingGoodNo = '"& goodsCode &"' "
								strSql = strSql & " and outmallOptCode = '" & goodsdtCode & "' "
								strSql = strSql & " ) "
								strSql = strSql & " BEGIN "
								strSql = strSql & " 	INSERT INTO db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] ([sellsite], [orderGb], [outmallorderserial], [orderGSeq], [orderDSeq], [orderWSeq], [shintvshoppingGoodNo], [outmallOptCode], [regdate]) "
								strSql = strSql & " 	VALUES ('"& sellsite &"', '"& orderGb &"', '"& orderNo &"', '"& orderGSeq &"', '"& orderDSeq &"', '"& orderWSeq &"', '"& goodsCode &"', '"& goodsdtCode &"', GETDATE()) "
								strSql = strSql & " END "
								dbget.Execute strSql
							Next
						End If
					Set orderList = nothing
				Else
					response.write "내역없음(0)<br />"
					GetOrderFrom_shintvshopping_Gubun1 = True
					Set strObj = Nothing
					rw "------------"
					Exit Function
				End If
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
			GetOrderFrom_shintvshopping_Gubun1 = True
			Set strObj = Nothing
			rw "------------"
			Exit Function
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_shintvshopping_Gubun2(outmallorderserial, orderGSeq, orderDSeq, orderWSeq, shintvshoppingGoodNo, outmallOptCode)
	dim sellsite : sellsite = "shintvshopping"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0

	addParam = ""
	addParam = addParam & "linkCode=" & linkCode				'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode				'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & entpId					'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & entpPass				'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&orderGb=10"							'#배송 구분 값 | 10:주문, 40:교환
	addParam = addParam & "&orderNo=" & outmallorderserial		'#주문번호
	addParam = addParam & "&orderGSeq=" & orderGSeq				'#상품순번 | 주문번호 G시퀀스 코드
	addParam = addParam & "&orderDSeq=" & orderDSeq				'#세트순번 | 주문번호 D시퀀스 코드
	addParam = addParam & "&orderWSeq=" & orderWSeq				'#처리순번 | 주문번호 W시퀀스 코드
	addParam = addParam & "&goodsCode=" & shintvshoppingGoodNo	'#판매상품코드
	addParam = addParam & "&goodsdtCode=" & outmallOptCode		'#판매단품코드

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", shintvshoppingAPIURL & "/partner/delivery/delivery-order", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,800000,800000,800000
		objXML.Send(addParam)
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				If returnCode = "200" Then
					rw "orderNo : " & outmallorderserial & " |||| goodsCode : " & shintvshoppingGoodNo & " [Deliver Ready Complete] "
				Else
					response.write BinaryToText(objXML.ResponseBody,"utf-8")
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_shintvshopping(selldate)
	dim sellsite : sellsite = "shintvshopping"
	dim xmlURL, xmlSelldate, iRbody, addParam
	dim objXML, xmlDOM
	dim strObj, iMessage, orderList, returnCode, returnStatus
	dim i, strSql
	dim successCnt : successCnt = 0
	Dim orderGb, orderNo, orderGSeq, orderDSeq, orderWSeq, goodsCode, goodsdtCode
	dim oMaster, oDetail, oDetailArr, tmpItemid, tmpItemOption
	dim tmpOptionSeq : tmpOptionSeq = 0
	Dim POS1, POS2, POS3
	GetOrderFrom_shintvshopping = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	addParam = ""
	addParam = addParam & "linkCode=" & linkCode			'#연결코드 | 샵링커: SLINK [ TCODE.LGROUP : xxx ]
	addParam = addParam & "&entpCode=" & entpCode			'#업체코드 | 신세계TV쇼핑에서 부여한 업체코드 6자리
	addParam = addParam & "&entpId=" & entpId				'#업체사용자ID | 신세계TV쇼핑에서 부여한 업체사용자 ID
	addParam = addParam & "&entpPass=" & entpPass			'#업체PASSWORD | 신세계TV쇼핑에서 등록한 업체사용자 비밀번호
	addParam = addParam & "&doFlag=30"						'주문 진행 단계 | 25:출하지시대상(default), 30:출고대상, 40:배송완료대상
	addParam = addParam & "&orderGb=10"						'배송 구분 값 | 00:전체(default), 10:주문, 40:교환
	addParam = addParam & "&bDate="& xmlSelldate			'#조회 시작일자 | 주문승인/교환접수일 기준, YYYYMMDD 타입. ex) 20140520
	addParam = addParam & "&eDate="& xmlSelldate			'#조회 마지막일자 | 주문승인/교환접수일 기준 YYYYMMDD 타입. ex) 20140520
'	addParam = addParam & "&orderNo="						'주문코드 | 주문번호를 이용한 검색. 숫자만 허용되며 사용시 14자리(orderNo) 또는 23자리(orderNo,orderGSeq,orderDSeq,orderWSeq)의 값만 허용
'	addParam = addParam & "&slipINo="						'운송장식별번호 | doFlag >= 30 일 경우에만 해당 매개변수로 조회 가능
'rw addParam
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", shintvshoppingAPIURL & "/partner/delivery/order-list?" & addParam , false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,800000,800000,800000
		objXML.Send()
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				returnStatus	= strObj.status
				returnCode		= strObj.code
				iMessage		= strObj.message
				rw BinaryToText(objXML.ResponseBody,"utf-8")
				If returnCode = "200" Then
					Set orderList = strObj.orderList
						If orderList.length > 0 Then
							For i=0 to orderList.length-1
								Set oMaster = new COrderMasterItem
									oMaster.FSellSite 			= sellsite
									oMaster.FOutMallOrderSerial = orderList.get(i).orderNo
									oMaster.FSellDate 			= Left(orderList.get(i).procDate, 8)
									oMaster.FSellDate			= Left(oMaster.FSellDate, 4) & "-" & Right(Left(oMaster.FSellDate,6), 2) & "-" & Right(oMaster.FSellDate, 2)
									oMaster.FPayType			= "50"
									oMaster.FPaydate			= oMaster.FSellDate
									oMaster.FOrderUserID		= ""
									oMaster.FOrderName			= LEFT(html2db(orderList.get(i).custName), 28)
									On Error Resume Next
										oMaster.FOrderTelNo		= LEFT(html2db(orderList.get(i).custTel), 16)
										If Err.number <> 0 Then
											oMaster.FOrderTelNo = ""
										End If
									On Error Goto 0

									On Error Resume Next
										oMaster.FOrderHpNo			= LEFT(html2db(orderList.get(i).custHp), 16)
										If Err.number <> 0 Then
											oMaster.FOrderHpNo = ""
										End If
									On Error Goto 0

									if Len(CStr(oMaster.FOrderTelNo)) <= 3 then
										oMaster.FOrderTelNo = oMaster.FOrderHpNo
									end if

									oMaster.FOrderEmail			= ""
									oMaster.FReceiveName		= LEFT(html2db(orderList.get(i).receiverName), 28)
									On Error Resume Next
										oMaster.FReceiveTelNo	= LEFT(html2db(orderList.get(i).receiverTel), 16)
										If Err.number <> 0 Then
											oMaster.FReceiveTelNo = ""
										End If
									On Error Goto 0

									On Error Resume Next
										oMaster.FReceiveHpNo		= LEFT(html2db(orderList.get(i).receiverHp), 16)
										If Err.number <> 0 Then
											oMaster.FReceiveHpNo = ""
										End If
									On Error Goto 0
									
									if Len(CStr(oMaster.FReceiveTelNo)) <= 3 then
										oMaster.FReceiveTelNo = oMaster.FReceiveHpNo
									end if
									
									On Error Resume Next
										oMaster.Fdeliverymemo		= html2db(orderList.get(i).msg)
										If Err.number <> 0 Then
											oMaster.Fdeliverymemo	= ""
										End If
									On Error Goto 0
									oMaster.FdeliverPay 		= orderList.get(i).shipCost
									
									On Error Resume Next
										oMaster.FReceiveZipCode		= html2db(orderList.get(i).receiverPostNo)
										If Err.number <> 0 Then
											oMaster.FReceiveZipCode		= ""
										End If
									On Error Goto 0

									oMaster.FReceiveAddr1		= html2db(orderList.get(i).receiverAddr)
									oMaster.FbeasongNum11st		= orderList.get(i).slipINo

									'''주소와 상세주소가 같은경우 3번째 Blank에서 끊음.
									POS1 = 0
									POS2 = 0
									POS3 = 0
									POS1 = InStr(oMaster.FReceiveAddr1," ")
									''rw "POS1="&POS1
									IF (POS1>0) then
										POS2 = InStr(MID(oMaster.FReceiveAddr1,POS1+1,512)," ")
										''rw "POS2="&POS2
										IF POS2>0 then
											POS3 = InStr(MID(oMaster.FReceiveAddr1,POS1+POS2+1,512)," ")
											IF POS3>0 then
												oMaster.FReceiveAddr2=MID(oMaster.FReceiveAddr1,POS1+POS2+POS3+1,512)
												oMaster.FReceiveAddr1=LEFT(oMaster.FReceiveAddr1, POS1 + POS2 + POS3 - 1)
											END IF
										END IF
									END IF

									Redim oDetailArr(0)
									Set oDetailArr(0) = new COrderDetail
										oDetailArr(0).FdetailSeq			= orderList.get(i).orderGSeq & "-" & orderList.get(i).orderDSeq & "-" & orderList.get(i).orderWSeq
										tmpItemid = ""
										tmpItemOption = ""
										strsql = ""
										strsql = strsql & " SELECT TOP 1 itemid "
										strsql = strsql & " FROM db_etcmall.dbo.tbl_shintvshopping_regitem "
										strsql = strsql & " WHERE shintvshoppingGoodno = '"& orderList.get(i).goodsCode &"' "
										rsget.CursorLocation = adUseClient
										rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
										If (Not rsget.EOF) Then
											tmpItemid = rsget("itemid")
										End If
										rsget.Close

										strsql = ""
										strsql = strsql & " SELECT TOP 1 itemoption "
										strsql = strsql & " FROM db_item.dbo.tbl_OutMall_regedoption "
										strsql = strsql & " WHERE itemid = '"& tmpItemid &"' "
										strsql = strsql & " and outmallOptCode = '"& orderList.get(i).goodsdtCode &"' "
										strsql = strsql & " and mallid = 'shintvshopping' "
										rsget.CursorLocation = adUseClient
										rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
										If (Not rsget.EOF) Then
											tmpItemOption = rsget("itemoption")
										Else
											tmpOptionSeq = tmpOptionSeq + 1
											tmpItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
										End If
										rsget.Close

										If tmpItemid = "" Then
											tmpItemid = "99999999"
										End If

										oDetailArr(0).FItemID				= tmpItemid
										oDetailArr(0).FItemOption			= tmpItemOption
										oDetailArr(0).FOutMallItemID		= orderList.get(i).goodsCode
										oDetailArr(0).FOutMallItemOption	= orderList.get(i).goodsdtCode
										oDetailArr(0).FOutMallItemName		= html2db(orderList.get(i).goodsName)
										oDetailArr(0).FOutMallItemOptionName = html2db(orderList.get(i).goodsdtInfo)
										oDetailArr(0).FItemNo				= orderList.get(i).syslast
										oDetailArr(0).Fitemcost				= Clng(orderList.get(i).salePrice)
										oDetailArr(0).FReducedPrice			= oDetailArr(0).Fitemcost
										oDetailArr(0).FOutMallCouponPrice	= 0
										oDetailArr(0).FTenCouponPrice		= 0
										oDetailArr(0).FrequireDetail 		= ""

										If (SaveOrderToDB(oMaster, oDetailArr) = True) Then
											successCnt = successCnt + 1
										End If
									Set oDetailArr = Nothing
								Set oMaster = nothing
							Next
							response.write "주문입력(" & successCnt & ")" & "<br />"
'							response.end
						End If
					Set orderList = nothing
				Else
					response.write "내역없음(0)<br />"
					GetOrderFrom_shintvshopping = True
					Set strObj = Nothing
					rw "------------"
					Exit Function
				End If
			Set strObj = nothing
		Else
			rw BinaryToText(objXML.ResponseBody,"utf-8")
			GetOrderFrom_shintvshopping = True
			Set strObj = Nothing
			rw "------------"
			Exit Function
		End If
	Set objXML= nothing
End Function

Function GetOrderFrom_lotteon(selldate)
	dim sellsite : sellsite = "lotteon"
	dim xmlURL, xmlSelldate
	dim objXML, strObj, objData, aKey, jParam, requireDetailObj, requireDetail
	dim masterCnt
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim tmpStr, pos, obj, orderList
	dim i, j, k, strsql
	dim found, deliveryDate
	dim successCnt : successCnt = 0
	Dim returnCode, dataVal
	Dim apiUrl, apiKey

	GetOrderFrom_lotteon = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	apiUrl = getApiUrl("lotteon")
	apiKey = getApiKey("lotteon")
	apiUrl = apiUrl & "/v1/openapi/delivery/v1/SellerDeliveryOrdersSearch"

	'// =======================================================================
	Set obj = jsObject()
		obj("srchStrtDt") = xmlSelldate&"000000"	'#검색시작일시 yyyymmddhhmmss 출고/회수지시 생성일
		obj("srchEndDt") = xmlSelldate&"235959"		'#검색종료일시 yyyymmddhhmmss
'		obj("srchStrtDt") = xmlSelldate				'#검색시작일시 yyyymmddhhmmss 출고/회수지시 생성일
'		obj("srchEndDt") = xmlSelldate				'#검색종료일시 yyyymmddhhmmss
		obj("odNo") = ""							'주문번호
		obj("odPrgsStepCd") = ""					'주문진행단계코드 | 11:출고지시,12:상품준비,13:발송완료,14:배송완료,15:수취완료,23:회수지시,24:회수진행,25:회수확정,26:반품완료
		obj("odTypCd") = "10"						'주문유형코드 | 10:주문, 30:교환, 40:반품, 50:AS
		obj("IfCplYN") = ""							'N 으로 요청시 연동완료여부 실패건만 조회
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Authorization", "Bearer " & apiKey
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Accept-Language", "ko"
		objXML.setRequestHeader "X-Timezone", "GMT+09:00"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
' rw xmlURL
' rw apiKey
' rw jParam
' rw "----------------------------"
' rw BinaryToText(objXML.ResponseBody,"utf-8")
' If xmlSelldate = "20230309" Then
' 	response.end
' End If

		If objXML.Status <> "200" Then
			If IsAutoScript Then
				response.write "ERROR : 통신오류"
			Else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			End If

			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
		Set strObj = JSON.parse(objData)
			returnCode		= strObj.returnCode
			If returnCode = "0000" Then
				deliveryDate = ""
				Set orderList = strObj.data.deliveryOrderList
					If orderList.length > 0 Then
						'rw "건수 : " & orderList.length & "!!!!!!!!!!!!!!!!!!!!!!"
						'  If orderList.length > 0 Then
						'  	rw BinaryToText(objXML.ResponseBody,"utf-8")
						'  End If
						' response.end

						For i=0 to orderList.length-1
							requireDetail = ""
'							On Error Resume Next
							If orderList.get(i).pdAdtnOptJsn <> "" Then
								Set requireDetailObj = JSON.parse(orderList.get(i).pdAdtnOptJsn)
									If requireDetailObj.length > 0 Then
										If orderList.get(i).odNo = "2023030919691552" Then
											requireDetail = "참 잘했어요, 김하림 선생님"
										Else
											requireDetail = requireDetailObj.get(0).adtnOptVal
										End If
									End If
								Set requireDetailObj = nothing
							End If

							If orderList.get(i).odTypCd = "10" Then
							' If Err.number <> 0 Then
							' 	requireDetail = ""
							' End If
							' On Error Goto 0
								Set oMaster = new COrderMasterItem
									oMaster.FSellSite 			= sellsite
									oMaster.FOutMallOrderSerial = orderList.get(i).odNo
									oMaster.FbeasongNum11st		= orderList.get(i).procSeq
									oMaster.FSellDate 			= Left(Now(), 10)
									oMaster.FPayType			= "50"
									oMaster.FPaydate			= oMaster.FSellDate
									oMaster.FOrderUserID		= ""
									oMaster.FOrderName			= html2db(orderList.get(i).odrNm)
									oMaster.FOrderTelNo			= html2db(orderList.get(i).telNo)
									oMaster.FOrderHpNo			= html2db(orderList.get(i).mphnNo)
									oMaster.FOrderEmail			= html2db(orderList.get(i).emlAddr)
									oMaster.FReceiveName		= html2db(orderList.get(i).dvpCustNm)
									If Len(oMaster.FReceiveName) > 32 Then
										oMaster.FReceiveName = oMaster.FOrderName
									End If

									oMaster.FReceiveTelNo		= html2db(orderList.get(i).dvpTelNo)
									oMaster.FReceiveHpNo		= html2db(orderList.get(i).dvpMphnNo)
									oMaster.Fdeliverymemo		= html2db(RemoveWhiteSpaceChar(orderList.get(i).dvMsg))
									If (oMaster.Fdeliverymemo = "null") Then
										oMaster.Fdeliverymemo = ""
									End if

									oMaster.FdeliverPay			= orderList.get(i).dvCst
									oMaster.FReceiveZipCode		= html2db(RemoveWhiteSpaceChar(orderList.get(i).dvpZipNo))

									If Len(oMaster.FReceiveZipCode) = 4 Then
										oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
									End if

	'								If Len(oMaster.FReceiveZipCode) <= 5 and orderList.get(i).dvpStnmZipAddr <> "" Then
										'// 도로명주소
										oMaster.FReceiveAddr1		= html2db(orderList.get(i).dvpStnmZipAddr)
										oMaster.FReceiveAddr2		= html2db(orderList.get(i).dvpStnmDtlAddr)
	'								Else
										'// 구주소..(※ 5/13부로 지번주소는 제외 됩니다. - 롯데on은 현재 도로명주소를 기준으로 관리되고 있습니다.)
	'									oMaster.FReceiveAddr1		= html2db(orderList.get(i).dvpJbZipAddr)
	'									oMaster.FReceiveAddr2		= html2db(orderList.get(i).dvpJbDtlAddr)
	'								End if
									deliveryDate = orderList.get(i).owhoDttm
									Redim oDetailArr(0)
									Set oDetailArr(0) = new COrderDetail
										oDetailArr(0).FdetailSeq			= orderList.get(i).odSeq
										oDetailArr(0).FItemID				= orderList.get(i).epdNo
										oDetailArr(0).FItemOption			= orderList.get(i).eitmNo
										oDetailArr(0).FOutMallItemID		= orderList.get(i).spdNo
										oDetailArr(0).FOutMallItemOption	= orderList.get(i).sitmNo
										oDetailArr(0).FOutMallItemName		= html2db(orderList.get(i).spdNm)
										oDetailArr(0).FOutMallItemOptionName = html2db(orderList.get(i).sitmNm)
										oDetailArr(0).FItemNo				= orderList.get(i).odQty
										oDetailArr(0).Fitemcost				= Clng(orderList.get(i).slPrc)
										oDetailArr(0).FReducedPrice			= Clng(orderList.get(i).slAmt / orderList.get(i).odQty)	'할인금액이 안 넘어옴
										oDetailArr(0).FOutMallCouponPrice	= 0
										oDetailArr(0).FTenCouponPrice		= 0
										If oDetailArr(0).FItemOption = "" Then
											oDetailArr(0).FItemOption = "0000"
											oDetailArr(0).FOutMallItemOption = "0000"
										End If
										oDetailArr(0).FrequireDetail = requireDetail
										If (SaveOrderToDB(oMaster, oDetailArr) = True) Then
											successCnt = successCnt + 1
											strsql = ""
											strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid, outMallGoodsNo, ItemOrderCount, outMallOptionNo, isTenConfirmSend, deliveryDate) "
											strsql = strsql & " VALUES ('"&orderList.get(i).odNo&"', '"&orderList.get(i).odSeq&"', '"&orderList.get(i).procSeq&"', 'N', getdate(), 'lotteon', '"&oDetailArr(0).FOutMallItemID&"', '"&orderList.get(i).odQty&"', '"&orderList.get(i).sitmNo&"', 'N', '"&deliveryDate&"')"
											dbget.Execute strSql
										End If
									Set oDetailArr = Nothing
								Set oMaster = nothing
							End If
						Next
						response.write "주문입력(" & successCnt & ")" & "<br />"
						rw "------------"
					Else
						''if IsAutoScript then
							response.write "내역없음(0)<br />"
						''end if

						GetOrderFrom_lotteon = True
						Set orderList = Nothing
						Set strObj = Nothing
						rw "------------"
						exit function
					End If
				Set objXML = nothing
			End If
		Set strObj = nothing
	Set objXML = nothing
'response.end
End Function

Function GetOrderFrom_wetoo1300k(selldate)
	dim sellsite : sellsite = "wetoo1300k"
	dim oMaster, oDetail, oDetailArr
	Dim POS1, POS2, POS3
    Dim objXML, obj, strParam, objData, strObj, categoryList, i, strSql, stDate, edDate, wetoo1300kAPIURL, company_auth, company_code, j
	Dim returnCode, iMessage, infoList, productList
	dim successCnt : successCnt = 0
	If application("Svr_Info") = "Dev" Then
		wetoo1300kAPIURL = "https://ts.1300k.com"
		company_auth = "1ac6e7cd04fc587cc26722b1cbaaa75c"
		company_code = "C927"
	Else
		wetoo1300kAPIURL = "http://api.1300k.com"
		company_auth = "f91f60a59e32425e4f22c3d20cf4f7b7"
		company_code = "C927"
	End If

	stDate = Replace(selldate, "-", "")
	edDate = Replace(Dateadd("d", 1, selldate), "-", "")

	Set obj = jsObject()
		Set obj("header") = jsObject()
			obj("header")("company_code") = company_code
			obj("header")("company_auth") = company_auth
			Set obj("order_date") = jsObject()
				obj("order_date")("st_date") = stDate&"0000"		'#시작시간 YYYYMMDDHHMM
				obj("order_date")("ed_date") = edDate&"0000"		'#종료시간 YYYYMMDDHHMM
				obj("order_date")("order_status") = ""
			strParam = obj.jsString
	Set obj = nothing

	GetOrderFrom_wetoo1300k = False

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", wetoo1300kAPIURL & "/enterstore/api/order_info.html", false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(strParam)
		If objXML.Status <> "200" Then
			If IsAutoScript Then
				response.write "ERROR : 통신오류"
			Else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			End If

			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
'rw objData
		Set strObj = JSON.parse(objData)
			returnCode		= strObj.code
			iMessage		= strObj.message

			If returnCode = "00" Then
				Set infoList = strObj.result.info
					If infoList.length > 0 Then
						For i=0 to infoList.length-1
							If NOT(isObject(infoList.get(i).product)) Then
								Exit for
							End If
							Set oMaster = new COrderMasterItem
								oMaster.FSellSite 			= sellsite
								oMaster.FOutMallOrderSerial = Trim(infoList.get(i).order.order_no)		'주문번호
								oMaster.FSellDate 			= Trim(infoList.get(i).order.order_date)	'주문일자
								oMaster.FPayType			= "50"
								oMaster.FPaydate			= Trim(infoList.get(i).order.payment_date)	'결제일자
								oMaster.FSellDate = Left(oMaster.FSellDate,4)&"-"&Mid(oMaster.FSellDate,5,2)&"-"&Mid(oMaster.FSellDate,7,2)&" "&Mid(oMaster.FSellDate,9,2)&":"&Mid(oMaster.FSellDate,11,2)&":"&Mid(oMaster.FSellDate,13,2)
								oMaster.FPaydate = Left(oMaster.FPaydate,4)&"-"&Mid(oMaster.FPaydate,5,2)&"-"&Mid(oMaster.FPaydate,7,2)&" "&Mid(oMaster.FPaydate,9,2)&":"&Mid(oMaster.FPaydate,11,2)&":"&Mid(oMaster.FPaydate,13,2)

								oMaster.FOrderUserID		= ""
								oMaster.FOrderName			= html2db(infoList.get(i).order.order_name)	'주문자명
								oMaster.FOrderTelNo			= html2db(infoList.get(i).order.order_tel)	'주문자전화변호
								oMaster.FOrderHpNo			= html2db(infoList.get(i).order.order_tel)	'주문자전화변호
								oMaster.FOrderEmail			= html2db(infoList.get(i).order.order_email)	'주문자이메일
								oMaster.FReceiveName		= html2db(infoList.get(i).order.recvr_name)	'받는사람명
								If Len(oMaster.FReceiveName) > 32 Then
									oMaster.FReceiveName = oMaster.FOrderName
								End If

								oMaster.FReceiveTelNo		= html2db(infoList.get(i).order.recvr_tel)	'받는사람전화번호
								oMaster.FReceiveHpNo		= html2db(infoList.get(i).order.recvr_cell)	'받는사람핸드폰
								oMaster.Fdeliverymemo		= html2db(RemoveWhiteSpaceChar(infoList.get(i).order.delivery_message))	'배송메세지
								If (oMaster.Fdeliverymemo = "null") Then
									oMaster.Fdeliverymemo = ""
								End if

								oMaster.FdeliverPay			= infoList.get(i).order.delivery_cost
								oMaster.FReceiveZipCode		= html2db(RemoveWhiteSpaceChar(infoList.get(i).order.recvr_zipcode))

								If Len(oMaster.FReceiveZipCode) = 4 Then
									oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
								End if
								oMaster.FReceiveAddr1		= html2db(infoList.get(i).order.recvr_addr)

								'''주소와 상세주소가 같은경우 3번째 Blank에서 끊음.
								POS1 = 0
								POS2 = 0
								POS3 = 0
								POS1 = InStr(oMaster.FReceiveAddr1," ")
								''rw "POS1="&POS1
								IF (POS1>0) then
									POS2 = InStr(MID(oMaster.FReceiveAddr1,POS1+1,512)," ")
									''rw "POS2="&POS2
									IF POS2>0 then
										POS3 = InStr(MID(oMaster.FReceiveAddr1,POS1+POS2+1,512)," ")
										IF POS3>0 then
											oMaster.FReceiveAddr2=MID(oMaster.FReceiveAddr1,POS1+POS2+POS3+1,512)
											oMaster.FReceiveAddr1=LEFT(oMaster.FReceiveAddr1, POS1 + POS2 + POS3 - 1)
										END IF
									END IF
								END IF

								Set productList = infoList.get(i).product
									If productList.length > 0 Then
										Redim oDetailArr(productList.length - 1)
										For j=0 to productList.length - 1
											Set oDetailArr(j) = new COrderDetail
												oDetailArr(j).FdetailSeq = productList.get(j).seq_no						'일련번호
												oDetailArr(j).FOutMallItemID = productList.get(j).product_code				'상품코드
												oDetailArr(j).FOutMallItemName = html2db(productList.get(j).product_name)	'상품명
												oDetailArr(j).FOutMallItemOption = productList.get(j).opt_no				'옵션번호
												oDetailArr(j).FOutMallItemOptionName = html2db(productList.get(j).opt_name)	'옵션명
												oDetailArr(j).FItemNo = productList.get(j).qty								'수량
												oDetailArr(j).Fitemcost = productList.get(j).sale_price - productList.get(j).dc_price	'판매가
												oDetailArr(j).FReducedPrice = productList.get(j).sale_price - productList.get(j).dc_price - productList.get(j).cpn_price
												oDetailArr(j).FItemID = productList.get(j).company_product_code				'업체상품코드
												oDetailArr(j).FItemOption = productList.get(j).company_opt_no				'업체옵션번호
'												rw productList.get(j).change_status											'추가상태
'												rw productList.get(j).change_seq_no											'원주문 일련번호
										Next
										If (SaveOrderToDB(oMaster, oDetailArr) = True) Then
											successCnt = successCnt + 1
										End If
									End If
								Set productList = nothing
							Set oMaster = nothing
						Next
						response.write "주문입력(" & successCnt & ")" & "<br />"
					Else
						response.write "내역없음(0)<br />"
						GetOrderFrom_wetoo1300k = True
						Set infoList = Nothing
						Set strObj = Nothing
						Set objXML = nothing
						rw "------------"
						exit function
					End If
				Set infoList = nothing
			End If
		Set strObj = nothing
	Set objXML = nothing
End Function

Function GetOrderFrom_lfmall(selldate)
	dim sellsite : sellsite = "lfmall"
	dim xmlURL, xmlSelldate
	dim objXML, strObj, objData, aKey, jParam, requireDetailObj, requireDetail
	dim masterCnt, optCode, optNm
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim tmpStr, pos, obj, orderList
	dim i, j, k, strsql
	dim found, deliveryDate
	dim successCnt : successCnt = 0
	Dim returnCode, dataVal
	Dim apiUrl, apiKey, istrParam
	dim tmpOptionSeq : tmpOptionSeq = 0
	GetOrderFrom_lfmall = False

	istrParam = ""
	istrParam = istrParam & "<?xml version=""1.0"" encoding=""UTF-8""?>"
	istrParam = istrParam & "<OrderInfo>"
	istrParam = istrParam & "	<Header>"
	istrParam = istrParam & "		<AuthId><![CDATA[tenten]]></AuthId>"
	istrParam = istrParam & "		<AuthKey><![CDATA[Ten1010*!!]]></AuthKey>"
	istrParam = istrParam & "		<Format>XML</Format>"
	istrParam = istrParam & "		<Charset>UTF-8</Charset>"
	istrParam = istrParam & "	</Header>"
	istrParam = istrParam & "	<Body>"
	istrParam = istrParam & "		<Order>"
	istrParam = istrParam & "			<OrdStartDate>"& Replace(selldate, "-", "") &"</OrdStartDate>"
	istrParam = istrParam & "            <OrdEndDate>"& Replace(selldate, "-", "") &"</OrdEndDate>"
	istrParam = istrParam & "            <OrdStatusCode>30</OrdStatusCode>"
	istrParam = istrParam & "		</Order>"
	istrParam = istrParam & "	</Body>"
	istrParam = istrParam & "</OrderInfo>"

    Dim  iRbody, iMessage
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://b2b.lfmall.co.kr/interface.do?cmd=getOrderList", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(REQUEST_XML)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

rw iRbody
'response.end
				Set obj = Nothing

				if (xmlDOM.getElementsByTagName("OrderInfo/Body/Order").length < 1) then
					''if IsAutoScript then
						response.write "내역없음 : 종료" & "<br />"
					''end if

					GetOrderFrom_lfmall = True
					Set xmlDOM = Nothing
					Set objXML = Nothing
					exit function
				else
					response.write "건수(" & xmlDOM.getElementsByTagName("OrderInfo/Body/Order").length & ") " & "<br />"
				end if

				set objMasterListXML = xmlDOM.getElementsByTagName("OrderInfo/Body/Order")
				For each objMasterOneXML in objMasterListXML
					optCode = ""
					optNm = ""
					Set oMaster = new COrderMasterItem

					oMaster.FSellSite 			= sellsite
					oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("OrdNo")(0).Text
					oMaster.FSellDate 			= Left(Now(), 10)
					oMaster.FPayType			= "50"
					oMaster.FPaydate			= oMaster.FSellDate
					oMaster.FOrderUserID		= ""

					oMaster.FOrderName			= LEFT(html2db(objMasterOneXML.getElementsByTagName("OrdererName")(0).Text), 28)
					oMaster.FOrderTelNo			= LEFT(html2db(objMasterOneXML.getElementsByTagName("OrdererPhone")(0).Text), 16)
					oMaster.FOrderHpNo			= LEFT(html2db(objMasterOneXML.getElementsByTagName("OrdererCellPhone")(0).Text), 16)
					if Len(CStr(oMaster.FOrderTelNo)) <= 3 then
						oMaster.FOrderTelNo = oMaster.FOrderHpNo
					end if

					oMaster.FOrderEmail			= ""
					oMaster.FReceiveName		= LEFT(html2db(objMasterOneXML.getElementsByTagName("ReceiverName")(0).Text), 28)
					if (objMasterOneXML.getElementsByTagName("ReceiverPhone").length > 0) then
						oMaster.FReceiveTelNo	= LEFT(html2db(objMasterOneXML.getElementsByTagName("ReceiverPhone")(0).Text), 16)
					Else
						oMaster.FReceiveTelNo	= ""
					End If
					oMaster.FReceiveHpNo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("ReceiverCellPhone")(0).Text), 16)
					if Len(CStr(oMaster.FReceiveTelNo)) <= 3 then
						oMaster.FReceiveTelNo = oMaster.FReceiveHpNo
					end if

					if (objMasterOneXML.getElementsByTagName("DeliveryMemo").length > 0) then
						oMaster.Fdeliverymemo	= html2db(objMasterOneXML.getElementsByTagName("DeliveryMemo")(0).Text)
					Else
						oMaster.Fdeliverymemo	= ""
					End If
					oMaster.FdeliverPay 		= objMasterOneXML.getElementsByTagName("SupplyEntrDeliveryFee")(0).Text

					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("ReceiverZipCode")(0).Text)
					oMaster.FReceiveAddr1		= html2db(objMasterOneXML.getElementsByTagName("ReceiverAddr1")(0).Text)
					oMaster.FReceiveAddr2		= html2db(objMasterOneXML.getElementsByTagName("ReceiverAddr2")(0).Text)

					if InStr(oMaster.FReceiveZipCode, "-") = 0 then
						oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
					end if

					'// 주소 수정
					oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
					oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
					tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
					pos = 0
					for k = 0 to 2
						pos = InStr(pos+1, tmpStr, " ")
						if (pos = 0) then
							exit for
						end if
					next

					if (pos > 0) then
						oMaster.FReceiveAddr1 = Left(tmpStr, pos)
						oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
					end if

					oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
					oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)


					Redim oDetailArr(0)
					Set oDetailArr(0) = new COrderDetail
						oDetailArr(0).FdetailSeq			= objMasterOneXML.getElementsByTagName("OrdDtlNo")(0).Text
						oDetailArr(0).FItemID				= objMasterOneXML.getElementsByTagName("SupplyProdCode")(0).Text

						If objMasterOneXML.getElementsByTagName("SizeCode")(0).Text = "단일상품" Then
							optCode = "0000"
							optNm = "단일상품"
						Else
							If InStr(objMasterOneXML.getElementsByTagName("SizeCode")(0).Text, "/") > 0 Then
								strsql = ""
								strsql = strsql & " SELECT itemoption "
								strsql = strsql & " FROM db_item.dbo.tbl_item_option "
								strsql = strsql & " WHERE itemid = '"& objMasterOneXML.getElementsByTagName("SupplyProdCode")(0).Text &"' "
								strsql = strsql & " and optionname = '"& Trim(Split(objMasterOneXML.getElementsByTagName("SizeCode")(0).Text, "/")(1)) &"' "
								rsget.CursorLocation = adUseClient
								rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
								If (Not rsget.EOF) Then
									optCode = rsget("itemoption")
									optNm = Trim(Split(objMasterOneXML.getElementsByTagName("SizeCode")(0).Text, "/")(1))
								Else
									tmpOptionSeq = tmpOptionSeq + 1
									optCode = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
									optNm = objMasterOneXML.getElementsByTagName("SizeCode")(0).Text
								End If
								rsget.Close
							Else
								tmpOptionSeq = tmpOptionSeq + 1
								optCode = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
								optNm = objMasterOneXML.getElementsByTagName("SizeCode")(0).Text
							End If
						End If

						oDetailArr(0).FItemOption			= optCode
						oDetailArr(0).FOutMallItemID		= objMasterOneXML.getElementsByTagName("ProductCode")(0).Text
						oDetailArr(0).FOutMallItemOption	= objMasterOneXML.getElementsByTagName("OptionCode")(0).Text
						oDetailArr(0).FOutMallItemName		= html2db(objMasterOneXML.getElementsByTagName("ItemName")(0).Text)
						oDetailArr(0).FOutMallItemOptionName = html2db(optNm)
						oDetailArr(0).FItemNo				= objMasterOneXML.getElementsByTagName("OrderQty")(0).Text
						oDetailArr(0).Fitemcost				= Clng(objMasterOneXML.getElementsByTagName("OrderAmt")(0).Text  / objMasterOneXML.getElementsByTagName("OrderQty")(0).Text )
						oDetailArr(0).FReducedPrice			= Clng(objMasterOneXML.getElementsByTagName("RealOrderAmt")(0).Text  / objMasterOneXML.getElementsByTagName("OrderQty")(0).Text)
						oDetailArr(0).FOutMallCouponPrice	= 0
						oDetailArr(0).FTenCouponPrice		= 0
						If oDetailArr(0).FItemOption = "" Then
							oDetailArr(0).FItemOption = "0000"
							oDetailArr(0).FOutMallItemOption = "0000"
						End If
						oDetailArr(0).FrequireDetail = ""

						If (SaveOrderToDB(oMaster, oDetailArr) = True) Then
						'	successCnt = successCnt + 1
							strsql = ""
							strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
							strsql = strsql & " VALUES ('"&oMaster.FOutMallOrderSerial&"', '"&oDetailArr(0).FdetailSeq&"', '', 'N', getdate(), 'lfmall')"
							dbget.Execute strSql

							if PlaceProductOrder_lfmall(oMaster.FOutMallOrderSerial, oDetailArr(0).FdetailSeq, sellsite) then
								successCnt = successCnt + 1
							end if
						Else
							rw "--------------------------------------"
							rw objMasterOneXML.getElementsByTagName("OrdNo")(0).Text
							rw objMasterOneXML.getElementsByTagName("SupplyProdCode")(0).Text
							rw "--------------------------------------"
						End If
					Set oDetailArr = Nothing
				next
					response.write "주문입력(" & successCnt & ")" & "<br />"
				GetOrderFrom_lfmall = True
				Set xmlDOM = Nothing
				Set objXML = Nothing
		Else
			if IsAutoScript then
				response.write "ERROR : 통신오류"
			else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			end if

			dbget.close : response.end
		End If
	Set objXML = nothing
End Function

'2019-04-29 김진영..
'# 롯데닷컴 누락건이 발생하여 신규주문으로 긁어왔더라도 한번 더 해당 날짜의 주문확인 리스트를 긁어오기
'# SelOption : 01:미발주(신규주문), 02:발주확인(상품준비)
'# 여기서는 02로 옵션을 주어 긁기
'# 해당 주문코드가 temp table에 존재시 입력하지 않고 temp table에 없다면 역으로 집어 넣는다.
function GetOrderFrom_lotteComConfirmList(selldate)
	dim sellsite : sellsite = "lotteCom"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim tmpStr, pos
	dim i, j, k
	dim found, successCnt, oSql, failCnt, oCnt
	successCnt = 0
	failCnt = 0
	oCnt = 0

	GetOrderFrom_lotteComConfirmList = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 전체내역 가져오기)
	xmlURL = "https://openapi.lotte.com"
	xmlURL = xmlURL + "/openapi/searchNewOrdLstOpenApi.lotte?subscriptionId=" + GetLotteAuthNo() + "&start_date=" + xmlSelldate + "&end_date=" + xmlSelldate + "&SelOption=02"

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 5000,80000,80000,80000
	objXML.Send()

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	else
		objData = BinaryToText(objXML.ResponseBody, "euc-kr")
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML (objData)

	if xmlDOM.getElementsByTagName("Response/Result/OrderInfo").length < 1 then
		''if IsAutoScript then
			response.write "주문확인 내역없음(0)<br />"
		''end if

		GetOrderFrom_lotteComConfirmList = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	masterCnt = xmlDOM.getElementsByTagName("Response/Result/OrderInfo").length
	''if IsAutoScript then
	'	response.write "건수(" & masterCnt & ") " & "<br />"
	''end if

	set objMasterListXML = xmlDOM.selectNodes("/Response/Result/OrderInfo")
	for each objMasterOneXML in objMasterListXML
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite 			= sellsite
		oMaster.FOutMallOrderSerial = Replace(objMasterOneXML.selectSingleNode("OrdNo").text, "-", "")
		oMaster.FSellDate 			= Left(Now(), 10)
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= oMaster.FSellDate
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("CardMemoSndrName").text))
		oMaster.FOrderTelNo			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("OrderTelNo").text))
		oMaster.FOrderHpNo			= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("OrderHpNo").text))
		oMaster.FOrderEmail			= ""
		oMaster.FReceiveName		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvName").text))
		if Len(oMaster.FReceiveName) > 32 then
			oMaster.FReceiveName = oMaster.FOrderName
		end if

		oMaster.FReceiveTelNo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvTel").text))
		oMaster.FReceiveHpNo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvHp").text))

		oMaster.Fdeliverymemo		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DlvMemoCont").text))
		if (oMaster.Fdeliverymemo = "null") then
			oMaster.Fdeliverymemo = ""
		end if

		'// 배송비 안넘어옴
		oMaster.FdeliverPay			= 0

		oMaster.FReceiveZipCode		= html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvPostCode").text))

		'// 우편번호 수정
		if Len(oMaster.FReceiveZipCode) = 4 then
			oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
		end if

		oMaster.FReceiveAddr1 = html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvAddr1").text))
		oMaster.FReceiveAddr2 = html2db(RemoveWhiteSpaceChar(objMasterOneXML.selectSingleNode("DelvInfo/recvAddr2").text))

		if InStr(oMaster.FReceiveZipCode, "-") = 0 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
		oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)


		'// FROM
		'// 견종선택:푸들,사이즈:S (23-30cm),끈타입:타입3 (믹스)
		'// TO
		'// 푸들,S (23-30cm),타입3 (믹스)
		dim regEx
		set regEx = New RegExp
		With regEx
			.Pattern = ",[^:]+:"
			.IgnoreCase = True
			.Global = True
		end with


		set objDetailListXML = objMasterOneXML.selectNodes("ProdInfo")
		for each objDetailOneXML in objDetailListXML
			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objDetailOneXML.selectSingleNode("ProdSeq").text
			oDetailArr(0).FItemID = ""
			oDetailArr(0).FItemOption = "0000"
			oDetailArr(0).FOutMallItemID = objDetailOneXML.selectSingleNode("ProdCode").text
			oDetailArr(0).FOutMallItemOption = "0000"
			oDetailArr(0).FOutMallItemName = html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("ProdName").text))
			oDetailArr(0).FOutMallItemOptionName = html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("prodOption").text))
			if (oDetailArr(0).FOutMallItemOptionName = "null") then
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			'// 롯데닷컴은 업체상품코드/옵션코드 모두 안준다.
			found = False
			sqlStr = ""
			sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
			sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] as m "
			sqlStr = sqlStr & " JOIN db_etcmall.dbo.tbl_lotteAddOption_regItem as r on m.idx = r.midx "
			sqlStr = sqlStr & " WHERE IsNULL(r.LotteGoodNo, r.LotteTmpGoodNo)= '"& oDetailArr(0).FOutMallItemID &"' "
			sqlStr = sqlStr & " and m.mallid = 'lotteCom' "
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			If (Not rsget.EOF) Then
				found = True
				oDetailArr(0).FItemID = rsget("itemid")
				oDetailArr(0).FItemOption = rsget("itemoption")
				oDetailArr(0).FOutMallItemOption = rsget("itemoption")
			End If
			rsget.Close

			if found = False then
		        sqlStr = " select top 2 itemid from db_item.dbo.tbl_lotte_regItem "
		        sqlStr = sqlStr & " where IsNULL(LotteGoodNo,LotteTmpGoodNo)='"& oDetailArr(0).FOutMallItemID &"'"
				rsget.CursorLocation = adUseClient
				rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly

				if (rsget.RecordCount = 1) and (Not rsget.EOF) then
					oDetailArr(0).FItemID = rsget("itemid")
				elseif (oDetailArr(0).FOutMallItemID = "19710092") then
					oDetailArr(0).FItemID = 481915
				end if
				rsget.Close

				if oDetailArr(0).FItemID <> "" then
					if oDetailArr(0).FOutMallItemOptionName <> "" then
						oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, mid(regEx.replace("," & oDetailArr(0).FOutMallItemOptionName, ","), 2, 1000))
					else
						oDetailArr(0).FItemOption = "0000"
					end if
					oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
				else
					'2018-10-02 15:43 김진영 추가..TimeOut등으로 lotte_regItem테이블에 값을 못 넣는 경우 발생
					'oDetailArr(0).FItemID = -1
					sqlStr = " select top 1 itemid from db_item.dbo.tbl_item "
					sqlStr = sqlStr & " where itemname = '"& html2db(RemoveWhiteSpaceChar(objDetailOneXML.selectSingleNode("ProdName").text)) &"'"
					rsget.CursorLocation = adUseClient
					rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
					If (Not rsget.EOF) Then
						oDetailArr(0).FItemID = rsget("itemid")
					End If
					rsget.Close

					if oDetailArr(0).FItemID <> "" then
						if oDetailArr(0).FOutMallItemOptionName <> "" then
							oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, mid(regEx.replace("," & oDetailArr(0).FOutMallItemOptionName, ","), 2, 1000))
						else
							oDetailArr(0).FItemOption = "0000"
						end if
						oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
					Else
						oDetailArr(0).FItemID = -1
					End If
				end if
			end if

			oDetailArr(0).FItemNo = CLng(objDetailOneXML.selectSingleNode("ordQty").text)

			oDetailArr(0).Fitemcost = objDetailOneXML.selectSingleNode("ordPrice").text
			oDetailArr(0).FReducedPrice = objDetailOneXML.selectSingleNode("ordPrice").text
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0

			oDetailArr(0).FrequireDetail = objDetailOneXML.selectSingleNode("GoodsChocDesc").text
			if (oDetailArr(0).FrequireDetail = "null") then
				oDetailArr(0).FrequireDetail = ""
			end if

			oMaster.ForderCsGbn = "0"
			if (objDetailOneXML.selectSingleNode("Exchange").text <> "일반") then
				oMaster.ForderCsGbn = "3"
			end if


			oSql = ""
			oSql = oSql & " SELECT COUNT(*) cnt "
			oSql = oSql & " FROM db_temp.dbo.tbl_xSite_TMPOrder "
			oSql = oSql & " WHERE sellsite = 'lotteCom' "
			oSql = oSql & " and OutMallOrderSerial = '"& oMaster.FOutMallOrderSerial &"' "
			oSql = oSql & " and OrgDetailKey = '"& oDetailArr(0).FdetailSeq &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open oSql, dbget, adOpenForwardOnly, adLockReadOnly
			if not rsget.Eof then
				oCnt	= rsget("cnt")
			end if
			rsget.close

			If oCnt > 0 Then
				failCnt = failCnt + 1
			Else
				if (SaveOrderToDB(oMaster, oDetailArr) = True) then
						successCnt = successCnt + 1
				end if
			End If
			Set oDetailArr = Nothing
		next
		Set oMaster = Nothing
	next

	''if IsAutoScript then
		response.write "주문 누락건 입력(" & successCnt & ")" & "<br />"
	''end if

	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

function GetLotteAuthNo()
	dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd, iisql

	GetLotteAuthNo = ""

	IF application("Svr_Info")="Dev" THEN
		'lotteAPIURL = "http://openapidev.lotte.com"	'' 테스트서버
		lotteAPIURL = "http://openapitest.lotte.com"	'' 테스트서버
		tenBrandCd = "14846"	'텐바(임시)
		tenDlvCd = "513484"		'배송정책코드
		CertPasswd = "1234"		'Dev는 비번 : 1234
	Else
		lotteAPIURL = "https://openapi.lotte.com"		'' 실서버
		tenBrandCd = "155112"	'텐바이텐
		tenDlvCd = "513484"
		CertPasswd = "store101010*!"
	End if
	lottenTenID = "124072"					'텐바이텐ID

	Dim updateAuth, dbAuthNo
	iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
	iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
	iisql = iisql & " where mallid='lotteCom'"&VbCRLF
	iisql = iisql & " and inikey='auth'"
	rsget.CursorLocation = adUseClient
	rsget.Open iisql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
	    dbAuthNo	= rsget("iniVal")
	    updateAuth	= rsget("lastupdate")
	end if
	rsget.close

	If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
		dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", lotteAPIURL & "/openapi/createCertification.lotte?strUserId=" & lottenTenID & "&strPassWd="&CertPasswd&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

			''on Error Resume Next
				GetLotteAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text
				if Err<>0 then
					Response.Write "통신오류(XML)"
					Response.End
				end if
			''on Error Goto 0
				iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
				iisql = iisql & " set iniVal='"&lotteAuthNo&"'"&VbCRLF
				iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
				iisql = iisql & " where mallid='lotteCom'"&VbCRLF
				iisql = iisql & " and inikey='auth'"
				dbget.Execute iisql

			Set xmlDOM = Nothing
		else
			Response.Write "통신오류"
			Response.End
		end if
		Set objXML = Nothing
	Else
		GetLotteAuthNo = dbAuthNo
	End If
end function

function GetOrderDetailFrom_auction1010(detailSeq)
	dim xmlURL, strRst
	dim objXML, xmlDOM, obj
	dim OrderTelNo, OrderHpNo

	xmlURL = "https://api.auction.co.kr"
	xmlURL = xmlURL + "/APIv1/Auctionservice.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	''strRst = strRst + "    <AuthenticationTicket xmlns=""http://www.auction.co.kr/Security"">"
	''strRst = strRst + "      <Value></Value>"
	''strRst = strRst + "    </AuthenticationTicket>"
	strRst = strRst + "    <EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst + "      <Value>" & auctionTicket & "</Value>"
	strRst = strRst + "    </EncryptedTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <GetShippingDetail xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
	strRst = strRst + "      <req OrderNo=""" & detailSeq & """>"
	strRst = strRst + "        <MemberTicket xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">"
	strRst = strRst + "          <Ticket></Ticket>"
	strRst = strRst + "        </MemberTicket>"
	strRst = strRst + "      </req>"
	strRst = strRst + "    </GetShippingDetail>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/GetShippingDetail"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))

	Set obj = xmlDOM.selectSingleNode("Envelope/Body/GetShippingDetailResponse/GetShippingDetailResult/Buyer")
	if Not IsNull(obj.GetAttribute("Tel")) then
		OrderTelNo = obj.attributes.GetNamedItem("Tel").text
	end if
	if Not IsNull(obj.GetAttribute("MobileTel")) then
		OrderHpNo = obj.attributes.GetNamedItem("MobileTel").text
	end if

	GetOrderDetailFrom_auction1010 = OrderTelNo & "|" & OrderHpNo

	Set objXML = Nothing
	Set xmlDOM = Nothing
end function

function GetOrderCouponDetailFrom_auction1010(detailSeq)
	dim xmlURL, strRst
	dim objXML, xmlDOM, obj
	dim OrderTelNo, OrderHpNo

	xmlURL = "https://api.auction.co.kr"
	xmlURL = xmlURL + "/APIv1/Auctionservice.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	''strRst = strRst + "    <AuthenticationTicket xmlns=""http://www.auction.co.kr/Security"">"
	''strRst = strRst + "      <Value>string</Value>"
	''strRst = strRst + "    </AuthenticationTicket>"
	strRst = strRst + "    <EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst + "      <Value>" & auctionTicket & "</Value>"
	strRst = strRst + "    </EncryptedTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <GetSettlementDetail xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
	strRst = strRst + "      <req orderNo=""" & detailSeq & """>"
	strRst = strRst + "        <MemberTicket xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">"
	strRst = strRst + "          <Ticket>string</Ticket>"
	strRst = strRst + "        </MemberTicket>"
	strRst = strRst + "      </req>"
	strRst = strRst + "    </GetSettlementDetail>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/GetSettlementDetail"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))
	''response.write objXML.responseText

	Set obj = xmlDOM.selectSingleNode("Envelope/Body/GetSettlementDetailResponse/GetSettlementDetailResult")

	'// (서비스이용료+정산예정금액) / 수량
	GetOrderCouponDetailFrom_auction1010 = Round(Round(CLng(obj.attributes.GetNamedItem("RemitExpectedMoney").text) + CLng(obj.attributes.GetNamedItem("SellFeeAmount").text)) / CLng(obj.selectSingleNode("OrderBase").attributes.GetNamedItem("AwardQty").text))

	Set objXML = Nothing
	Set xmlDOM = Nothing
end function

function GetOrderDetailConfirmFrom_auction1010(detailSeq)
	dim xmlURL, strRst
	dim objXML, xmlDOM, obj
	dim OrderTelNo, OrderHpNo

	xmlURL = "https://api.auction.co.kr"
	xmlURL = xmlURL + "/APIv1/Auctionservice.asmx"

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	''strRst = strRst + "    <AuthenticationTicket xmlns=""http://www.auction.co.kr/Security"">"
	''strRst = strRst + "      <Value>string</Value>"
	''strRst = strRst + "    </AuthenticationTicket>"
	strRst = strRst + "    <EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst + "      <Value>" & auctionTicket & "</Value>"
	strRst = strRst + "    </EncryptedTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <ConfirmReceivingOrder xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
	strRst = strRst + "      <req OrderNo=""" & detailSeq & """>"
	strRst = strRst + "        <MemberTicket xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">"
	strRst = strRst + "          <Ticket></Ticket>"
	strRst = strRst + "        </MemberTicket>"
	strRst = strRst + "      </req>"
	strRst = strRst + "    </ConfirmReceivingOrder>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/ConfirmReceivingOrder"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		GetOrderDetailConfirmFrom_auction1010 = False

		Set objXML = Nothing
		Set xmlDOM = Nothing
		exit function
	end if

	GetOrderDetailConfirmFrom_auction1010 = True

	Set objXML = Nothing
	Set xmlDOM = Nothing
end function

function GetOrderFrom_auction1010(selldate)
	dim sellsite : sellsite = "auction1010"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0

	GetOrderFrom_auction1010 = False


	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "https://api.auction.co.kr"
	xmlURL = xmlURL + "/APIv1/Auctionservice.asmx"
	''response.write xmlURL

	strRst = ""
	strRst = strRst + "<?xml version=""1.0"" encoding=""utf-8""?>"
	strRst = strRst + "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
	strRst = strRst + "  <soap:Header>"
	''strRst = strRst + "    <AuthenticationTicket xmlns=""http://www.auction.co.kr/Security"">"
	''strRst = strRst + "      <Value></Value>"
	''strRst = strRst + "    </AuthenticationTicket>"
	strRst = strRst + "    <EncryptedTicket xmlns=""http://www.auction.co.kr/Security"">"
	strRst = strRst + "      <Value>" & auctionTicket & "</Value>"
	strRst = strRst + "    </EncryptedTicket>"
	strRst = strRst + "  </soap:Header>"
	strRst = strRst + "  <soap:Body>"
	strRst = strRst + "    <GetPaidOrderList xmlns=""http://www.auction.co.kr/APIv1/AuctionService"">"
	strRst = strRst + "      <req DurationType=""ReceiptDate"" SearchType=""Nothing"" SearchValue="""" CategoryID="""">"
	strRst = strRst + "        <MemberTicket xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"">"
	strRst = strRst + "          <Ticket></Ticket>"
	strRst = strRst + "        </MemberTicket>"
	strRst = strRst + "        <SearchDuration StartDate=""" & selldate & """ EndDate=""" & selldate & """ xmlns=""http://schema.auction.co.kr/Arche.APISvc.xsd"" />"
	strRst = strRst + "      </req>"
	strRst = strRst + "    </GetPaidOrderList>"
	strRst = strRst + "  </soap:Body>"
	strRst = strRst + "</soap:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "Content-Length", LenB(strRst)
	objXML.setRequestHeader "SOAPAction", "http://www.auction.co.kr/APIv1/AuctionService/GetPaidOrderList"
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(Replace(objXML.responseText,"soap:",""))
	''response.write objXML.responseText & "<br /><br />"

	Set obj = Nothing
	If xmlDOM.selectSingleNode("Envelope/Body").firstChild.nodeName = "GetPaidOrderListResponse" Then
		Set obj = xmlDOM.selectNodes("Envelope/Body/GetPaidOrderListResponse")
		If xmlDOM.selectSingleNode("Envelope/Body/GetPaidOrderListResponse").firstChild.nodeName = "GetPaidOrderListResult" Then
			Set obj = xmlDOM.selectNodes("Envelope/Body/GetPaidOrderListResponse/GetPaidOrderListResult")
		else
			Set obj = Nothing
		end if
	end if

	if obj is Nothing then
		if IsAutoScript then
			response.write "내역없음 : 종료"
		end if

		GetOrderFrom_auction1010 = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	masterCnt = (xmlDOM.selectNodes("Envelope/Body/GetPaidOrderListResponse/GetPaidOrderListResult/PaidOrder").length)
	''response.write masterCnt

	if masterCnt = 0 then
		if IsAutoScript then
			response.write "내역없음<br />"
		end if

		GetOrderFrom_auction1010 = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.selectNodes("Envelope/Body/GetPaidOrderListResponse/GetPaidOrderListResult/PaidOrder")
	masterCnt = objMasterListXML.length

	''if IsAutoScript then
		response.write "건수(" & masterCnt & ") " & "<br />"
	''end if

	for i = 0 to masterCnt - 1
		set objMasterOneXML = objMasterListXML.item(i)
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite 			= sellsite
		oMaster.FOutMallOrderSerial = objMasterOneXML.attributes.GetNamedItem("PayNo").value
		oMaster.FSellDate 			= Left(objMasterOneXML.attributes.GetNamedItem("ReceiptDate").text, 10)
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= oMaster.FSellDate
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= html2db(Trim(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("BuyerName").text))
		oMaster.FOrderTelNo			= ""
		oMaster.FOrderHpNo			= ""
		oMaster.FOrderEmail			= ""
		oMaster.FReceiveName		= html2db(Trim(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("Name").value))
		oMaster.FReceiveTelNo		= html2db(Trim(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("Tel").value))
		oMaster.FReceiveHpNo		= html2db(Trim(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("MobileTel").value))

		oMaster.Fdeliverymemo		= html2db(objMasterOneXML.attributes.GetNamedItem("DeliveryRemark").value)
		oMaster.FdeliverPay			= CLng(objMasterOneXML.attributes.GetNamedItem("DeliveryFeeAmount").value)

		oMaster.FReceiveZipCode		= html2db(Trim(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("PostNo").value))

		'// 우편번호 수정
		if Len(oMaster.FReceiveZipCode) = 4 then
			oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
		end if

		if Len(oMaster.FReceiveZipCode) <= 5 and objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("AddressRoadName").text <> "" then
			'// 도로명주소
			oMaster.FReceiveAddr1		= html2db(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("AddressRoadName").value)
			oMaster.FReceiveAddr2		= ""
		else
			'// 구주소
			oMaster.FReceiveAddr1		= html2db(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("AddressPost").value)
			oMaster.FReceiveAddr2		= html2db(objMasterOneXML.getElementsByTagName("AddressBase")(0).attributes.GetNamedItem("AddressDetail").value)
		end if

		if Len(oMaster.FReceiveZipCode) > 4 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
		oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)

		redim oDetailArr(0)
		Set oDetailArr(0) = new COrderDetail
		oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("OrderNo").text
		oDetailArr(0).FItemID = objMasterOneXML.attributes.GetNamedItem("ItemCode").text
		oDetailArr(0).FItemOption = objMasterOneXML.attributes.GetNamedItem("SellerStockCode").text
		oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("ItemID").text
		oDetailArr(0).FOutMallItemOption = objMasterOneXML.attributes.GetNamedItem("SellerStockCode").text
		oDetailArr(0).FOutMallItemName = html2db(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("ItemName").text)
		oDetailArr(0).FOutMallItemOptionName = html2db(objMasterOneXML.attributes.GetNamedItem("RequestOption").text)

		oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardQty").text)

		oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardAmount").text) / oDetailArr(0).FItemNo
		oDetailArr(0).FReducedPrice = oDetailArr(0).Fitemcost
		oDetailArr(0).FOutMallCouponPrice = 0
		oDetailArr(0).FTenCouponPrice = 0

		if oDetailArr(0).FItemOption = "" then
			oDetailArr(0).FItemOption = "0000"
			oDetailArr(0).FOutMallItemOption = "0000"
		end if

		if (oDetailArr(0).FItemOption <> "0000") then
			if Not GetCheckItemOptionValid(oDetailArr(0).FItemID, oDetailArr(0).FItemOption) then
				'// 잘못된 옵션.
				tmpOptionSeq = tmpOptionSeq + 1
				oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
				oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
			end if
		end if

		'// 주문제작문구 수정
		if InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") <> 0 then
			oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FOutMallItemOptionName, InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") + Len("텍스트를 입력하세요"), 1000)
			if left(oDetailArr(0).FrequireDetail, 1) = "：" then
				oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FrequireDetail, 2, 1000)
			end if
		end if

		'// 구매자 연락처
		tmpStr = GetOrderDetailFrom_auction1010(oDetailArr(0).FdetailSeq)
		tmpStr = Split(tmpStr, "|")
		oMaster.FOrderTelNo			= tmpStr(0)
		oMaster.FOrderHpNo			= tmpStr(1)

		'// 실판매가
		oDetailArr(0).FReducedPrice = GetOrderCouponDetailFrom_auction1010(oDetailArr(0).FdetailSeq)

		if (SaveOrderToDB(oMaster, oDetailArr) = True) then
			if GetOrderDetailConfirmFrom_auction1010(oDetailArr(0).FdetailSeq) then
				successCnt = successCnt + 1
			end if
		end if

		Set oMaster = Nothing
		Set oDetailArr = Nothing
	next

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

end function

'// /outmall/gseshop/gseshopItemcls.asp 참조
CONST CGSShopCompanyCode = 1003890	'' 협력사코드
function GetOrderFrom_gseshop(selldate)
	dim sellsite : sellsite = "gseshop"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData, strParam
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst, buf

	GetOrderFrom_gseshop = False


	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	'// tnsType : 주문구분(주문/반품 : S, 취소 : C)
	'// 개발 : test1 운영 : ecb2b
	if (application("Svr_Info") = "Dev") then
		xmlURL = "http://test1.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=S"
	else
		xmlURL = "http://ecb2b.gsshop.com/SupSendOrderInfo.gs?supCd=" + CStr(CGSShopCompanyCode) + "&sdDt=" + CStr(xmlSelldate) + "&tnsType=S"
	end if
	''response.write xmlURL & "<br />"
	''dbget.close : response.end


	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.setTimeouts 50000,90000,90000,90000
	objXML.send()

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	'// 전송요청만 한다.(XML 수신X)

end function

Function GetOrderFrom_gseshopNew(selldate)
	dim sellsite : sellsite = "gseshop"
	dim xmlURL, xmlSelldate, obj
	dim objXML, strObj, objData, jParam, requireDetailObj, requireDetail
	dim i, j, k, strsql
	dim successCnt : successCnt = 0
	Dim returnCode, resultMsg
	Dim apiUrl

	GetOrderFrom_gseshopNew = False

	'// =======================================================================
	'// 날짜형식
	xmlSelldate = Replace(selldate, "-", "")
	If (application("Svr_Info") = "Dev") Then
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
		'apiUrl = "http://testapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	Else
		apiUrl = "http://realapi.gsshop.com/b2b/SupSendOrderInfo.gs"
	End If

	'// =======================================================================
	Set obj = jsObject()
		obj("sender") = "TBT"					'전송자 (협력사별 부여되는 ID로 GS에서 제공)
		obj("receiver") = "GS SHOP"				'수신자 | GS SHOP
		obj("documentId") = "ORDINF"			'문서ID | ORDINF
		obj("processType") = "S"				'전송구분 | A:전체, S:주문/반품, C:취소
		obj("supCd") = ""&CGSShopCompanyCode&""	'협력사코드	| (협력사번호 GS에서 제공)
		obj("sdDt") = xmlSelldate				'조회일자
		jParam = obj.jsString
	Set obj = nothing

	'// 데이타 가져오기
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", apiUrl, false
		objXML.setRequestHeader "Accept", "application/json"
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(jParam)
' rw jParam
' rw "----------------------------"
'rw BinaryToText(objXML.ResponseBody,"utf-8")
' response.end

		If objXML.Status <> "200" Then
			response.write "ERROR : 통신오류"
			dbget.close : response.end
		Else
			objData = BinaryToText(objXML.ResponseBody,"utf-8")
		End If

	'// =======================================================================
	'// Json 파싱
		Set strObj = JSON.parse(objData)
			returnCode		= strObj.resultCd
			resultMsg		= strObj.resultMsg
			If returnCode <> "S" Then
				response.write "ERROR : 오류" & resultMsg
				dbget.close : response.end
			End If
		Set strObj = nothing
	Set objXML = nothing
'response.end
End Function

CONST sabangnetAPIURL = "http://r.sabangnet.co.kr"
CONST sabangnetID = "tenbyten"
CONST sabangnetAPIKEY = "PTxNV3d9CXPXBNu60X72EbSNYTJd5955b"
CONST sabangnetWapiURL = "http://wapi.10x10.co.kr"
function GetOrderFrom_sabangnet(selldate)
	dim sellsite : sellsite = "sabangnet"
	dim xmlURL, xmlSelldate
	Dim fso,tFile, istrParam, dataURL
	dim objXML, xmlDOM, objData, strParam
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	''dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst, buf
	dim tmpOptionSeq, sqlStr

	GetOrderFrom_sabangnet = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2018-05-11"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = sabangnetAPIURL
	xmlURL = xmlURL + "/RTL_API/xml_order_info.html"
	''response.write xmlURL

	Dim opath : opath = "/outmall/sabangnet/sabangnetXML/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	CALL CheckFolderCreate(defaultPath)
	Dim fileName

	fileName = "GetOrder" &"_"& getCurrDateTimeFormat&".xml"

	response.write selldate
	istrParam = getSabangnetOrderParameter(selldate)
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.WriteLine istrParam
		Set tFile = nothing
	Set fso = nothing

	dataURL = "?xml_url="&sabangnetWapiURL&opath&FileName

	''On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "" & sabangnetAPIURL&"/RTL_API/xml_order_info.html" & dataURL, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send(istrParam)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				objData = BinaryToText(objXML.ResponseBody, "euc-kr")
				If session("ssBctID")="kjy8517" Then
					response.write objData
				End If
				xmlDOM.LoadXML objData

				'2019-07-02 김진영..사방넷 xml 저장
				Dim cfso,ctFile
				Dim copath : copath = "/outmall/order/lib/"
				Dim cdefaultPath : cdefaultPath = server.mappath(copath) + "\"
				Dim cfileName : cfileName = "testSabang"&left(now(),10)&"_"&timer()&".xml"
				Set cfso = CreateObject("Scripting.FileSystemObject")
					Set ctFile = cfso.CreateTextFile(cdefaultPath & cfileName )
						ctFile.WriteLine objData
					Set ctFile = nothing
				Set cfso = nothing
				'2019-07-02 김진영..사방넷 xml 저장 끝

				Set objMasterListXML = xmlDOM.selectNodes("SABANG_ORDER_LIST/DATA")
				masterCnt = objMasterListXML.length

				''if IsAutoScript then
					response.write "건수(" & masterCnt & ") " & "<br />"
				''end if

				if (masterCnt > 0) then
					tmpOptionSeq = 0
					for i = 0 to masterCnt - 1
						set objMasterOneXML = objMasterListXML.item(i)
						Set oMaster = new COrderMasterItem

						oMaster.FSellSite 			= "unknown"
						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "GS fresh") then
							oMaster.FSellSite 			= "gsisuper"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "LG패션") Then
							oMaster.FSellSite 			= "LFmall"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "YES24") Then
							oMaster.FSellSite 			= "yes24"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "ALPHA MALL") Then
							oMaster.FSellSite 			= "alphamall"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "오늘의집") Then
							oMaster.FSellSite 			= "ohou1010"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "스마트스토어") Then
							oMaster.FSellSite 			= "wadsmartstore"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굳닷컴") Then
							oMaster.FSellSite 			= "casamia_good_com"
						' ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "무신사") Then
						' 	oMaster.FSellSite 			= "musinsa22"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then
							oMaster.FSellSite 			= "kakaostore"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "Wconcept") Then
							oMaster.FSellSite 			= "wconcept1010"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "자연이랑") Then
							oMaster.FSellSite 			= "withnature1010"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿샵") Then
							oMaster.FSellSite 			= "goodshop1010"
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿웨어몰") Then
							oMaster.FSellSite 			= "goodwearmall10"
						End If

						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then
							oMaster.FOutMallOrderSerial = ""
						Else
							oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("ORDER_ID")(0).text
						End If
						
						oMaster.FshoplinkerOrderID = objMasterOneXML.getElementsByTagName("IDX")(0).text

						'// 20180511 => 2018-05-11
						oMaster.FSellDate 			= Left(objMasterOneXML.getElementsByTagName("REG_DATE")(0).text,8)
						oMaster.FSellDate			= Left(oMaster.FSellDate, 4) & "-" & Right(Left(oMaster.FSellDate,6), 2) & "-" & Right(oMaster.FSellDate, 2)
						oMaster.FPayType			= "50"
						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then
							oMaster.FPaydate		= objMasterOneXML.getElementsByTagName("ORDER_DATE")(0).text
							oMaster.FPaydate 		= Left(oMaster.FPaydate,4) & "-" & Mid(oMaster.FPaydate,5,2) & "-" & Mid(oMaster.FPaydate,7,2) & " " & Mid(oMaster.FPaydate,9,2) & ":" & Mid(oMaster.FPaydate,11,2) & ":" & Mid(oMaster.FPaydate,13,2)
						Else
							oMaster.FPaydate		= oMaster.FSellDate
						End If
						oMaster.FOrderUserID		= ""
						oMaster.FOrderName 			= html2db(objMasterOneXML.getElementsByTagName("USER_NAME")(0).text)
						oMaster.FOrderTelNo 		= html2db(objMasterOneXML.getElementsByTagName("USER_TEL")(0).text)
						oMaster.FOrderHpNo 			= html2db(objMasterOneXML.getElementsByTagName("USER_CEL")(0).text)
						oMaster.FOrderEmail			= ""
						oMaster.FReceiveName 		= html2db(objMasterOneXML.getElementsByTagName("RECEIVE_NAME")(0).text)
						oMaster.FReceiveTelNo 		= html2db(objMasterOneXML.getElementsByTagName("RECEIVE_TEL")(0).text)
						oMaster.FReceiveHpNo 		= html2db(objMasterOneXML.getElementsByTagName("RECEIVE_CEL")(0).text)

						oMaster.Fdeliverymemo 		= html2db(objMasterOneXML.getElementsByTagName("DELV_MSG")(0).text)
						oMaster.FdeliverPay			= CLng(objMasterOneXML.getElementsByTagName("DELV_COST")(0).text)

						oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("RECEIVE_ZIPCODE")(0).text)

						'// 우편번호 수정
						if Len(oMaster.FReceiveZipCode) = 4 then
							oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
						end if

						oMaster.FReceiveAddr1		= objMasterOneXML.getElementsByTagName("RECEIVE_ADDR")(0).text

						tmpStr = oMaster.FReceiveAddr1
						pos = 0
						for k = 0 to 2
							pos = InStr(pos+1, tmpStr, " ")
							if (pos = 0) then
								exit for
							end if
						next

						if (pos > 0) then
							oMaster.FReceiveAddr1 = Left(tmpStr, pos)
							oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
						end if

						oMaster.FReceiveAddr1 = html2db(Trim(oMaster.FReceiveAddr1))
						oMaster.FReceiveAddr2 = html2db(Trim(oMaster.FReceiveAddr2))


						redim oDetailArr(0)
						Set oDetailArr(0) = new COrderDetail
						'2019-07-02 김진영 추가..LFmall은 MALL_ORDER_SEQ 값이 공란임..IDX로 대체
						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "LG패션") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "스마트스토어") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굳닷컴") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "자연이랑") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿샵") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿웨어몰") Then
							oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("IDX")(0).text
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then			'2022-08-05 김진영..주문번호 변경..
							oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("MALL_ORDER_SEQ")(0).text	'카카오톡스토어의 결제번호인듯
							oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("ORDER_ID")(0).text				'카카오톡스토어의 주문번호인듯
						Else
							oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("MALL_ORDER_SEQ")(0).text
						End If

						oDetailArr(0).FItemID = objMasterOneXML.getElementsByTagName("COMPAYNY_GOODS_CD")(0).text

						If Len(oDetailArr(0).FItemID) < 3 Then
							Select Case objMasterOneXML.getElementsByTagName("PRODUCT_NAME")(0).text
								Case "[Peanuts] 머그컵_스누피와 친구들 (6종)"										oDetailArr(0).FItemID = "2785591"
								Case "[Peanuts] 스누피 레트로 브런치 플레이트"										oDetailArr(0).FItemID = "3471393"
								Case "[Peanuts] 스누피 레트로 컵세트(2인조)"										oDetailArr(0).FItemID = "3471386"
								Case "[Peanuts] 스누피 코지 밥솥"													oDetailArr(0).FItemID = "4572790"
								Case "[Peanuts] 스누피 코지 시리얼볼 세트"											oDetailArr(0).FItemID = "4442989"
								Case "[Peanuts] 스누피 코지 전기포트"												oDetailArr(0).FItemID = "4572966"
								Case "[Peanuts] 스누피 코지 타이머 와플/샌드위치 메이커"							oDetailArr(0).FItemID = "4495193"
								Case "[Peanuts] 스누피 코지 타이머 와플/샌드위치 메이커_플레이트"					oDetailArr(0).FItemID = "4495197"
								Case "[Peanuts] 스누피 코지 타이머 와플/샌드위치 메이커_플레이트"					oDetailArr(0).FItemID = "4495197"
								Case "[Peanuts] 스누피 코지 토스터기"												oDetailArr(0).FItemID = "4572867"
								Case "[Peanuts] 스누피 코지 플레이트 (2P)"											oDetailArr(0).FItemID = "4442990"
								Case "[Peanuts] 스누피 터치 무드등"													oDetailArr(0).FItemID = "3313868"
								Case "[Peanuts] 스누피 테이블 (M)"													oDetailArr(0).FItemID = "3019218"
								Case "[Peanuts] 스누피 테이블 (S)"													oDetailArr(0).FItemID = "3019191"
								Case "[Peanuts] 스누피 트레이"														oDetailArr(0).FItemID = "3074724"
								Case "[Peanuts] 스누피 페이스 쿠션"													oDetailArr(0).FItemID = "2953019"
								Case "[Peanuts] 스누피와 친구들 사각접시"											oDetailArr(0).FItemID = "3671855"
								Case "[Peanuts] 스터디 텀블러 (6종)"												oDetailArr(0).FItemID = "4460500"
								Case "[Peanuts] 프렌즈 유리컵"														oDetailArr(0).FItemID = "3649588"
								Case "[라이브구성] 스누피 코지 시리얼볼세트+접시2P 구성"							oDetailArr(0).FItemID = "4730390"
								Case "[라이브구성] 스누피 코지 와플/샌드위치 메이커 (+플레이트 1종 추가증정)"		oDetailArr(0).FItemID = "4950717"
								Case "[라이브구성] 스누피 코지가전 3종세트 (밥솥+토스터기+전기포트)"				oDetailArr(0).FItemID = "4883156"
								Case "[라이브방송] 스누피 코지 와플/샌드위치 메이커 (플레이트4종증정)"				oDetailArr(0).FItemID = "4950717"
								Case "Peanuts 스누피 코지 시리얼볼+접시세트"										oDetailArr(0).FItemID = "4730390"
							End Select
						End If

						oDetailArr(0).FItemOption = ""
						oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("MALL_PRODUCT_ID")(0).text
						oDetailArr(0).FOutMallItemOption = ""
						oDetailArr(0).FOutMallItemName = objMasterOneXML.getElementsByTagName("PRODUCT_NAME")(0).text
						oDetailArr(0).FOutMallItemOptionName = objMasterOneXML.getElementsByTagName("SKU_VALUE")(0).text
						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "스마트스토어") Then
							oDetailArr(0).FOutMallItemOptionName = Trim(Replace(oDetailArr(0).FOutMallItemOptionName, "옵션:", ""))
						ElseIf (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then
							oDetailArr(0).FOutMallItemOptionName = Trim(Replace(oDetailArr(0).FOutMallItemOptionName, "옵션/", ""))
						End If
						oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("P_EA")(0).text)

						If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "YES24") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "Wconcept") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "스마트스토어") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "ALPHA MALL") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "오늘의집") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굳닷컴") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "자연이랑") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿샵") OR (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "굿웨어몰") Then
							If objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "ALPHA MALL" Then
								'알파몰은 상품 단가가 안 넘어옴..
								'tenbyten1010는 PB상품 ID임..20%, 일반은 15%
								If objMasterOneXML.getElementsByTagName("MALL_USER_ID")(0).text = "tenbyten1010" Then
									oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("MALL_WON_COST")(0).text) / 0.8
								Else
									oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("MALL_WON_COST")(0).text) / 0.85
								End If
							Else
								oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("SALE_COST")(0).text)
							End If

							If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "오늘의집") Then
								If CLng(objMasterOneXML.getElementsByTagName("DELV_COST")(0).text) > 0 Then
									oDetailArr(0).FReducedPrice = Clng( (objMasterOneXML.getElementsByTagName("PAY_COST")(0).text - CLng(objMasterOneXML.getElementsByTagName("DELV_COST")(0).text)) / objMasterOneXML.getElementsByTagName("P_EA")(0).text)
								Else
									oDetailArr(0).FReducedPrice = Clng(objMasterOneXML.getElementsByTagName("PAY_COST")(0).text / objMasterOneXML.getElementsByTagName("P_EA")(0).text)
								End If
							Else
								oDetailArr(0).FReducedPrice = Clng(objMasterOneXML.getElementsByTagName("PAY_COST")(0).text / objMasterOneXML.getElementsByTagName("P_EA")(0).text)
							End If

							If (objMasterOneXML.getElementsByTagName("MALL_ID")(0).text = "카카오톡스토어") Then
								oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("SALE_COST")(0).text)
								oDetailArr(0).FReducedPrice = Clng(objMasterOneXML.getElementsByTagName("PAY_COST")(0).text  / objMasterOneXML.getElementsByTagName("P_EA")(0).text )
							End If

							oDetailArr(0).FOutMallCouponPrice = 0
							oDetailArr(0).FTenCouponPrice = 0
							If (oDetailArr(0).FOutMallItemOptionName = "" OR oDetailArr(0).FOutMallItemOptionName="single type" OR oDetailArr(0).FOutMallItemOptionName="단품" OR oDetailArr(0).FOutMallItemOptionName="- 단품") then
								oDetailArr(0).FItemOption = "0000"
								oDetailArr(0).FOutMallItemOption = "0000"
							Else
								oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, oDetailArr(0).FOutMallItemOptionName)
								If (oDetailArr(0).FItemOption = "0000") then
									tmpOptionSeq = tmpOptionSeq + 1
									oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
									oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
								End If
							End If
						Else
							'2019-05-31 16:21 김진영 상품가격 변경..LFmall 경우 MALL_WON_COST이 0원으로 나옴
							If Clng(objMasterOneXML.getElementsByTagName("MALL_WON_COST")(0).text) < 1 Then
								oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("SALE_COST")(0).text)
							Else
								oDetailArr(0).Fitemcost = Clng(objMasterOneXML.getElementsByTagName("MALL_WON_COST")(0).text)
							End If
							oDetailArr(0).FReducedPrice = Clng(objMasterOneXML.getElementsByTagName("SALE_COST")(0).text)
							oDetailArr(0).FOutMallCouponPrice = 0
							oDetailArr(0).FTenCouponPrice = 0

							if (oDetailArr(0).FOutMallItemOptionName = "") then
								oDetailArr(0).FItemOption = "0000"
								oDetailArr(0).FOutMallItemOption = "0000"
							else
								oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, oDetailArr(0).FOutMallItemOptionName)

								if (oDetailArr(0).FItemOption = "0000") then
									tmpOptionSeq = tmpOptionSeq + 1
									oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
									oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
								end if
							end if
						End If

						'// 주문제작문구 수정
						''if InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") <> 0 then
						''	oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FOutMallItemOptionName, InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") + Len("텍스트를 입력하세요"), 1000)
						''	if left(oDetailArr(0).FrequireDetail, 1) = "：" then
						''		oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FrequireDetail, 2, 1000)
						''	end if
						''end if

						''response.write "oMaster.FdeliverPay : " & oMaster.FdeliverPay & "<br />"
						''dbget.close : response.end

						if (oMaster.FSellSite = "unknown") then
							response.write "미지정 판매처 : " & objMasterOneXML.getElementsByTagName("MALL_ID")(0).text & "<br />"
						else
							if (SaveOrderToDB(oMaster, oDetailArr) = True) then
								''if GetOrderDetailConfirmFrom_auction1010(oDetailArr(0).FdetailSeq) then
									successCnt = successCnt + 1
								''end if
							end if
						end if
					next
				end if

				''if IsAutoScript then
					response.write "주문입력(" & successCnt & ")" & "<br />"
				''end if

			Set xmlDOM = Nothing
		Else
			if IsAutoScript then
				response.write "ERROR : 통신오류"
			else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			end if
		End If
	Set objXML = Nothing
	''On Error Goto 0
If session("ssBctID") <> "kjy8517" Then
	Call DelAPITMPFile(sabangnetWapiURL&opath&FileName)
End If
	''dbget.close : response.end

end function

Function getSabangnetOrderParameter(selldate)
	dim strRst, reqFields

	'// ORDER_STATUS
	'// 주문(배송)상태 코드	상태명
	'// 001	신규주문
	'// 002	주문확인
	'// 003	출고대기
	'// 004	출고완료
	'// 006	배송보류
	'// 007	취소접수
	'// 008	교환접수
	'// 009	반품접수
	'// 010	취소완료
	'// 011	교환완료
	'// 012	반품완료
	'// 021	교환발송준비
	'// 022	교환발송완료
	'// 023	교환회수준비
	'// 024	교환회수완료
	'// 025	반품회수준비
	'// 026	반품회수완료
	'// 999	폐기

	'// IDX	주문번호(사방넷)
	'// ORDER_ID	주문번호(쇼핑몰)
	'// MALL_ID	쇼핑몰명
	'// MALL_USER_ID	쇼핑몰ID
	'// ORDER_STATUS	주문상태
	'// USER_ID	주문자ID
	'// USER_NAME	주문자명
	'// USER_TEL	주문자전화번호
	'// USER_CEL	주문자핸드폰번호
	'// USER_EMAIL	주문자이메일주소
	'// RECEIVE_TEL	수취인전화번호
	'// RECEIVE_CEL	수취인핸드폰번호
	'// RECEIVE_EMAIL	수취인이메일주소
	'// DELV_MSG	배송메세지
	'// RECEIVE_NAME	수취인명
	'// RECEIVE_ZIPCODE	수취인우편번호
	'// RECEIVE_ADDR	수취인주소
	'// TOTAL_COST	주문금액
	'// PAY_COST	결제금액
	'// ORDER_DATE	주문일자
	'// PARTNER_ID	매입처명
	'// DPARTNER_ID	물류처명
	'// MALL_PRODUCT_ID	상품코드(쇼핑몰)
	'// PRODUCT_ID	품번코드(사방넷)
	'// SKU_ID	단품코드(사방넷)
	'// P_PRODUCT_NAME	상품명(확정)
	'// P_SKU_VALUE	옵션(확정)
	'// PRODUCT_NAME	상품명(수집)
	'// SALE_COST	판매가(수집)
	'// MALL_WON_COST	공급단가
	'// WON_COST	원가
	'// SKU_VALUE	옵션(수집)
	'// SALE_CNT	수량
	'// DELIVERY_METHOD_STR	배송구분
	'// DELV_COST	배송비(수집)
	'// COMPAYNY_GOODS_CD	자체상품코드
	'// SKU_ALIAS	옵션별칭
	'// BOX_EA	EA(상품)
	'// JUNG_CHK_YN	정산대조확인여부
	'// MALL_ORDER_SEQ	주문순번
	'// MALL_ORDER_ID	부주문번호
	'// ETC_FIELD3	수정된 수집옵션명
	'// ORDER_GUBUN	주문구분
	'// P_EA	EA(확정)
	'// REG_DATE	수집일자
	'// ORDER_ETC_1	자사몰필드1
	'// ORDER_ETC_2	자사몰필드2
	'// ORDER_ETC_3	자사몰필드3
	'// ORDER_ETC_4	자사몰필드4
	'// ORDER_ETC_5	자사몰필드5
	'// ORDER_ETC_6	자사몰필드6
	'// ORDER_ETC_7	자사몰필드7
	'// ORDER_ETC_8	자사몰필드8
	'// ORDER_ETC_9	자사몰필드9
	'// ORDER_ETC_10	자사몰필드10
	'// ORDER_ETC_11	자사몰필드11
	'// ORDER_ETC_12	자사몰필드12
	'// ORDER_ETC_13	자사몰필드13
	'// ORDER_ETC_14	자사몰필드14
	'// ord_field2	세트분리주문구분
	'// copy_idx	원주문번호(사방넷)
	'// GOODS_NM_PR	출력상품명
	'// GOODS_KEYWORD	상품약어
	'// ORD_CONFIRM_DATE	주문 확인일자
	'// RTN_DT	반품 완료일자
	'// CHNG_DT	교환 완료일자
	'// DELIVERY_CONFIRM_DATE	출고 완료일자
	'// CANCEL_DT	취소 완료일자
	'// CLASS_CD1	대분류코드
	'// CLASS_CD2	중분류코드
	'// CLASS_CD3	소분류코드
	'// CLASS_CD4	세분류코드
	'// BRAND_NM	브랜드명
	'// DELIVERY_ID	택배사코드
	'// INVOICE_NO	송장번호
	'// HOPE_DELV_DATE	배송희망일자
	'// FLD_DSP	주문엑셀용
	'// INV_SEND_MSG	운송장 전송 결과 메시지
	'// MODEL_NO	모델NO
	'// SET_GUBUN	상품구분
	'// ETC_MSG	기타메세지
	'// DELV_MSG1	배송메세지
	'// MUL_DELV_MSG	물류메세지
	'// BARCODE	바코드
	'// INV_SEND_DM	송장전송일자

	reqFields = "IDX|ORDER_ID|MALL_ID|MALL_USER_ID|ORDER_STATUS|USER_ID|USER_NAME|USER_TEL|USER_CEL|USER_EMAIL|RECEIVE_TEL|RECEIVE_CEL|RECEIVE_EMAIL"
	reqFields = reqFields + "|DELV_MSG|RECEIVE_NAME|RECEIVE_ZIPCODE|RECEIVE_ADDR|TOTAL_COST|PAY_COST|ORDER_DATE|PARTNER_ID|DPARTNER_ID|MALL_PRODUCT_ID|PRODUCT_ID|SKU_ID"
	reqFields = reqFields + "|P_PRODUCT_NAME|P_SKU_VALUE|PRODUCT_NAME|SALE_COST|MALL_WON_COST|WON_COST|SKU_VALUE|SALE_CNT|DELIVERY_METHOD_STR|DELV_COST|COMPAYNY_GOODS_CD"
	reqFields = reqFields + "|SKU_ALIAS|BOX_EA|JUNG_CHK_YN|MALL_ORDER_SEQ|MALL_ORDER_ID|ETC_FIELD3|ORDER_GUBUN|P_EA|REG_DATE"
	reqFields = reqFields + "|ORDER_ETC_1|ORDER_ETC_2|ORDER_ETC_3|ORDER_ETC_4|ORDER_ETC_5|ORDER_ETC_6|ORDER_ETC_7|ORDER_ETC_8|ORDER_ETC_9|ORDER_ETC_10|ORDER_ETC_11|ORDER_ETC_12|ORDER_ETC_13|ORDER_ETC_14"
	reqFields = reqFields + "|ord_field2|copy_idx|GOODS_NM_PRGOODS_KEYWORD|ORD_CONFIRM_DATE|RTN_DT|CHNG_DT|DELIVERY_CONFIRM_DATE|CANCEL_DT"
	reqFields = reqFields + "|CLASS_CD1|CLASS_CD2|CLASS_CD3|CLASS_CD4|BRAND_NM|DELIVERY_ID|INVOICE_NO|HOPE_DELV_DATE|FLD_DSP|INV_SEND_MSG|MODEL_NO|SET_GUBUN"
	reqFields = reqFields + "|ETC_MSG|DELV_MSG1|MUL_DELV_MSG|BARCODE|INV_SEND_DM"

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	strRst = strRst & "<SABANG_ORDER_LIST>"
	strRst = strRst & "	<HEADER>"
	strRst = strRst & "		<SEND_COMPAYNY_ID>"&sabangnetID&"</SEND_COMPAYNY_ID>"				'#사방넷 로그인 아이디
	strRst = strRst & "		<SEND_AUTH_KEY>"&sabangnetAPIKEY&"</SEND_AUTH_KEY>"					'#사방넷에서 발급 받은 인증키
	strRst = strRst & "		<SEND_DATE>"&Replace(Date(), "-", "")&"</SEND_DATE>"				'#전송일자 | YYYYMMDD
	strRst = strRst & "	</HEADER>"
	strRst = strRst & "	<DATA>"
	strRst = strRst & "		<ORD_ST_DATE>"&Replace(selldate, "-", "")&"</ORD_ST_DATE>"	'// 사방넷에 수집된 주문의 수집일자 기준의 검색조건
	strRst = strRst & "		<ORD_ED_DATE>"&Replace(selldate, "-", "")&"</ORD_ED_DATE>"
	strRst = strRst & "		<ORD_FIELD><![CDATA[" & reqFields & "]]></ORD_FIELD>"
	strRst = strRst & "		<ORDER_STATUS><![CDATA['001]]></ORDER_STATUS>"							'// 반드시 따옴표 ' << 이거로 시작해야 함
	strRst = strRst & "	</DATA>"
	strRst = strRst & "</SABANG_ORDER_LIST>"

	getSabangnetOrderParameter = strRst
end Function

Function getCurrDateTimeFormat()
	Dim nowtimer : nowtimer= timer()
	getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile = Server.CreateObject("Scripting.FileSystemObject")
	If NOT objfile.FolderExists(sFolderPath) Then
		objfile.CreateFolder sFolderPath
	End If
	Set objfile = Nothing
End Sub

Function DelAPITMPFile(iFileURI)
	Dim iFullPath
	iFullPath = server.mappath(replace(iFileURI,"http://wapi.10x10.co.kr",""))

	Dim FSO, iFile
	Set FSO = CreateObject("Scripting.FileSystemObject")
		Set iFile = FSO.GetFile(iFullPath)
			If (iFile <> "") Then iFile.Delete
		Set iFile = Nothing
	Set FSO = Nothing
End Function

public function RequestArrayToArray(reqObj)
	dim obj, objArr()
	dim i
	Set obj = reqObj

	ReDim objArr(obj.Count - 1)

	For i = 0 To obj.Count - 1
		objArr(i) = obj(i+1)
	Next

	RequestArrayToArray = objArr
end function

function RemovePrecedingZero(str)
	dim result
	result = str
	do while (Left(result, 1) = "0")
		result = Mid(result, 2, 1000)
	loop
	RemovePrecedingZero = result
end function

public function GetOrderFrom_gseshop_Recv()
	dim sellsite : sellsite = "gseshop"
	dim oMaster, oDetailArr(), sqlStr, isOptAddGS
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim successCnt : successCnt = 0
	dim i, j, k, pos
	dim tmpStr, strSql, iAssignedRow
	dim orgOrdNo, orgOrdItemNo, CSDetailKey


	'// 주문내역 중복으로 보냄, skyer9, 2018-01-26
	for i = 1 to 1
		Set oMaster = new COrderMasterItem

		'// 111111111111111
		oMaster.FSellSite 			= sellsite
		'oMaster.FOutMallOrderSerial = CStr(CLng(Request("ordNo")(i)))
		oMaster.FOutMallOrderSerial = CStr(Request("ordNo")(i))
		oMaster.FSellDate 			= Left(Now(), 10)
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= oMaster.FSellDate
		oMaster.FOrderUserID		= ""
		oMaster.FOrderName			= html2db(Trim(Request("rlOrdPrsnNm")(i)))
		oMaster.FOrderTelNo			= html2db(Trim(Request("rlOrdPrsnHomTel")(i)))
		oMaster.FOrderHpNo			= html2db(Trim(Request("rlOrdPrsnCelTel")(i)))
		oMaster.FOrderEmail			= ""
		oMaster.FReceiveName		= html2db(Trim(Request("custPrsnNm")(i)))
		oMaster.FReceiveTelNo		= html2db(Trim(Request("custPrsnHomTel")(i)))
		oMaster.FReceiveHpNo		= html2db(Trim(Request("custPrsnCelTel")(i)))

		oMaster.Fdeliverymemo		= html2db(Trim(Request("delivMsg")(i)))
		oMaster.FdeliverPay			= CLng(0)

		oMaster.FReceiveZipCode		= html2db(Trim(Request("delivZip")(i)))

		'// 우편번호 수정
		if Len(oMaster.FReceiveZipCode) = 4 then
			oMaster.FReceiveZipCode = "0" & oMaster.FReceiveZipCode
		end if

		oMaster.FReceiveAddr1		= html2db(Trim(Request("delivAddr1")(i)))
		oMaster.FReceiveAddr2		= html2db(Trim(Request("delivAddr2")(i)))

		if InStr(oMaster.FReceiveZipCode, "-") = 0 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
		oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)


		redim oDetailArr(0)
		Set oDetailArr(0) = new COrderDetail
		oDetailArr(0).FdetailSeq = Trim(Request("ordItemNo")(i))
		if Right(oDetailArr(0).FdetailSeq, 1) = "0" then
			oDetailArr(0).FdetailSeq = Left(oDetailArr(0).FdetailSeq, Len(oDetailArr(0).FdetailSeq) - 1)
			oDetailArr(0).FdetailSeq = CLng(oDetailArr(0).FdetailSeq)
		end if

		oDetailArr(0).FItemID = Trim(Request("supPrdCd")(i))
		oDetailArr(0).FItemOption = Trim(Request("dtlSupPrdCd")(i))
		oDetailArr(0).FOutMallItemOptionName = ""
		if Trim(Request("attrTypNm1")(i)) <> "" and Trim(Request("attrTypNm1")(i)) <> "None" then
			oDetailArr(0).FOutMallItemOptionName = Trim(Request("attrTypNm1")(i))
		end if
		if Trim(Request("attrTypNm2")(i)) <> "" and Trim(Request("attrTypNm2")(i)) <> "None" then
			oDetailArr(0).FOutMallItemOptionName = oDetailArr(0).FOutMallItemOptionName + "," + Trim(Request("attrTypNm2")(i))
		end if
		if Trim(Request("attrTypNm3")(i)) <> "" and Trim(Request("attrTypNm3")(i)) <> "None" then
			oDetailArr(0).FOutMallItemOptionName = oDetailArr(0).FOutMallItemOptionName + "," + Trim(Request("attrTypNm3")(i))
		end if
		if Trim(Request("attrTypNm4")(i)) <> "" and Trim(Request("attrTypNm4")(i)) <> "None" then
			oDetailArr(0).FOutMallItemOptionName = oDetailArr(0).FOutMallItemOptionName + "," + Trim(Request("attrTypNm4")(i))
		end if
		oDetailArr(0).FOutMallItemOptionName = html2db(oDetailArr(0).FOutMallItemOptionName)

		sqlStr = ""
		sqlStr = sqlStr & " SELECT TOP 1 itemid, itemoption"
		sqlStr = sqlStr & " FROM db_etcmall.[dbo].[tbl_Outmall_option_Manager] "
		sqlStr = sqlStr & " WHERE convert(varchar(20),itemid) + convert(varchar(20),itemoption) = '" & oDetailArr(0).FItemID & "' "
		sqlStr = sqlStr & " and mallid = 'gsshop' "
		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		If (Not rsget.EOF) Then
			oDetailArr(0).FItemID = rsget("itemid")
			oDetailArr(0).FItemOption = rsget("itemoption")
		End If
		rsget.Close

		if (oDetailArr(0).FItemOption="") then
			if (oDetailArr(0).FOutMallItemOptionName <> "") then
				oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, oDetailArr(0).FOutMallItemOptionName)
			else
				oDetailArr(0).FItemOption = "0000"
			end if
		end if

		if (oDetailArr(0).FItemOption = "0000") and (oDetailArr(0).FOutMallItemOptionName <> "") then
			oDetailArr(0).FOutMallItemOptionName = ""
		end if

		oDetailArr(0).FOutMallItemID = Trim(Request("prdCd")(i))
		for k = 0 to 100
			if Left(oDetailArr(0).FOutMallItemID,1) = "0" then
				oDetailArr(0).FOutMallItemID = Mid(oDetailArr(0).FOutMallItemID, 2, 1000)
			else
				exit for
			end if
		next
		oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption

		oDetailArr(0).FOutMallItemName = html2db(Trim(Request("prdNm")(i)))

		oDetailArr(0).FItemNo = Trim(Request("ordQty")(i))

		oDetailArr(0).Fitemcost = Trim(Request("stdUprc")(i))
		oDetailArr(0).FReducedPrice = Trim(Request("salePrc")(i))
		if (oDetailArr(0).FReducedPrice = "0") or (oDetailArr(0).FReducedPrice = "") then
			oDetailArr(0).FReducedPrice = oDetailArr(0).Fitemcost
		end if

		oDetailArr(0).FOutMallCouponPrice = 0
		oDetailArr(0).FTenCouponPrice = 0

		if oDetailArr(0).FItemOption = "" then
			oDetailArr(0).FItemOption = "0000"
			oDetailArr(0).FOutMallItemOption = "0000"
		end if

		if (oDetailArr(0).FItemOption <> "0000") and (application("Svr_Info") <> "Dev") then
			if Not GetCheckItemOptionValid(oDetailArr(0).FItemID, oDetailArr(0).FItemOption) then
				'// 잘못된 옵션.
				tmpOptionSeq = tmpOptionSeq + 1
				oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
				oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
			end if
		end if

		'// 주문제작문구 수정
		''if InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") <> 0 then
		''	oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FOutMallItemOptionName, InStr(oDetailArr(0).FOutMallItemOptionName, "텍스트를 입력하세요") + Len("텍스트를 입력하세요"), 1000)
		''	if left(oDetailArr(0).FrequireDetail, 1) = "：" then
		''		oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FrequireDetail, 2, 1000)
		''	end if
		''end if

		if (Trim(Request("ordTypeCd")(i)) = "O" and Trim(Request("ordTypeCdG")(i)) = "TA") then
			'// 일반 주문은 스킵 : 배송료가 없어서 엑셀로 주문입력한다. skyer9, 2019-02-07
			''if (SaveOrderToDB(oMaster, oDetailArr) = True) then
			''	''if GetOrderDetailConfirmFrom_auction1010(oDetailArr(0).FdetailSeq) then
			''		successCnt = successCnt + 1
			''	''end if
			''end if
		elseif (Trim(Request("ordTypeCd")(i)) = "C") then
			'// 1111111111111111111111
			'// 취소
	        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and CSDetailKey = '" & CStr("") & "' and OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('A008', '없음', '" & sellsite & "', '" & CStr(oMaster.FOutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(oDetailArr(0).FdetailSeq) & "', '" & CStr("") & "', " & oDetailArr(0).FItemNo & ") "
			strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"

			if (iAssignedRow > 0) then
				successCnt = successCnt + 1

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
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
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr("") & "' and c.OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		elseif (Trim(Request("ordTypeCd")(i)) = "R") then
			'// 반품

			'// orgOrdNo, orgOrdItemNo

			orgOrdNo = CStr(RemovePrecedingZero(Request("orgOrdNo")(i)))
			orgOrdItemNo = Trim(Request("orgOrdItemNo")(i))
			if Right(orgOrdItemNo, 1) = "0" then
				orgOrdItemNo = Left(orgOrdItemNo, Len(orgOrdItemNo) - 1)
				orgOrdItemNo = CLng(orgOrdItemNo)
			end if

			CSDetailKey = oMaster.FOutMallOrderSerial & "_" & oDetailArr(0).FdetailSeq
			oMaster.FOutMallOrderSerial = orgOrdNo
			oDetailArr(0).FdetailSeq = orgOrdItemNo

	        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('A004', '없음', '" & sellsite & "', '" & CStr(oMaster.FOutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(oDetailArr(0).FdetailSeq) & "', '" & CStr(CSDetailKey) & "', " & oDetailArr(0).FItemNo & ") "
			strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"

			if (iAssignedRow > 0) then
				successCnt = successCnt + 1

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
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
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		elseif (Trim(Request("ordTypeCd")(i)) = "X") then
			'// 교환

			'// orgOrdNo, orgOrdItemNo

			'orgOrdNo = CStr(CLng(Request("orgOrdNo")(i)))
			orgOrdNo = CStr(Request("orgOrdNo")(i))
			orgOrdItemNo = Trim(Request("orgOrdItemNo")(i))
			if Right(orgOrdItemNo, 1) = "0" then
				orgOrdItemNo = Left(orgOrdItemNo, Len(orgOrdItemNo) - 1)
				orgOrdItemNo = CLng(orgOrdItemNo)
			end if

			CSDetailKey = oMaster.FOutMallOrderSerial & "_" & oDetailArr(0).FdetailSeq
			oMaster.FOutMallOrderSerial = orgOrdNo
			oDetailArr(0).FdetailSeq = orgOrdItemNo

	        strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' ) "
			strSql = strSql & " BEGIN "
			strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
			strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
			strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
			strSql = strSql & " 	('A000', '없음', '" & sellsite & "', '" & CStr(oMaster.FOutMallOrderSerial) & "', '', '', '"& "" &"', '"& "" &"', '', "
			strSql = strSql & "		'', '', '', '', '', '' "
			strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(oDetailArr(0).FdetailSeq) & "', '" & CStr(CSDetailKey) & "', " & oDetailArr(0).FItemNo & ") "
			strSql = strSql & " END "
			dbget.Execute strSql,iAssignedRow
			''response.write strSql & "<br />"

			if (iAssignedRow > 0) then
				successCnt = successCnt + 1

				'' CS 마스터정보. 업데이트
				strSql = " update c "
				strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
				strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
				strSql = strSql + " , c.OrderName = o.OrderName "
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
				strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(oMaster.FOutMallOrderSerial) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(oDetailArr(0).FdetailSeq) & "' "
				''response.write strSql & "<br />"
				dbget.Execute strSql
			end if
		end if
		Set oMaster = Nothing
	next
end function

public function GetOrderFromJson_gseshop_Recv(objData)
	Dim sellsite : sellsite = "gseshop"
	Dim oMaster, oDetailArr(), sqlStr, isOptAddGS
	Dim tmpOptionSeq : tmpOptionSeq = 0
	Dim successCnt : successCnt = 0
	Dim i, j, k, pos
	Dim tmpStr, strSql, iAssignedRow
	Dim orgOrdNo, orgOrdItemNo, CSDetailKey, orderList
	Dim strObj

	Set strObj = JSON.parse(objData)
		'rw strObj.processType	'전송구분
		'rw strObj.sdDt			'조회일자
		'rw strObj.resultCnt	'처리결과수
		If strObj.resultCd = "S" Then	'처리결과코드 | (성공 : S, 실패 :E)
			Set orderList = strObj.resultList
				If orderList.length > 0 Then
					For i=0 to orderList.length - 1
						If orderList.get(i).ordTypeCd = "C" Then		'O:주문, R:반품(교환반품포함), X:교환주문, C:취소
							strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and CSDetailKey = '" & CStr("") & "' and OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' ) "
							strSql = strSql & " BEGIN "
							strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
							strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
							strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
							strSql = strSql & " 	('A008', '없음', '" & sellsite & "', '" & CStr(orderList.get(i).ordNo) & "', '', '', '"& "" &"', '"& "" &"', '', "
							strSql = strSql & "		'', '', '', '', '', '' "
							strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(orderList.get(i).ordItemNo) & "', '" & CStr("") & "', " & orderList.get(i).ordQty & ") "
							strSql = strSql & " END "
							dbget.Execute strSql, iAssignedRow
							''response.write strSql & "<br />"

							if (iAssignedRow > 0) then
								successCnt = successCnt + 1

								'' CS 마스터정보. 업데이트
								strSql = " update c "
								strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
								strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
								strSql = strSql + " , c.OrderName = o.OrderName "
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
								strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and c.CSDetailKey = '" & CStr("") & "' and c.OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' "
								''response.write strSql & "<br />"
								dbget.Execute strSql
							end if
						ElseIf orderList.get(i).ordTypeCd = "R" Then	'O:주문, R:반품(교환반품포함), X:교환주문, C:취소
							'// 반품
							'// orgOrdNo, orgOrdItemNo
							orgOrdNo = CStr(RemovePrecedingZero(orderList.get(i).orgOrdNo))
							orgOrdItemNo = Trim(orderList.get(i).orgOrdItemNo)

							CSDetailKey = orderList.get(i).ordNo & "_" & orderList.get(i).ordItemNo
							orderList.get(i).ordNo = orgOrdNo
							orderList.get(i).ordItemNo = orgOrdItemNo

							strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' ) "
							strSql = strSql & " BEGIN "
							strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
							strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
							strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
							strSql = strSql & " 	('A004', '없음', '" & sellsite & "', '" & CStr(orderList.get(i).ordNo) & "', '', '', '"& "" &"', '"& "" &"', '', "
							strSql = strSql & "		'', '', '', '', '', '' "
							strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(orderList.get(i).ordItemNo) & "', '" & CStr(CSDetailKey) & "', " & orderList.get(i).ordQty & ") "
							strSql = strSql & " END "
							dbget.Execute strSql,iAssignedRow
							''response.write strSql & "<br />"

							If (iAssignedRow > 0) Then
								successCnt = successCnt + 1

								'' CS 마스터정보. 업데이트
								strSql = " update c "
								strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
								strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
								strSql = strSql + " , c.OrderName = o.OrderName "
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
								strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' "
								''response.write strSql & "<br />"
								dbget.Execute strSql
							End If
						ElseIf orderList.get(i).ordTypeCd = "X" Then	'O:주문, R:반품(교환반품포함), X:교환주문, C:취소
							'// 교환
							'// orgOrdNo, orgOrdItemNo
							'orgOrdNo = CStr(CLng(Request("orgOrdNo")(i)))
							orgOrdNo = CStr(RemovePrecedingZero(orderList.get(i).orgOrdNo))
							orgOrdItemNo = Trim(orderList.get(i).orgOrdItemNo)

							CSDetailKey = orderList.get(i).ordNo & "_" & orderList.get(i).ordItemNo
							orderList.get(i).ordNo = orgOrdNo
							orderList.get(i).ordItemNo = orgOrdItemNo

							strSql = " IF NOT EXISTS (SELECT idx FROM db_temp.dbo.tbl_xSite_TMPCS WHERE SellSite = '" & sellsite & "' and OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and CSDetailKey = '" & CStr(CSDetailKey) & "' and OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' ) "
							strSql = strSql & " BEGIN "
							strSql = strSql & " 	INSERT INTO db_temp.dbo.tbl_xSite_TMPCS (divcd, gubunname, SellSite, OutMallOrderSerial, OrderName, OrderEmail, OrderTelNo, "
							strSql = strSql & " 	OrderHpNo, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, deliverymemo "
							strSql = strSql & "		, OutMallRegDate, OrgDetailKey, CSDetailKey, itemno) VALUES "
							strSql = strSql & " 	('A000', '없음', '" & sellsite & "', '" & CStr(orderList.get(i).ordNo) & "', '', '', '"& "" &"', '"& "" &"', '', "
							strSql = strSql & "		'', '', '', '', '', '' "
							strSql = strSql & "		, '" & Left(Now, 10) & "', '" & CStr(orderList.get(i).ordItemNo) & "', '" & CStr(CSDetailKey) & "', " & orderList.get(i).ordQty & ") "
							strSql = strSql & " END "
							dbget.Execute strSql,iAssignedRow
							''response.write strSql & "<br />"

							if (iAssignedRow > 0) then
								successCnt = successCnt + 1

								'' CS 마스터정보. 업데이트
								strSql = " update c "
								strSql = strSql + " set c.OrderSerial = o.OrderSerial, c.ItemID = o.matchItemID, c.itemoption = o.matchitemoption "
								strSql = strSql + " , c.OutMallItemName = o.orderItemName, c.OutMallItemOptionName = o.orderItemOptionName "
								strSql = strSql + " , c.OrderName = o.OrderName "
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
								strSql = strSql + " 	and c.SellSite = '" & sellsite & "' and c.OutMallOrderSerial = '" & CStr(orderList.get(i).ordNo) & "' and c.CSDetailKey = '" & CStr(CSDetailKey) & "' and c.OrgDetailKey = '" & CStr(orderList.get(i).ordItemNo) & "' "
								''response.write strSql & "<br />"
								dbget.Execute strSql
							end if
						End If
					Next
				End If
			Set orderList = nothing
		End If
	Set strObj = nothing
end function

Public Function getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iserv, ioper)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			iaccessLicense = "01000100004b035a25d67f991849cad1c7042b8da528d13e9ddce6878f2e43ac88080e0a5e" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		Else
			iaccessLicense = "010001000019133c715650b9c85b820961612f2b90b431ddd8654b42c097c4df1a43d0be09" 'AccessLicense Key 입력, PDF파일참조
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			iTimestamp = cryptoLib.getTimestamp()
			isignature = cryptoLib.generateSign(iTimestamp & iserv & ioper, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

Public Function generateKey_nvstorefarm(iTimestamp)
	Dim cryptoLib, oLicense, osecretKey, otimeStamp, osignature
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
		If (application("Svr_Info") = "Dev") Then
			osecretKey = "AQABAAAWPWagCrPjFQnFEtxs5j+oyZFwuzomdNq0XZSricPuMw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		Else
			osecretKey = "AQABAADX6Hz/wORFJS5pSIy4KQXkH83gC9G1aXChxBjcnUMqWw=="  'SecreKey 입력, PDF파일참조
			generateKey_nvstorefarm = cryptoLib.generateKey(iTimestamp, osecretKey)
		End If
	Set cryptoLib = nothing
End Function

Function PlaceProductOrder_lfmall(iOrderNo, iOrdDtlNo, isellsite)
	dim sellsite : sellsite = "lfmall"
	dim xmlURL, xmlSelldate
	dim objXML, strObj, objData, aKey, jParam, requireDetailObj, requireDetail
	dim masterCnt, optCode, optNm
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim tmpStr, pos, obj, orderList
	dim i, j, k, strsql
	dim found, deliveryDate
	dim successCnt : successCnt = 0
	Dim returnCode, dataVal
	Dim apiUrl, apiKey, istrParam
	dim tmpOptionSeq : tmpOptionSeq = 0

	istrParam = ""
	istrParam = istrParam & "<?xml version=""1.0"" encoding=""UTF-8""?>"
	istrParam = istrParam & "<ConfirmInfo>"
	istrParam = istrParam & "	<Header>"
	istrParam = istrParam & "		<AuthId><![CDATA[tenten]]></AuthId>"
	istrParam = istrParam & "		<AuthKey><![CDATA[Ten1010*!!]]></AuthKey>"
	istrParam = istrParam & "		<Format>XML</Format>"
	istrParam = istrParam & "		<Charset>UTF-8</Charset>"
	istrParam = istrParam & "	</Header>"
	istrParam = istrParam & "	<Body>"
	istrParam = istrParam & "		<Confirm>"
	istrParam = istrParam & "			<OrdNo>"& iOrderNo &"</OrdNo>"
	istrParam = istrParam & "			<OrdDtlNo>"& iOrdDtlNo &"</OrdDtlNo>"
	istrParam = istrParam & "		</Confirm>"
	istrParam = istrParam & "	</Body>"
	istrParam = istrParam & "</ConfirmInfo>"

    Dim  iRbody, iMessage
	Dim xmlDOM, retCode
	Dim REQUEST_XML
	REQUEST_XML = "REQUEST_XML=" & Server.URLEncode(istrParam)

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://b2b.lfmall.co.kr/interface.do?cmd=updaterConfirm", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(REQUEST_XML)
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "utf-8")
				xmlDOM.LoadXML iRbody

				retCode = xmlDOM.getElementsByTagName("ConfirmInfo/Body/Confirm/ResultCode").item(0).text
				If retCode = "SUCCESS" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
					strSql = strSql & " SET isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&iOrderNo&"'  "
					strSql = strSql & " and OrgDetailKey = '"&iOrdDtlNo&"' "
					strSql = strSql & " and mallid = 'lfmall' "
					dbget.Execute strSql
				End If
			Set xmlDOM = nothing
		End If
	Set objXML= nothing
End Function

'// 주문 발주처리
function PlaceProductOrder_nvstorefarm(ProductOrderID, isellsite)
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL
	dim strRst, objXML, xmlDOM

	PlaceProductOrder_nvstorefarm = False

	iServ		= "SellerService41"
	iCcd		= "PlaceProductOrder"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If isellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf isellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:PlaceProductOrderRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	strRst = strRst & "			<sel:ProductOrderID>"&ProductOrderID&"</sel:ProductOrderID>"
	strRst = strRst & "		</sel:PlaceProductOrderRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	PlaceProductOrder_nvstorefarm = True

	''set objMasterListXML = Nothing
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

function GetOrderDetailList_nvstorefarm(selldate, LastChangedStatusCode, isellsite)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL
	dim strRst, objXML, xmlDOM
	dim objMasterListXML, objMasterOneXML
	dim PrdOrderList(), i
	dim tmpXml, sellUtcDate

	Dim testStr1, testStr2
	testStr1 = request("testStr1")
	testStr2 = request("testStr2")

	redim PrdOrderList(-1)
	GetOrderDetailList_nvstorefarm = PrdOrderList

	iServ		= "SellerService41"
	iCcd		= "GetChangedProductOrderList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
If testStr1 <> "" Then
	strRst = strRst & "			<sel:InquiryTimeFrom>"& testStr1 &"</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"& testStr2 &"</sel:InquiryTimeTo>"										'조회 종료 일시(해당 시각 포함하지 않음)
Else
	sellUtcDate = Left(DateAdd("d", -1, CDate(selldate)), 10)
	strRst = strRst & "			<sel:InquiryTimeFrom>"&sellUtcDate&"T15:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
	strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(sellUtcDate)), 10)&"T15:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)

'	strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
'	strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
End If
	strRst = strRst & "			<sel:LastChangedStatusCode>" & LastChangedStatusCode & "</sel:LastChangedStatusCode>"				'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환 | PAYED : 결제완료)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'판매자 아이디
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"
If session("ssBctID")="kjy8517" Then
	rw objXML.responseText & "<br /><br />"
	rw "==================="
End If
	''dbget.close : response.end

	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) = 0 then
		''if IsAutoScript then
			response.write "내역없음<br />"
		''end if

		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")

	i = 0
	redim PrdOrderList(objMasterListXML.length - 1)
	For each objMasterOneXML in objMasterListXML
		PrdOrderList(i) = objMasterOneXML.getElementsByTagName("n:ProductOrderID")(0).Text
		i = i + 1
	next

	GetOrderDetailList_nvstorefarm = PrdOrderList

	set objMasterListXML = Nothing
	Set xmlDOM = Nothing
	Set objXML = Nothing
end function

function GetOrderFrom_nvstorefarm(isellsite, selldate)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If

	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim PrdOrderList, PrdOrder
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim cryptoLib
	dim keyGenerated
	Dim strSql, isDisCountYn, maySellPrice


	GetOrderFrom_nvstorefarm = False

	PrdOrderList = GetOrderDetailList_nvstorefarm(selldate, "PAYED", sellsite)

	response.write "건수(" & UBound(PrdOrderList) + 1 & ") " & "<br />"

	if UBound(PrdOrderList) < 0 then
		exit function
	end if

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
	For each PrdOrder in PrdOrderList
		strRst = strRst & "			<sel:ProductOrderIDList>" & PrdOrder & "</sel:ProductOrderIDList>" + vbCrLf
	next
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
If session("ssBctID")="kjy8517" Then
	response.write objXML.responseText & "<br /><br />"
'	response.end
End If
	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (UBound(PrdOrderList) + 1) then
		response.write "건수 불일치 오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	keyGenerated = generateKey_nvstorefarm(iTimestamp)
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
	set objMasterListXML = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")
	i = 0
	For each objMasterOneXML in objMasterListXML

		if objMasterOneXML.getElementsByTagName("n:CancelInfo").length > 0 then
			'// 취소주문
		elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length < 1) then
			'// 주소입력 안된 주문(선물하기 주문은 받는 사람이 주소를 입력한 이후에 끌어와야 한다.)
		else
			Set oMaster = new COrderMasterItem
			isDisCountYn = "N"
			maySellPrice = ""
			oMaster.FSellSite 			= sellsite
			oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("n:Order/n:OrderID")(0).Text

			If oMaster.FOutMallOrderSerial = "2020121995761581" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2020121995761581_1"
			ElseIf oMaster.FOutMallOrderSerial = "2021033148499521" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2021033148499521_1"
			ElseIf oMaster.FOutMallOrderSerial = "2022020275424071" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2022020275424071_1"
			ElseIf oMaster.FOutMallOrderSerial = "2022061268242041" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2022061268242041_1"
			End If

			oMaster.FSellDate 			= Left(Now(), 10)
			oMaster.FPayType			= "50"
			oMaster.FPaydate			= oMaster.FSellDate
			oMaster.FOrderUserID		= ""
			oMaster.FOrderName			= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererName")(0).Text)), 28)
			if (objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2").length > 0) then
				oMaster.FOrderTelNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2")(0).Text))
			else
				oMaster.FOrderTelNo = html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			end if
			oMaster.FOrderHpNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			oMaster.FOrderEmail			= ""
			''response.Write objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length
			''response.end
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name")(0).Text)), 28)
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name")(0).Text)), 28)
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if
			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2")(0).Text))
			elseif objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2")(0).Text))
			else
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if
			end if

			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo").length > 0 then
				oMaster.Fdeliverymemo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo")(0).Text), 180)
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount").length > 0 then
				oMaster.FdeliverPay = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount")(0).Text
			end if

			If sellsite <> "nvstorefarmclass" Then
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode")(0).Text)
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode")(0).Text)
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress")(0).Text))
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress").length > 0) then
					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress")(0).Text))
				else
					oMaster.FReceiveAddr2		= "" '아래 주석 부분으로 했더니 출고지 주소가 출력 됨 (내용 -> 도봉동 여인닷컴)
'				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress").length > 0) then
'					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress")(0).Text))
				end if
				if InStr(oMaster.FReceiveZipCode, "-") = 0 then
					oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
				end if

				'// 주소 수정
				oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
				oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
				tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
				pos = 0
				for k = 0 to 2
					pos = InStr(pos+1, tmpStr, " ")
					if (pos = 0) then
						exit for
					end if
				next

				if (pos > 0) then
					oMaster.FReceiveAddr1 = Left(tmpStr, pos)
					oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
				end if

				oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
				oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)
			End If

			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOrderID")(0).Text
			oDetailArr(0).FItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode").length > 0) then
				oDetailArr(0).FItemOption = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode")(0).Text
			else
				oDetailArr(0).FItemOption = "0000"
			end if

			oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductID")(0).Text
			oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
			oDetailArr(0).FOutMallItemName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductName")(0).Text)
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption").length > 0) then
				oDetailArr(0).FOutMallItemOptionName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption")(0).Text)
			else
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:Quantity")(0).Text)

			'2019-08-06 김진영 아래 조건 추가
			'스토어팜 매입이면서 할인기간이라면 판매가(itemcost)를 실판매가(reducedprice)와 동일하게 저장
			'If left(now(),10) >= "2019-10-2" and left(now(),10) < "2019-09-24" Then
			'2019-10-21 김진영, 위 now()에서 Date로 변경 / Case SellerProductCode CSTR문자 변환, Trim 처리
			'2020-09-10 김진영, 스토어팜 특가관리에 추가했다면 할인가격으로 변경되게 수정
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
			strSql = strSql & " WHERE mallgubun = '"& sellsite &"' "
			strSql = strSql & " and itemid = '"& CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text)) &"' "
			strSql = strSql & " and GETDATE() >= startDate and GETDATE() <= endDate "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") > 0 Then
					isDisCountYn = "Y"
				Else
					isDisCountYn = "N"
				End If
			rsget.Close

			If isDisCountYn = "Y" Then
'				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 시작
				maySellPrice = Clng(objMasterOneXML.getElementsByTagName("n:UnitPrice")(0).Text) + Clng(objMasterOneXML.getElementsByTagName("n:OptionPrice")(0).Text)
				If (objMasterOneXML.getElementsByTagName("n:ProductImediateDiscountAmount").length > 0) then
					oDetailArr(0).Fitemcost = maySellPrice - Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductImediateDiscountAmount")(0).Text) / oDetailArr(0).FItemNo)
				Else
					oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
				End If
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 끝
			Else
				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			End If

			' If Date() = "2019-11-04" Then
			' 	SELECT CASE CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text))
			' 		Case "2420183","2420289","2420141","1849622","2462605","2209031","2432006","2498125","2498123","2275592","2191086","2428957","2100299","2215001","2157875","2214948","2330710","2215004","2275586","2275593","2215003","2275682","2100304","2241174","2209032","2000213","2080162","2423170","2185356","2515560","2420169","2157871","2150273","1921244","2423290","2374165","2078920","2191087","2422589","1683251","1921247","2431879","2333117","2207738","2493633","2431875"
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 		Case Else
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 	End Select
			' ElseIf Date() = "2020-08-10" Then
			' 	SELECT CASE CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text))
			' 		Case "2784155", "2849183", "2843167", "2445376", "2420294", "2733604", "1936993", "2351405", "3078814", "3078813", "3078811", "3078809", "3008917", "2791972", "2365294", "2844462", "2438078", "1908698", "2725225", "2328248", "2857532", "1802492", "2365035", "2832840", "2213896", "2501196", "2225256", "1646072", "2252408", "2357727", "2784157", "2931152", "2780193", "2523634", "2662135", "2777555", "2896860", "2662083", "2770775", "2662109", "2833873", "2770780", "2543514", "2778552", "2523612", "2796450", "2819587", "2706807", "2856290", "2785591", "2574519", "2701283", "2774770", "2788824", "2420193", "2777618", "2432006", "2708339", "2360463", "2819538", "2819537", "2733614", "2816524", "2852519", "2850549", "2689404", "2862196", "2920248", "1906618", "2215001", "2877882", "2731812", "2420186", "2878137", "2551434", "2445387", "2445386", "2814192", "2780188", "2445366", "2445365",   "2783818","2445385","2793715","2878524","2601184","2601239","2929782","2926754","2391691","2383744","2383739","2374180","2924063","2394946","2786467","2786466","2601007","2584266","2527155","2527149","2527148","2423290","2330680","2241269","2241252","2241214","2241174"
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 		Case Else
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 	End Select
			' ElseIf (Date() >= "2020-09-14" AND Date() <= "2020-09-20") Then
			' 	SELECT CASE CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text))
			' 		Case "3019191", "3019218", "2927565", "2927567", "2927753", "2927556", "2927727", "2780188", "2777618", "2788823"
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 		Case Else
			' 			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			' 	End Select
			' Else
			' 	oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			' End If
'			oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)

			oDetailArr(0).FReducedPrice = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0

			'// 주문제작문구 수정
			''if InStr(oDetailArr(0).FOutMallItemOptionName, "직접입력") <> 0 then
			''	oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FOutMallItemOptionName, InStr(oDetailArr(0).FOutMallItemOptionName, "직접입력") + Len("직접입력"), 1000)
			''end if

			if (SaveOrderToDB(oMaster, oDetailArr) = True) or (oMaster.FOutMallOrderSerial = "2017111282590911") or (oMaster.FOutMallOrderSerial = "2017120682960501") then
				if PlaceProductOrder_nvstorefarm(oDetailArr(0).FdetailSeq, sellsite) then
					successCnt = successCnt + 1
				end if
			end if
			i = i + 1
		end if
	next
	Set cryptoLib = Nothing

	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	GetOrderFrom_nvstorefarm = True
	Set xmlDOM = Nothing
	Set objXML = Nothing

end function

function ChangeOrderStatus_ezwel(orderNum, orderGoodsNum)
	dim xmlURL, postParam
	dim strRst
	dim objXML, xmlDOM

	ChangeOrderStatus_ezwel = False

	xmlURL = "http://api.ezwel.com/if/api/orderStatusInfoAPI.ez"
	postParam = "cspCd=" & cspCd & "&crtCd=" & crtCd & "&dataSet="

	strRst = ""
	strRst = strRst & "<?xml version=""1.0"" encoding=""euc-kr""?>"
	strRst = strRst & "<dataSet>"
	strRst = strRst & "       <arrOrderStatusInfo>"
	strRst = strRst & "              <orderNum>" & orderNum & "</orderNum>"
	strRst = strRst & "              <orderGoodsNum>" & orderGoodsNum & "</orderGoodsNum>"
	strRst = strRst & "              <orderStatus>1002</orderStatus>"
	strRst = strRst & "              <orderMemo></orderMemo>"
	strRst = strRst & "       </arrOrderStatusInfo>"
	strRst = strRst & "</dataSet>"
	''response.write strRst
	''dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL ''& "?" & postParam
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send(postParam & strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if

	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	If xmlDOM.getElementsByTagName("resultSet/resultCode").item(0).text <> "200" Then
		response.write "주문상태 전송오류 : " & xmlDOM.getElementsByTagName("resultSet/resultMsg").item(0).text & "<br />"
		exit function
	end if

	ChangeOrderStatus_ezwel = True

end function

CONST cspCd		= "10040413"							'CP업체코드(이지웰 발급)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'보안코드(이지웰 발급)
function GetOrderFrom_ezwel(selldate)
	dim sellsite : sellsite = "ezwel"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim postParam
	dim tmpXML, oSql

	GetOrderFrom_ezwel = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(기간동안의 주문 가져오기)
	xmlURL = "http://api.ezwel.com/if/api/orderListAPI.ez"
	''response.write xmlURL

	postParam = "cspCd=" & cspCd & "&crtCd=" & crtCd
	postParam = postParam & "&startDate=" & Replace(selldate, "-", "") & "000000"
	postParam = postParam & "&endDate=" & Replace(Left(DateAdd("d", 1, CDate(selldate)), 10), "-", "") & "000000"
	postParam = postParam & "&orderStatus=1001"
	''response.write postParam

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send(postParam)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
	''response.write objXML.responseText & "<br /><br />"

	Set obj = Nothing

	if (xmlDOM.getElementsByTagName("resultSet/arrOrderList").length < 1) then
		''if IsAutoScript then
			response.write "내역없음 : 종료" & "<br />"
		''end if

		GetOrderFrom_ezwel = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	else
		response.write "건수(" & xmlDOM.getElementsByTagName("resultSet/arrOrderList").length & ") " & "<br />"
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("resultSet/arrOrderList")
	For each objMasterOneXML in objMasterListXML
		Set oMaster = new COrderMasterItem

		oMaster.FSellSite 			= sellsite
		oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("orderNum")(0).Text
		oMaster.FSellDate 			= Left(Now(), 10)
		oMaster.FPayType			= "50"
		oMaster.FPaydate			= oMaster.FSellDate
		oMaster.FOrderUserID		= ""

		oMaster.FOrderName			= LEFT(html2db(objMasterOneXML.getElementsByTagName("sndNm")(0).Text), 28)
		oMaster.FOrderTelNo			= LEFT(html2db(objMasterOneXML.getElementsByTagName("sndTelNum")(0).Text), 16)
		oMaster.FOrderHpNo			= LEFT(html2db(objMasterOneXML.getElementsByTagName("sndMobile")(0).Text), 16)
		if Len(CStr(oMaster.FOrderTelNo)) <= 3 then
			oMaster.FOrderTelNo = oMaster.FOrderHpNo
		end if

		oMaster.FOrderEmail			= ""
		oMaster.FReceiveName		= LEFT(html2db(objMasterOneXML.getElementsByTagName("rcvrNm")(0).Text), 28)
		oMaster.FReceiveTelNo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("rcvrTelNum")(0).Text), 16)
		oMaster.FReceiveHpNo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("rcvrMobile")(0).Text), 16)
		if Len(CStr(oMaster.FReceiveTelNo)) <= 3 then
			oMaster.FReceiveTelNo = oMaster.FReceiveHpNo
		end if

		oMaster.Fdeliverymemo		= html2db(objMasterOneXML.getElementsByTagName("orderReqContent")(0).Text)
		oMaster.FdeliverPay 		= objMasterOneXML.getElementsByTagName("arrOrderGoods/dlvrPrice")(0).Text

		oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("rcvrPost")(0).Text)
		oMaster.FReceiveAddr1		= html2db(objMasterOneXML.getElementsByTagName("rcvrAddr1")(0).Text)
		oMaster.FReceiveAddr2		= html2db(objMasterOneXML.getElementsByTagName("rcvrAddr2")(0).Text)

		if InStr(oMaster.FReceiveZipCode, "-") = 0 then
			oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
		end if

		'// 주소 수정
		oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
		oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
		tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
		pos = 0
		for k = 0 to 2
			pos = InStr(pos+1, tmpStr, " ")
			if (pos = 0) then
				exit for
			end if
		next

		if (pos > 0) then
			oMaster.FReceiveAddr1 = Left(tmpStr, pos)
			oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
		end if

		oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
		oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)


		set objDetailListXML = objMasterOneXML.getElementsByTagName("arrOrderGoods")
		For each objDetailOneXML in objDetailListXML
			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail

			oDetailArr(0).FdetailSeq = objDetailOneXML.getElementsByTagName("orderGoodsNum")(0).Text

			oDetailArr(0).FItemID = objDetailOneXML.getElementsByTagName("cspGoodsCd")(0).Text
			oDetailArr(0).FItemOption = ""																	'// 옵션코드 안줌
			oDetailArr(0).FOutMallItemID = objDetailOneXML.getElementsByTagName("goodsCd")(0).Text
			oDetailArr(0).FOutMallItemOption = ""
			oDetailArr(0).FOutMallItemName = html2db(objDetailOneXML.getElementsByTagName("goodsNm")(0).Text)

			oDetailArr(0).FOutMallItemOptionName = html2db(Trim(objDetailOneXML.getElementsByTagName("optionContent")(0).Text))
			if Right(oDetailArr(0).FOutMallItemOptionName,1) = "^" then
				oDetailArr(0).FOutMallItemOptionName = Left(oDetailArr(0).FOutMallItemOptionName, Len(oDetailArr(0).FOutMallItemOptionName) - 1)
			end if
			if Left(oDetailArr(0).FOutMallItemOptionName,3) = "선택:" then
				oDetailArr(0).FOutMallItemOptionName = Mid(oDetailArr(0).FOutMallItemOptionName, 4, 1000)
			end if

			oDetailArr(0).FItemNo = objDetailOneXML.getElementsByTagName("orderQty")(0).Text
			oDetailArr(0).Fitemcost = CLng(objDetailOneXML.getElementsByTagName("salePrice")(0).Text)
			oDetailArr(0).FReducedPrice = CLng(objDetailOneXML.getElementsByTagName("salePrice")(0).Text)

			if objDetailOneXML.getElementsByTagName("dccpnPrice")(0).getAttribute("class") = "string" then
				oDetailArr(0).Fitemcost = oDetailArr(0).Fitemcost + CLng(objDetailOneXML.getElementsByTagName("dccpnPrice")(0).Text)
			end if
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0

			'// 주문제작문구 수정
			''if InStr(oDetailArr(0).FOutMallItemOptionName, "직접입력") <> 0 then
			''	oDetailArr(0).FrequireDetail = Mid(oDetailArr(0).FOutMallItemOptionName, InStr(oDetailArr(0).FOutMallItemOptionName, "직접입력") + Len("직접입력"), 1000)
			''end if

			''''옶션명 CASE ":^" //2019/02/11
			if (oDetailArr(0).FOutMallItemOptionName = ":") then oDetailArr(0).FOutMallItemOptionName=""

			if objDetailOneXML.getElementsByTagName("cancelDt")(0).getAttribute("class") = "string" then
				'// 취소주문
			else
				if (oDetailArr(0).FOutMallItemOptionName = "") then
					oDetailArr(0).FItemOption = "0000"
					oDetailArr(0).FOutMallItemOption = "0000"
				else
					oDetailArr(0).FItemOption = GetItemOptionWithOptionName(sellsite, oDetailArr(0).FItemID, oDetailArr(0).FOutMallItemOptionName)

					if (oDetailArr(0).FItemOption = "0000") then
						tmpOptionSeq = tmpOptionSeq + 1
						oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
						oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
					end if
				end if

'				response.Write "oMaster.FOutMallOrderSerial : " & oMaster.FOutMallOrderSerial & "<br />"
'				response.Write "oDetailArr(0).FItemID : " & oDetailArr(0).FItemID & "<br />"
'				response.Write "oDetailArr(0).FItemOption : " & oDetailArr(0).FItemOption & "<br />"
'				response.Write "oDetailArr(0).FOutMallItemID : " & oDetailArr(0).FOutMallItemID & "<br />"
'				response.Write "oDetailArr(0).FOutMallItemOption : " & oDetailArr(0).FOutMallItemOption & "<br />"

				'2019-04-25 김진영 하단 추가..종료한 상품인데 판매가 됨..
				If oDetailArr(0).FOutMallItemID = "1016606096" Then
					oDetailArr(0).FItemID = "1891124"
				End If

				'#################### 2019-04-29 작업 이전 ####################
' 				if (oDetailArr(0).FItemID = "") then
' '					response.Write "상품코드 없음 : <br />주문번호 : " & oMaster.FOutMallOrderSerial "<br />" & "제휴몰 상품코드 : " & oDetailArr(0).FOutMallItemID
' 					response.Write "상품코드 없음 : <br />제휴몰 상품코드 : " & oDetailArr(0).FOutMallItemID
' 					dbget.close() : response.end
' 				end if

' 				if (SaveOrderToDB(oMaster, oDetailArr) = True) then
' 					'// 배송비는 한번만 입력
' 					oMaster.FdeliverPay = 0
' 					if ChangeOrderStatus_ezwel(oMaster.FOutMallOrderSerial, oDetailArr(0).FdetailSeq) then
' 						successCnt = successCnt + 1
' 					end if
' 				end if
				'#################### 2019-04-29 작업 이후 ####################
				If (oDetailArr(0).FItemID = "") Then
					oSql = ""
					oSql = oSql & " IF NOT EXISTS (SELECT TOP 1 idx FROM db_etcmall.dbo.tbl_outmall_sms_log WHERE outmallGoodNo = '"& oDetailArr(0).FOutMallItemID &"') " & vbcrlf
					oSql = oSql & " 	BEGIN " & vbcrlf
					oSql = oSql & " 		INSERT INTO db_etcmall.dbo.tbl_outmall_sms_log (mallid, outmallGoodNo, regdate) " & vbcrlf
					oSql = oSql & " 		VALUES ('ezwel', '"& oDetailArr(0).FOutMallItemID &"', GETDATE()) " & vbcrlf

					' 발송 수기로 박지 말것.
					'oSql = oSql & " 		INSERT INTO smsdb.db_infosms.dbo.em_smt_tran (date_client_req, content, callback, service_type, broadcast_yn, msg_status, recipient_num) " & vbcrlf
					'oSql = oSql & " 		VALUES (GETDATE(), '"& oDetailArr(0).FOutMallItemID &" 상품코드 누락', '1644-6030', '0', 'N', '1', '010-9972-8517') " & vbcrlf
					oSql = oSql & " 	END  "
					dbget.Execute(oSql)

 					response.Write "상품코드 없음 : <br />제휴몰 상품코드 : " & oDetailArr(0).FOutMallItemID

					call SendNormalSMS_LINK("010-9972-8517", "1644-6030", oDetailArr(0).FOutMallItemID & " 상품코드 누락")

 					dbget.close() : response.end
				Else
					if (SaveOrderToDB(oMaster, oDetailArr) = True) then
						'// 배송비는 한번만 입력
						oMaster.FdeliverPay = 0
						if ChangeOrderStatus_ezwel(oMaster.FOutMallOrderSerial, oDetailArr(0).FdetailSeq) then
							successCnt = successCnt + 1
						end if
					end if
				end if
				'#################### 2019-04-29 작업 이후 끝 ####################
			end if

			''response.write "oDetailArr(0).Fitemcost : " & oDetailArr(0).Fitemcost & "<br />"
			''response.write "oDetailArr(0).FReducedPrice : " & oDetailArr(0).FReducedPrice & "<br />"
		next

		''response.write "FOutMallOrderSerial : " & oMaster.FOutMallOrderSerial & "<br />"
		''response.write "FReceiveAddr2 : " & oMaster.FReceiveAddr2 & "<br />"
	next


	''if IsAutoScript then
		response.write "주문입력(" & successCnt & ")" & "<br />"
	''end if

	GetOrderFrom_ezwel = True
	Set xmlDOM = Nothing
	Set objXML = Nothing

end function

Function GetOrder_nvstorefarm(isellsite, currdate, hasMoreData, chgCode, lastOrderNo, lastTime, ixmlCheck)
	Dim sellsite, sellUtcDate
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim xmlURL, strSql
	dim strRst, objXML, xmlDOM
	dim objMasterListXML, objMasterOneXML, i

	iServ		= "SellerService41"
	iCcd		= "GetChangedProductOrderList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">"
	strRst = strRst & "	<soapenv:Header/>"
	strRst = strRst & "	<soapenv:Body>"
	strRst = strRst & "		<sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "			<sel:AccessCredentials>"
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>"
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>"
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>"
	strRst = strRst & "			</sel:AccessCredentials>"
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>"
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>"															'#돌려받는 데이터의 상세 정도(Compact / Full)
	strRst = strRst & "			<sel:Version>4.1</sel:Version>"
	If hasMoreData = "Y" Then
		sellUtcDate = Left(DateAdd("d", -1, CDate(currdate)), 10)
		strRst = strRst & "			<sel:InquiryTimeFrom>"&lastTime&"</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(sellUtcDate)), 10)&"T15:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
'		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
		strRst = strRst & "			<sel:InquiryExtraData>"&lastOrderNo&"</sel:InquiryExtraData>"
	Else
		sellUtcDate = Left(DateAdd("d", -1, CDate(currdate)), 10)
		strRst = strRst & "			<sel:InquiryTimeFrom>"&sellUtcDate&"T15:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(sellUtcDate)), 10)&"T15:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)

'		strRst = strRst & "			<sel:InquiryTimeFrom>"&selldate&"T00:00:00</sel:InquiryTimeFrom>"									'#조회 시작 일시(해당 시각 포함)
'		strRst = strRst & "			<sel:InquiryTimeTo>"& Left(DateAdd("d", 1, CDate(selldate)), 10)&"T00:00:00</sel:InquiryTimeTo>"	'조회 종료 일시(해당 시각 포함하지 않음)
	End If
	strRst = strRst & "			<sel:LastChangedStatusCode>" & chgCode & "</sel:LastChangedStatusCode>"								'최종 상품 주문 상태 코드 (CANCELED | 취소, RETURNED | 반품, EXCHANGED : 교환 | PAYED : 결제완료)
	strRst = strRst & "			<sel:MallID>"&reqID&"</sel:MallID>"																	'판매자 아이디
	strRst = strRst & "		</sel:GetChangedProductOrderListRequest>"
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	 If session("ssBctID")="kjy8517" and ixmlCheck = "Y" Then
	 	response.write strRst
	 	rw "===================1"
	 End If
	''dbget.close : response.end

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "POST", xmlURL
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
		objXML.send(strRst)
 		If objXML.Status = 200 Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML(objXML.responseText)
				'  If session("ssBctID")="kjy8517" Then
				'   	rw objXML.responseText & "<br /><br />"
				'   	rw "===================2"
				'  End If
				ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
				If ResponseType = "SUCCESS" Then
					If xmlDOM.getElementsByTagName("n:HasMoreData").item(0).text = "true" Then
						hasMoreData = "Y"
						lastOrderNo	= xmlDOM.getElementsByTagName("n:InquiryExtraData").item(0).text
						lastTime	= xmlDOM.getElementsByTagName("n:MoreDataTimeFrom").item(0).text
					Else
						hasMoreData = "N"
					End If

					If CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) > 0 Then
						Set objMasterListXML = xmlDOM.getElementsByTagName("n:ChangedProductOrderInfoList")
							For Each objMasterOneXML in objMasterListXML
								strSql = ""
								strSql = strSql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] ([sellsite], [OutMallOrderSerial], [regdate]) "
								strSql = strSql & " VALUES ('"& sellsite &"', '"& objMasterOneXML.getElementsByTagName("n:ProductOrderID")(0).Text &"', '"& currdate &"') "
								dbget.Execute strSql
								'rw strSql
							Next
						Set objMasterListXML = nothing
					Else
						''if IsAutoScript then
							response.write "내역없음<br />"
						''end if

						Set xmlDOM = Nothing
						Set objXML = Nothing
						exit function
					End If
				Else
					response.write "오류 : 종료"
					Set xmlDOM = Nothing
					Set objXML = Nothing
					dbget.close : response.end
				End If
			Set xmlDOM = Nothing
		Else
			If IsAutoScript then
				response.write "ERROR : 통신오류"
			Else
				response.write "ERROR : 통신오류" & objXML.Status
				response.write "<script>alert('ERROR : 통신오류.');</script>"
			End If
			dbget.close : response.end
		End If
	Set objXML = Nothing
End Function

Function GetOrderFrom_NewCall_nvstorefarm(isellsite, currdate, ilp, ixmlCheck)
	dim sellsite
	If isellsite = "nvstorefarm" Then
		sellsite = "nvstorefarm"
	ElseIf isellsite = "nvstoregift" Then
		sellsite = "nvstoregift"
	ElseIf isellsite = "Mylittlewhoopee" Then
		sellsite = "Mylittlewhoopee"
	Else
		sellsite = "nvstoremoonbangu"
	End If

	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, objMasterOneXML
	dim objDetailListXML, objDetailOneXML
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst, arrRows
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim iaccessLicense, iTimestamp, isignature, iServ, iCcd, reqID, ResponseType
	dim cryptoLib
	dim keyGenerated
	Dim strSql, isDisCountYn, maySellPrice, mayCnt
	Dim storeOrderDate, storeIpkumDate
	GetOrderFrom_NewCall_nvstorefarm = False

	strSql = ""
	strSql = strSql & " SELECT COUNT(*) as cnt "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] "
	strSql = strSql & " WHERE sellsite = '"& sellsite &"' "
	strSql = strSql & " and regdate = '"& currdate &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		mayCnt = rsget("cnt")
	rsget.Close
	response.write "Order Count(" & mayCnt & ")<br />"

	If mayCnt = 0 Then
		exit function
	End If

	strSql = ""
	strSql = strSql & " SELECT outmallOrderSerial "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_storefarm] "
	strSql = strSql & " WHERE sellsite = '"& sellsite &"' "
	strSql = strSql & " and regdate = '"& currdate &"' "
	strSql = strSql & " and num = '"& ilp &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	iServ		= "SellerService41"
	iCcd		= "GetProductOrderInfoList"

	Call getsecretKey_nvstorefarm(iaccessLicense, iTimestamp, isignature, iServ, iCcd)

	'// =======================================================================
	'// API URL(기간동안의 주문 가져오기)
	If (application("Svr_Info") = "Dev") Then
		xmlURL = "http://sandbox.api.naver.com/ShopN/"&iServ
	Else
		xmlURL = "http://ec.api.naver.com/ShopN/"&iServ
	End If
	''response.write xmlURL

	If (application("Svr_Info") = "Dev") Then
		reqID = "qa2tc329"
	Else
		If sellsite = "nvstorefarm" Then
			reqID = "tenten"
		ElseIf sellsite = "nvstoregift" Then
			reqID = "ncp_1o1934_01"
		Else
			reqID = "ncp_1np6kl_01"
		End If
	End If

	strRst = ""
	strRst = strRst & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:sel=""http://seller.shopn.platform.nhncorp.com/"">" + vbCrLf
	strRst = strRst & "	<soapenv:Header/>" + vbCrLf
	strRst = strRst & "	<soapenv:Body>" + vbCrLf
	strRst = strRst & "		<sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "			<sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "				<sel:AccessLicense>"&iaccessLicense&"</sel:AccessLicense>" + vbCrLf
	strRst = strRst & "				<sel:Timestamp>"&iTimestamp&"</sel:Timestamp>" + vbCrLf
	strRst = strRst & "				<sel:Signature>"&isignature&"</sel:Signature>" + vbCrLf
	strRst = strRst & "			</sel:AccessCredentials>" + vbCrLf
	strRst = strRst & "			<sel:RequestID>"&reqID&"</sel:RequestID>" + vbCrLf
	strRst = strRst & "			<sel:DetailLevel>Full</sel:DetailLevel>" + vbCrLf
	strRst = strRst & "			<sel:Version>4.1</sel:Version>" + vbCrLf
	For i = 0 To ubound(arrRows,2)
		strRst = strRst & "			<sel:ProductOrderIDList>" & arrRows(0,i) & "</sel:ProductOrderIDList>" + vbCrLf
	next
	strRst = strRst & "		</sel:GetProductOrderInfoListRequest>" + vbCrLf
	strRst = strRst & "	</soapenv:Body>"
	strRst = strRst & "</soapenv:Envelope>"
	'response.write strRst
	' dbget.close : response.end

	'// =======================================================================
	'// 데이타 가져오기
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	objXML.setRequestHeader "SOAPAction", iServ & "#" & iccd
	objXML.send(strRst)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : 통신오류"
		else
			response.write "ERROR : 통신오류" & objXML.Status
			response.write "<script>alert('ERROR : 통신오류.');</script>"
		end if

		dbget.close : response.end
	end if


	'// =======================================================================
	'// XML DOM 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML(objXML.responseText)
If session("ssBctID")="kjy8517" and ixmlCheck = "Y" Then
	response.write objXML.responseText & "<br /><br />"
'	response.end
End If
	ResponseType = xmlDOM.getElementsByTagName("n:ResponseType").item(0).text
	If ResponseType <> "SUCCESS" Then
		response.write "오류 : 종료"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		dbget.close : response.end
	end if

	' if CLng(xmlDOM.getElementsByTagName("n:ReturnedDataCount").item(0).text) <> (mayCnt) then
	' 	response.write "건수 불일치 오류 : 종료"
	' 	Set xmlDOM = Nothing
	' 	Set objXML = Nothing
	' 	dbget.close : response.end
	' end if

	keyGenerated = generateKey_nvstorefarm(iTimestamp)
	Set cryptoLib = Server.CreateObject("NHNAPIPlatform.SimpleCryptoLib")
	set objMasterListXML = xmlDOM.getElementsByTagName("n:ProductOrderInfoList")
	i = 0
	For each objMasterOneXML in objMasterListXML

		if objMasterOneXML.getElementsByTagName("n:CancelInfo").length > 0 then
			'// 취소주문
		elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length < 1) then
			'// 주소입력 안된 주문(선물하기 주문은 받는 사람이 주소를 입력한 이후에 끌어와야 한다.)
		else
			Set oMaster = new COrderMasterItem
			isDisCountYn = "N"
			maySellPrice = ""
			storeOrderDate = ""
			storeIpkumDate = ""

			oMaster.FSellSite 			= sellsite
			oMaster.FOutMallOrderSerial = objMasterOneXML.getElementsByTagName("n:Order/n:OrderID")(0).Text
			storeOrderDate = Replace(LEFT(objMasterOneXML.getElementsByTagName("n:Order/n:OrderDate")(0).Text, 19), "T", " ")		'주문 일시
			storeIpkumDate = Replace(LEFT(objMasterOneXML.getElementsByTagName("n:Order/n:PaymentDate")(0).Text, 19), "T", " ")		'결제 일시(최종 결제)

			storeOrderDate = dateconvert(dateadd("h", 9, storeOrderDate))
			storeIpkumDate = dateconvert(dateadd("h", 9, storeIpkumDate))

			If oMaster.FOutMallOrderSerial = "2020121995761581" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2020121995761581_1"
			ElseIf oMaster.FOutMallOrderSerial = "2021033148499521" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2021033148499521_1"
			ElseIf oMaster.FOutMallOrderSerial = "2022020275424071" AND sellsite = "nvstoregift" Then
				oMaster.FOutMallOrderSerial = "2022020275424071_1"
			End If
			'oMaster.FSellDate 			= Left(Now(), 10)
			oMaster.FSellDate 			= storeOrderDate		'2021-12-09 김진영 수정

			oMaster.FPayType			= "50"
'			oMaster.FPaydate			= oMaster.FSellDate
			oMaster.FPaydate			= storeIpkumDate		'2021-12-09 김진영 수정
			oMaster.FOrderUserID		= ""
			oMaster.FOrderName			= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererName")(0).Text)), 28)
			if (objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2").length > 0) then
				oMaster.FOrderTelNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel2")(0).Text))
			else
				oMaster.FOrderTelNo = html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			end if
			oMaster.FOrderHpNo			= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:Order/n:OrdererTel1")(0).Text))
			oMaster.FOrderEmail			= ""
			''response.Write objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length
			''response.end
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Name")(0).Text)), 28)
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name").length > 0) then
				oMaster.FReceiveName		= LEFT(html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Name")(0).Text)), 28)
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if
			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel2")(0).Text))
			elseif objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2").length > 0 then
				oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel2")(0).Text))
			else
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
					oMaster.FReceiveTelNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if
			end if

			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:Tel1")(0).Text))
			elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1").length > 0) then
				oMaster.FReceiveHpNo		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:Tel1")(0).Text))
			else
				response.Write "ERROR : 시스템팀 문의"
				response.end
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo").length > 0 then
				oMaster.Fdeliverymemo		= LEFT(html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingMemo")(0).Text), 180)
			end if

			if objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount").length > 0 then
				oMaster.FdeliverPay = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:DeliveryFeeAmount")(0).Text
			end if

			If sellsite <> "nvstorefarmclass" Then
				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:ZipCode")(0).Text)
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode").length > 0) then
					oMaster.FReceiveZipCode		= html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:ZipCode")(0).Text)
				else
					response.Write "ERROR : 시스템팀 문의"
					response.end
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:BaseAddress")(0).Text))
				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress").length > 0) then
					oMaster.FReceiveAddr1		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:BaseAddress")(0).Text))
				end if

				if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress").length > 0) then
					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ShippingAddress/n:DetailedAddress")(0).Text))
				else
					oMaster.FReceiveAddr2		= "" '아래 주석 부분으로 했더니 출고지 주소가 출력 됨 (내용 -> 도봉동 여인닷컴)
'				elseif (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress").length > 0) then
'					oMaster.FReceiveAddr2		= html2db(cryptoLib.decrypt(keyGenerated, objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TakingAddress/n:DetailedAddress")(0).Text))
				end if
				if InStr(oMaster.FReceiveZipCode, "-") = 0 then
					oMaster.FReceiveZipCode = Left(oMaster.FReceiveZipCode,3) & "-" & Mid(oMaster.FReceiveZipCode,4,10)
				end if

				'// 주소 수정
				oMaster.FReceiveAddr1 = TRIM(Replace(oMaster.FReceiveAddr1,"  "," "))
				oMaster.FReceiveAddr2 = TRIM(Replace(oMaster.FReceiveAddr2,"  "," "))
				tmpStr = oMaster.FReceiveAddr1 & " " & oMaster.FReceiveAddr2
				pos = 0
				for k = 0 to 2
					pos = InStr(pos+1, tmpStr, " ")
					if (pos = 0) then
						exit for
					end if
				next

				if (pos > 0) then
					oMaster.FReceiveAddr1 = Left(tmpStr, pos)
					oMaster.FReceiveAddr2 = Mid(tmpStr, pos+1, 1000)
				end if

				oMaster.FReceiveAddr1 = Trim(oMaster.FReceiveAddr1)
				oMaster.FReceiveAddr2 = Trim(oMaster.FReceiveAddr2)
			End If

			redim oDetailArr(0)
			Set oDetailArr(0) = new COrderDetail
			oDetailArr(0).FdetailSeq = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOrderID")(0).Text
			oDetailArr(0).FItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode").length > 0) then
				oDetailArr(0).FItemOption = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:OptionManageCode")(0).Text
			else
				oDetailArr(0).FItemOption = "0000"
			end if

			oDetailArr(0).FOutMallItemID = objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductID")(0).Text
			oDetailArr(0).FOutMallItemOption = oDetailArr(0).FItemOption
			oDetailArr(0).FOutMallItemName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductName")(0).Text)
			if (objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption").length > 0) then
				oDetailArr(0).FOutMallItemOptionName = html2db(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductOption")(0).Text)
			else
				oDetailArr(0).FOutMallItemOptionName = ""
			end if

			oDetailArr(0).FItemNo = CLng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:Quantity")(0).Text)

			'2019-08-06 김진영 아래 조건 추가
			'스토어팜 매입이면서 할인기간이라면 판매가(itemcost)를 실판매가(reducedprice)와 동일하게 저장
			'If left(now(),10) >= "2019-10-2" and left(now(),10) < "2019-09-24" Then
			'2019-10-21 김진영, 위 now()에서 Date로 변경 / Case SellerProductCode CSTR문자 변환, Trim 처리
			'2020-09-10 김진영, 스토어팜 특가관리에 추가했다면 할인가격으로 변경되게 수정
			strSql = ""
			strSql = strSql & " SELECT COUNT(*) as cnt "
			strSql = strSql & " FROM db_etcmall.dbo.tbl_outmall_mustPriceItem "
			strSql = strSql & " WHERE mallgubun = '"& sellsite &"' "
			strSql = strSql & " and itemid = '"& CSTR(Trim(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:SellerProductCode")(0).Text)) &"' "
			strSql = strSql & " and '"& storeOrderDate &"' BETWEEN startdate AND enddate "
			'rw strSql
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
				If rsget("cnt") > 0 Then
					isDisCountYn = "Y"
				Else
					isDisCountYn = "N"
				End If
			rsget.Close

			If isDisCountYn = "Y" Then
'				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 시작
				maySellPrice = Clng(objMasterOneXML.getElementsByTagName("n:UnitPrice")(0).Text) + Clng(objMasterOneXML.getElementsByTagName("n:OptionPrice")(0).Text)
				If (objMasterOneXML.getElementsByTagName("n:ProductImediateDiscountAmount").length > 0) then
					oDetailArr(0).Fitemcost = maySellPrice - Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:ProductImediateDiscountAmount")(0).Text) / oDetailArr(0).FItemNo)
				Else
					oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
				End If
'######## 2020-10-08 김진영 // 할인상품 판매가 아래처럼 수정 끝
			Else
				oDetailArr(0).Fitemcost = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalProductAmount")(0).Text) / oDetailArr(0).FItemNo)
			End If

			oDetailArr(0).FReducedPrice = Round(Clng(objMasterOneXML.getElementsByTagName("n:ProductOrder/n:TotalPaymentAmount")(0).Text) / oDetailArr(0).FItemNo)
			oDetailArr(0).FOutMallCouponPrice = 0
			oDetailArr(0).FTenCouponPrice = 0


			if (SaveOrderToDB(oMaster, oDetailArr) = True) then
				if PlaceProductOrder_nvstorefarm(oDetailArr(0).FdetailSeq, sellsite) then
					successCnt = successCnt + 1
				end if
			end if
			i = i + 1
		end if
	next
	Set cryptoLib = Nothing

	''if IsAutoScript then
		response.write "Order Insert(" & successCnt & ")" & "<br />"
	''end if

	GetOrderFrom_NewCall_nvstorefarm = True
	Set xmlDOM = Nothing
	Set objXML = Nothing
End Function

Function getStorefarmOrderNumUpd()
	Dim sqlStr
	sqlStr = "exec [db_temp].[dbo].[usp_Ten_Nvstorefarm_Num_Upd]"
	dbget.Execute sqlStr
End Function

Function getMaxPageStorefarm()
	Dim sqlStr
	sqlStr = ""
	sqlStr = sqlStr & " select max(num) as maxnum "
	sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_TMPOrder_storefarm "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If (Not rsget.EOF) Then
		getMaxPageStorefarm = rsget("maxnum")
	Else
		getMaxPageStorefarm = 0
	End If
	rsget.Close
End Function


function GetCheckStatus(byVal sellsite, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp](sellsite, lastcheckdate, issuccess) "
	strSql = strSql + "		values('" & sellsite & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N') "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function GetCheckItemOptionValid(byVal itemid, byVal itemoption)
	dim strSql

	GetCheckItemOptionValid = False

    strSql = " select top 1 i.itemid "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    strSql = strSql + " 	and o.itemoption = '" & itemoption & "' "

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetCheckItemOptionValid = True
	end if
	rsget.Close
end function

function GetItemOptionWithOptionName(byVal sellsite, byVal itemid, byVal itemoptionname)
	dim strSql, found

	found = False
	GetItemOptionWithOptionName = "0000"



	'// 모델명:SMN-204 you're in
	itemoptionname = Replace(itemoptionname, "'", "''")


	if (sellsite = "ezwel") then
		strSql = "exec [db_temp].[dbo].[usp_TEN_xSiteOrder_OptionMapping_EzWel] '"&itemid&"','"&itemoptionname&"'"

		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			GetItemOptionWithOptionName = rsget("itemoption")
			found = True
		end if
		rsget.Close
	end if

	if found then
		exit function
	end if

    strSql = " select top 1 o.itemoption "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    strSql = strSql + " 	and o.optionname = '" & itemoptionname & "' "
    ''response.Write strSql

	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetItemOptionWithOptionName = rsget("itemoption")
		found = True
	end if
	rsget.Close

	if found then
		exit function
	end if

	'사방넷을 통한 LFmall 옵션명에  None[XX]: 글자가 심어서 넘어옴
	If Instr(itemoptionname, "None[XX]:") > 0 Then
		itemoptionname = Replace(itemoptionname, "None[XX]:", "")
	End If

    strSql = " select top 1 o.itemoption "
    strSql = strSql + " from "
    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
    strSql = strSql + " 	on "
    strSql = strSql + " 		i.itemid = o.itemid "
    strSql = strSql + " where "
    strSql = strSql + " 	1 = 1 "
    strSql = strSql + " 	and i.itemid = " & itemid
    ''strSql = strSql + " 	and (o.optionname = '" & Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/") & "') "
    strSql = strSql + " 	and (Replace(Replace(o.optionname, ',', ''), ':', '') = '" & Replace(Replace(Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/"), ",", ""), ":", "") & "') "

	rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	if Not rsget.Eof then
		GetItemOptionWithOptionName = rsget("itemoption")
		found = True
	end if
	rsget.Close

	if found then
		exit function
	end if

	if (sellsite = "lotteCom") and False then
	    strSql = " select top 1 o.itemoption "
	    strSql = strSql + " from "
	    strSql = strSql + " 	[db_item].[dbo].[tbl_item] i "
	    strSql = strSql + " 	join [db_item].[dbo].[tbl_item_option] o "
	    strSql = strSql + " 	on "
	    strSql = strSql + " 		i.itemid = o.itemid "
	    strSql = strSql + " where "
	    strSql = strSql + " 	1 = 1 "
	    strSql = strSql + " 	and i.itemid = " & itemid
	    strSql = strSql + " 	and (Replace(Replace(o.optionname, ',', ''), ':', '') = '" & Replace(Replace(Replace(Replace(itemoptionname, "&amp;", "&"), "&times;", "/"), ",", ""), ":", "") & "') "

		rsget.CursorLocation = adUseClient
    	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			GetItemOptionWithOptionName = rsget("itemoption")
			found = True
		end if
		rsget.Close

		if found then
			exit function
		end if
	end if

end function

function SetCheckStatus(sellsite, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "' "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

function arrayMerge(left, right)
	dim right_size
	dim total_size
	dim i
	dim merged
	''// Convert "left" to an array
	if not isArray(left) then
		left = Array(left)
	end if
	''// Convert "right" to an array
	if not isArray(right) then
		right = Array(right)
	end if
	''// Start with "left" and add the elements of "right"

	right_size = ubound(right)
	total_size = ubound(left) + right_size + 1

	merged = array()
	redim merged(total_size)
	dim counter : counter = 0

	for i = lbound(left) to ubound(left)
		if isobject(left(i))then
			set merged(counter) = left(i)
		else
			merged(counter) = left(i)
		end if
		counter=counter+1
	next

	for i = lbound(right) to ubound(right)
		if isobject(right(i))then
			set merged(counter) = right(i)
		else
			merged(counter) = right(i)
		end if
	next


	''// Return value
	arrayMerge = merged
end function


public function getDelimCharCount(orgStr, delim)
    dim retCNT : retCNT = 0
    dim buf
    buf = split(orgStr,delim)

    if IsArray(buf) then
        retCNT = UBound(buf)
    end if
    getDelimCharCount = retCNT
end function

'' SSG 매칭된 옵션코드 리턴 [빨강/XL/소재/3] [화이트] [빨강/L,1:^:주문제작문구:^:문구작성] [,1:^:주문제작문구:^:주문문구1,2:^:주문제작문구:^:주문문구2]  '' spliter ,  / :^:
function getOptionCodByOptionNameSSG(iitemid,outmalloptionName,byref requiredtl, uitemid)
    dim retStr, sqlStr : retStr=""
    dim ichrCnt, IsDoubleOption, IsTreepleOption
    dim ioptionname, ireqdrlname

    if (outmalloptionName="") then
        requiredtl = ""
        getOptionCodByOptionNameSSG = "0000"
        Exit function
    end if

    ' if (outmalloptionName="") then
    '     requiredtl = ""
    '     ' getOptionCodByOptionNameSSG = "0000"

    '     sqlStr = "select top 1 itemoption"
	' 	sqlStr = sqlStr & " from db_item.dbo.tbl_OutMall_regedoption"
	' 	sqlStr = sqlStr & " where mallid = 'ssg'"
	' 	sqlStr = sqlStr & " and itemid ="&iitemid
	' 	sqlStr = sqlStr & " and outmallOptCode = '"&uitemid&"'"
	' 	rsget.CursorLocation = adUseClient
	' 	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	' 	if (Not rsget.EOF) then
	' 		getOptionCodByOptionNameSSG = rsget("itemoption")
	' 	Else
	' 		getOptionCodByOptionNameSSG = "0000"
	' 	end if
	' 	rsget.Close
    '     Exit function
    ' end if

    ioptionname = outmalloptionName
    ichrCnt = getDelimCharCount(ioptionname,",")

''////////////////////////////////////////////////////// 예전 버전 ////////////////////////////////////////////////////////
'     IF (ichrCnt>=1) THEN ''주문제작 문구가 있는 상품
'         ioptionname = split(outmalloptionName,",")(0)
'         requiredtl  = replace(split(outmalloptionName,",")(1),"1:^:주문제작문구:^:","")
'         ''requiredtl  = replace(split(outmalloptionName,",")(1),"1:^:asdasd:^:","")

'         if ichrCnt>1 then
'             requiredtl = requiredtl + ","+replace(split(outmalloptionName,",")(2),"2:^:주문제작문구:^:","")
'             ''requiredtl = requiredtl + ","+replace(split(outmalloptionName,",")(2),"2:^:asdasdddd:^:","")
'         end if
'   ''rw "[requiredtl]"&requiredtl
'         'rw "ioptionname:"&ioptionname
'         'rw "requiredtl:"&requiredtl
'     end if
''////////////////////////// 수정 버전 2019-12-11 11:40 김진영 수정 주문번호 :19120982908 문제 발생  ////////////////////////////
    IF (ichrCnt>=1) THEN ''주문제작 문구가 있는 상품
        ioptionname = split(outmalloptionName,",")(0)
		If instr(outmalloptionName, "1:^:주문제작문구:^:") > 0 Then
			requiredtl = Split(outmalloptionName, "1:^:주문제작문구:^:")(1)
			If instr(requiredtl, "2:^:주문제작문구:^:") > 0 then
				requiredtl = Replace(requiredtl, "2:^:주문제작문구:^:", "")
			end if
		End If
    end if
''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if (ioptionname="") then  ''주문제작문구를 발라낸후 옵션명이 없으면.
        getOptionCodByOptionNameSSG = "0000"
        Exit function
    end if

    IF (getDelimCharCount(ioptionname,"/")=1) THEN
        IsDoubleOption = TRUE
    ELSEIF (getDelimCharCount(ioptionname,"/")=2) THEN  '''빨강/XL/소재/3 = 옵션명에 / 가 있을경우 못발라낼 수 있음.(소재/3)
        IsTreepleOption = TRUE
    ENd IF


    ioptionname= replace(ioptionname,"'","''")   '' like this CASE : 모델명:SMN-204 you're in
    IF (IsDoubleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and optionname='"&replace(ioptionname,"/",",")&"'"   ''replace(optionname,'*','')
    ELSEIF (IsTreepleOption) THEN
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and optionTypename='복합옵션'"
        sqlStr = sqlStr & " and optionname='"&replace(ioptionname,"/",",")&"'"   ''replace(optionname,'*','')
    ELSE
        sqlStr = "select top 1 itemoption "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        sqlStr = sqlStr & " and Replace(optionname,',','')='"&ioptionname&"'"
    END IF

''response.write sqlstr & "<Br>"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if (Not rsget.EOF) then
	    retStr = rsget("itemoption")
	end if
    rsget.Close

	''옵션명에 "/" 가 있는 CASE===============================================================================
	If (retStr="") THEN
		sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteOrder_OptionMapping_SSG] '"&iitemid&"','"&replace(Trim(outmalloptionName),"'","''")&"'"

		rsget.CursorLocation = adUseClient
		rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
		if Not rsget.Eof then
			retStr = rsget("itemoption")
		end if
		rsget.Close
	END IF
	''=====================================================================================================

    If (retStr="") THEN
       ''옵션 매칭이 안되었을때. 수기매칭으로 진행 ?  0000 맞나?
        sqlStr = "select count(*) as CNT "
        sqlStr = sqlStr & " from db_item.dbo.tbl_item_option "
        sqlStr = sqlStr & " where itemid="&iitemid&VbcrLF
        rsget.CursorLocation = adUseClient
        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
        if (Not rsget.EOF) then
    	    if (rsget("CNT")>0) THEN
    	        retStr = "FF00" '"0000"=>FF00
    	    else
    	        retStr = "0000"
    	    end if
    	end if
        rsget.Close
    END IF

    getOptionCodByOptionNameSSG = retStr
end function

function RemoveWhiteSpaceChar(str)
	dim retVal
	If isNull(str) Then
		RemoveWhiteSpaceChar = ""
		Exit Function
	End If

	retVal = str
	retVal = Replace(retVal, Chr(13), "")
	retVal = Replace(retVal, Chr(10), "")
	retVal = Replace(retVal, vbTab, " ")
	retVal = Trim(retVal)
	RemoveWhiteSpaceChar = retVal
end function

Function getApiUrl(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiUrl = "https://dev-openapi.lotteon.com"
			Else
				getApiUrl = "https://openapi.lotteon.com"
			End If
	End Select
End Function

Function getApiKey(mallid)
	Select Case mallid
		Case "lotteon"
			If application("Svr_Info") = "Dev" Then
				getApiKey = "5d5b2cb498f3d20001665f4e5451c4d923ac4e2c95df619996f35476"
			Else
				getApiKey = "5d5b2cb498f3d20001665f4e18a41621005d4c1ba262804ec7a10732"
			End If
	End Select
End Function

Function getShintvshoppingGubun1
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT outmallorderserial, orderGSeq, orderDSeq, orderWSeq, shintvshoppingGoodNo, outmallOptCode "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] "
	strSql = strSql & " WHERE sellsite = 'shintvshopping' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getShintvshoppingGubun1 = rsget.getRows()
	End If
	rsget.Close
End Function

Function getSkstoaGubun1
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT outmallorderserial, orderGSeq, orderDSeq, orderWSeq, shintvshoppingGoodNo, outmallOptCode "
	strSql = strSql & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_shintvshopping] "
	strSql = strSql & " WHERE sellsite = 'skstoa' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If not rsget.EOF Then
		getSkstoaGubun1 = rsget.getRows()
	End If
	rsget.Close
End Function

function skstoaAPIURL()
	If application("Svr_Info") = "Dev" Then
		skstoaAPIURL = "http://dev-sel.skstoa.com"
	Else
		skstoaAPIURL = "https://open-api.skstoa.com"
	End If
end function

function skstoalinkCode()
	If application("Svr_Info") = "Dev" Then
		skstoalinkCode = "TENBYTEN"
	Else
		skstoalinkCode = "TENBYTEN"
	End If
end function

function skstoaentpCode()
	If application("Svr_Info") = "Dev" Then
		skstoaentpCode = "112644"
	Else
		skstoaentpCode = "112644"
	End If
end function

function skstoaentpId()
	If application("Svr_Info") = "Dev" Then
		skstoaentpId = "E112644"
	Else
		skstoaentpId = "E112644"
	End If
end function

function skstoaentpPass()
	Dim skstoaStrSql
	skstoaStrSql = ""
	skstoaStrSql = skstoaStrSql & " SELECT TOP 1 isnull(iniVal, '') as iniVal "
	skstoaStrSql = skstoaStrSql & " FROM db_etcmall.dbo.tbl_outmall_ini " & VbCRLF
	skstoaStrSql = skstoaStrSql & " where mallid='skstoa' " & VbCRLF
	skstoaStrSql = skstoaStrSql & " and inikey='pass'"
	rsget.CursorLocation = adUseClient
	rsget.Open skstoaStrSql, dbget, adOpenForwardOnly, adLockReadOnly
	if not rsget.Eof then
		skstoaentpPass	= rsget("iniVal")
	end if
	rsget.close
end function

%>
