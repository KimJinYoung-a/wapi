<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 XML 주문처리
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteOrderXMLCls.asp"-->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/hmall/hmallItemcls.asp"-->
<!-- #include virtual="/outmall/hmall/inchmallFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Function fnHmallConfirmOrder(vOrderserial, vOrgDetailKey, vBeasongNum11st)
	Dim objXML, xmlDOM, iRbody, strSql, istrParam, iDlvstNo, iDlvstPtcSeq
	iDlvstNo		= Trim(Split(vBeasongNum11st, "!_!")(0))
	iDlvstPtcSeq	= Trim(Split(vBeasongNum11st, "!_!")(1))
	'ProcGb | P1:주문확인, P2:출고완료, P3:배송완료
	''istrParam = "DlvstNo="&iDlvstNo&"&DlvstPtcSeq=" & iDlvstPtcSeq & "&OrdNo=" & vOrderserial & "&OrdPtcSeq=" & vOrgDetailKey & "&ProcGb=P1&DsrvDlvcoCd=&InvcNo="

    istrParam = ""
    istrParam = istrParam & "{"
    istrParam = istrParam & "  ""DlvstNo"": """ & iDlvstNo & ""","
    istrParam = istrParam & "  ""DlvstPtcSeq"": """ & iDlvstPtcSeq & ""","
    istrParam = istrParam & "  ""OrdNo"": """ & vOrderserial & ""","
    istrParam = istrParam & "  ""OrdPtcSeq"": """ & vOrgDetailKey & ""","
    istrParam = istrParam & "  ""ProcGb"": ""P1"","
    istrParam = istrParam & "  ""DsrvDlvcoCd"": """","
    istrParam = istrParam & "  ""InvcNo"": """""
    istrParam = istrParam & "}"

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "http://xapi.10x10.co.kr:8080/Orders/Hmall/actionoutput", false
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)

		If Err.number <> 0 Then
			iErrStr = ivendorItemId
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		rw "###############"

		If objXML.Status = "200" OR objXML.Status = "201" Then
			strSql = ""
			strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
			strSql = strSql & " isbaljuConfirmSend = 'Y' "
			strSql = strSql & " , lastUpdate = getdate() "
			strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
			strSql = strSql & " and beasongNum11st = '"&vBeasongNum11st&"' "
			strSql = strSql & " and OrgDetailKey = '"&vOrgDetailKey&"' "
			strSql = strSql & " and mallid = 'hmall1010' "
			dbget.Execute strSql
			fnHmallConfirmOrder= true
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

Function getTenOptionCode(iitemid, ipartnerOptionName)
	Dim strSql, retOptionCode, mayOptTypeName, maySingleOption
	maySingleOption = "N"

	If ipartnerOptionName = "단일옵션" Then
		retOptionCode = "0000"
	Else
		mayOptTypeName = Trim(Split(ipartnerOptionName, "/")(0))

		strSql = ""
		strSql = strSql & " SELECT COUNT(*) as Cnt"
		strSql = strSql & " FROM db_item.dbo.tbl_item_option "
		strSql = strSql & " WHERE itemid = '"& iitemid &"' "
		strSql = strSql & " and optionTypeName = '"& mayOptTypeName &"' "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If rsget("Cnt") > 0 Then
			maySingleOption = "Y"
		End If
		rsget.Close

		If maySingleOption = "Y" Then
			strSql = ""
			strSql = strSql & " SELECT itemoption "
			strSql = strSql & " FROM db_item.dbo.tbl_item_option "
			strSql = strSql & " WHERE itemid = '"& iitemid &"' "
			strSql = strSql & " and optionname = '"& Trim(Split(ipartnerOptionName, "/")(1)) &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				retOptionCode = rsget("itemoption")
			End If
			rsget.Close
		Else
			strSql = ""
			strSql = strSql & " SELECT itemoption "
			strSql = strSql & " FROM db_item.dbo.tbl_item_option "
			strSql = strSql & " WHERE itemid = '"& iitemid &"' "
			strSql = strSql & " and optionname = '"& REPLACE(ipartnerOptionName, "/", ",") &"' "
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not(rsget.EOF or rsget.BOF) then
				retOptionCode = rsget("itemoption")
			End If
			rsget.Close
		End If
	End If

	If retOptionCode = "" Then
		retOptionCode = "0000"
	End If

	getTenOptionCode = retOptionCode
End Function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr, beasongNum11st, reserve01, outMallOptionNo)
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
		,Array("@reserve01"	,adVarchar, adParamInput,32, reserve01) _
		,Array("@beasongNum11st"	,adVarchar, adParamInput,16, beasongNum11st) _
		,Array("@outMallOptionNo"	,adVarchar, adParamInput,16, outMallOptionNo) _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.[dbo].[usp_API_Hmall_OrderReg_Add]"
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

Dim sqlStr, buf, i, j, mode, sellsite
Dim divcd, idx
Dim objXML, xmlDOM, retCode, iMessage
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)

Dim strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim OrgDetailKey, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2, OrderName, OrderTelNo, OrderHpNo
Dim OutMallOrderSerial, SellDate, outMallGoodsNo, matchItemID, partnerItemName, SellPrice, RealSellPrice, ItemOrderCount, orderDlvPay, requireDetail, matchItemOption, outMallOptionNo
Dim partnerOptionName, SalePrice, beasongNum11st, reserve01
Dim regOrderCnt, strObj, iRbody
Dim iSellDate, iIsSuccess, fromDate, nowDate, searchDate, orderCount

Dim dlvstNo, dlvstPtcSeq, ordNo, lastDlvstPrgrGbcd, dlvTypeGbcd, POS1, POS2, POS3, ReceiveAddr, dlvCnclYn

Call GetCheckStatus("hmall1010", iSellDate, iIsSuccess)
searchDate = replace(iSellDate, "-", "")
rw searchDate & " Order START"
' searchDate = "20231118"
If sellsite = "hmall1010" Then
	Dim istrParam
		istrParam = ""
		istrParam = istrParam & "<?xml version=""1.0"" encoding=""utf-8""?>"
		istrParam = istrParam & "<Root xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		istrParam = istrParam & "	<Dataset id=""dsInput"">"
		istrParam = istrParam & "		<rows>"
		istrParam = istrParam & "			<row>"
		istrParam = istrParam & "				<venCd>002569</venCd>"
		istrParam = istrParam & "				<fromDate>"&searchDate&"</fromDate>"
		istrParam = istrParam & "				<toDate>"&searchDate&"</toDate>"
		istrParam = istrParam & "				<prgrGb>P0</prgrGb>"
		istrParam = istrParam & "			</row>"
		istrParam = istrParam & "		</rows>"
		istrParam = istrParam & "	</Dataset>"
		istrParam = istrParam & "</Root>"

'	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", "https://openapi.hmall.com/front/sc/scb/scbd/selectOshpDtlList.do", false
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "oauserId", "002569"
		objXML.setRequestHeader "oauseKey", "23439A336B4FC812A1ED415489F185A2"
		objXML.Send(istrParam)

		If objXML.Status = "200" OR objXML.Status = "201" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = objXML.ResponseText
				xmlDOM.LoadXML iRbody
				response.write "req : <textarea cols=40 rows=10>"&istrParam&"</textarea>"
				response.write "res : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				Dim Nodes, SubNodes
				If xmlDOM.getElementsByTagName("Dataset").item(1).attributes(0).nodeValue = "dsOutput" Then
					If xmlDOM.selectNodes("//Dataset[@id='dsOutput']/rows/row/dlvstNo").length > 0 Then
						Set Nodes = xmlDOM.selectNodes("//Dataset[@id='dsOutput']/rows/row")
							For each SubNodes in Nodes
								ReceiveAddr = ""
								ReceiveAddr1 = ""
								ReceiveAddr2 = ""

								orderCsGbn			= 0
								beasongNum11st		= SubNodes.selectSingleNode("dlvstNo").Text				'배송지시번호
								reserve01			= SubNodes.selectSingleNode("dlvstPtcSeq").Text			'배송지시상세번호
								OutMallOrderSerial	= SubNodes.selectSingleNode("ordNo").Text				'주문번호
								OrgDetailKey		= SubNodes.selectSingleNode("ordPtcSeq").Text			'주문일련번호
								outMallGoodsNo		= SubNodes.selectSingleNode("slitmCd").Text				'판매상품코드
								partnerItemName		= SubNodes.selectSingleNode("slitmNm").Text				'상품명
								outMallOptionNo		= SubNodes.selectSingleNode("uitmCd").Text				'상품속성코드
								partnerOptionName	= SubNodes.selectSingleNode("uitmTotNm").Text			'상품속성명
								lastDlvstPrgrGbcd	= SubNodes.selectSingleNode("lastDlvstPrgrGbcd").Text	'최종배송지시진행구분코드 | 25:출고대기, 30:출고진행, 45:출고, 50:배송완료
								dlvCnclYn			= SubNodes.selectSingleNode("dlvCnclYn").Text			'배송취소여부
								SellPrice			= SubNodes.selectSingleNode("sellUprc").Text				'판매단가
								RealSellPrice		= SubNodes.selectSingleNode("sellUprc").Text				'진영Comment : 실 판매가가 안 넘어옴...그냥 판매가로 넣음

								dlvTypeGbcd 		= SubNodes.selectSingleNode("dlvTypeGbcd").Text			'배송유형(주문구분) | 10:주문출력; 40:교환배송
								ReceiveName			= Left(SubNodes.selectSingleNode("rcvrNm").Text, 28)		'인수자명 | 수취자명
								ReceiveHpNo			= SubNodes.selectSingleNode("rcvrHp").Text				'인수자전화 | 수취자전화번호(Astrisk)
								ReceiveTelNo		= SubNodes.selectSingleNode("rcvrTel").Text				'인수자전화 | 수취자전화번호(Astrisk)
								OrderName			= Left(SubNodes.selectSingleNode("dlvApltNm").Text, 28)	'주문자 | 배송신청자명(Astrisk)
								OrderHpNo			= SubNodes.selectSingleNode("dlvApltTel").Text			'주문자전화 | 배송신청자전화번호(Astrisk)
								OrderTelNo			= SubNodes.selectSingleNode("dlvApltTel").Text			'주문자전화 | 배송신청자전화번호(Astrisk)
								ReceiveZipCode		= SubNodes.selectSingleNode("dstnPostNo").Text			'배송지우편번호 | (astrisk)
								ReceiveAddr			= SubNodes.selectSingleNode("dstnAdr").Text				'배달지 | 배송지주소(astrisk)

								'''주소와 상세주소가 같은경우 3번째 Blank에서 끊음.
								POS1 = 0
								POS2 = 0
								POS3 = 0
								POS1 = InStr(ReceiveAddr, " ")
								If (POS1 > 0) Then
									POS2 = InStr(MID(ReceiveAddr, POS1+1, 512)," ")
									If POS2>0 Then
										POS3 = InStr(MID(ReceiveAddr, POS1 + POS2 + 1 ,512)," ")
										If POS3 > 0 Then
											ReceiveAddr1 = LEFT(ReceiveAddr, POS1 + POS2 + POS3 - 1)
											ReceiveAddr2 = MID(ReceiveAddr, POS1 + POS2 + POS3 + 1, 512)
										End If
									End If
								End If

								SellDate			= LEFT(SubNodes.selectSingleNode("ptcOrdDtm").Text, 10)	'상세주문일자
								deliverymemo		= SubNodes.selectSingleNode("dlvPaonMsg").Text			'배송메세지
								matchItemID			= SubNodes.selectSingleNode("venItemCd").Text			'협력사 상품관리코드
								matchItemID 		= replace(matchItemID, "TEST_", "")
								ItemOrderCount		= SubNodes.selectSingleNode("ordQty").Text				'주문 수량
								matchItemOption		= getTenOptionCode(matchItemID, partnerOptionName)

								If (dlvCnclYn <> "Y") AND (dlvTypeGbcd <> "40") Then	'배송취소여부가 Y가 아니고, 교환주문이 아니면 저장
									retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
											, OrderName, OrderTelNo, OrderHpNo _
											, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
											, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
											, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
											, errCode, errStr, beasongNum11st, reserve01, outMallOptionNo)

									If (retVal) Then
										succCNT = succCNT + 1
										strsql = ""
										strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
										strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', '" & beasongNum11st & "!_!" & reserve01 & "', 'N', getdate(), 'hmall1010')"
										dbget.Execute strSql
									Else
										failCNT = failCNT + 1
									End If
								End If
							Next
						Set Nodes = nothing
					End If
				End If
			Set xmlDOM = nothing

			If (failCNT <> 0) Then
			    rw "["&failCNT&"] 건 실패(주문조회)"
			End if

			If (succCNT <> 0) then
			    rw "["&succCNT&"] 건 성공(주문조회)"
			    Dim arrList, lp, ret1
			    Dim OKcnt, NOcnt
			    OKcnt = 0
			    NOcnt = 0

				strsql = ""
				strsql = strsql & " update T "
				strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				strsql = strsql & " From db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T "
				strsql = strsql & " JOIN db_temp.dbo.tbl_xsite_tmporder as O on T.outmallorderserial = O.OutMallOrderSerial and T.OrgDetailKey = O.OrgDetailKey "
				strsql = strsql & " where T.isbaljuConfirmSend <> 'Y' "
				strsql = strsql & " and O.sendState = 1 "
				strsql = strsql & " and O.matchstate in ('O') "
				strsql = strsql & " and T.mallid = 'hmall1010' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " update T "
				strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
				strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
				strsql = strsql & " WHERE M.cancelyn ='Y' "
				strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
				strsql = strsql & " and T.mallid = 'hmall1010' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " SELECT TOP 1000 outmallorderserial, OrgDetailKey, beasongNum11st FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
				strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
				strsql = strsql & " and mallid = 'hmall1010' "
				rsget.CursorLocation = adUseClient
				rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			    if not rsget.Eof then
			        arrList = rsget.getRows()
			    end if
			    rsget.close

				For lp = 0 To Ubound(arrList, 2)
					ret1 = fnHmallConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp))

	                If (ret1) then
	                    OKcnt = OKcnt + 1
	                Else
	                    NOcnt = NOcnt + 1
	                End If
				Next

				If OKcnt <> 0 then
					rw "["&OKcnt&"] 건 성공(발주확인)"
				End If

				If NOcnt <> 0 then
					rw "["&NOcnt&"] 건 실패(발주확인)"
				End If
			End If
'			response.end
			If (iSelldate < Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
			ElseIf (iSellDate = Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, iSellDate, "Y")
			End If
		Else
			rw "주문연동 실패..잠시 후 시도 요망"
		End If
	Set objXML = nothing
End If
rw searchDate & " Order End"

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
