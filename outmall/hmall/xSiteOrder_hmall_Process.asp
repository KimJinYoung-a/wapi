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
'searchDate = "20181119"
If sellsite = "hmall1010" Then
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'prgrGb | P0:출고대기, P1:출고진행, P2:출고, P3:배송완료
		objXML.open "GET", "http://xapi.10x10.co.kr:8080/Orders/Hmall/output?startdate="&searchDate&"&enddate="&searchDate&"&prgrGb=P0", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||" & Err.Description
		End If
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Dim obj1
			Set strObj = JSON.parse(iRbody)
				orderCount = strObj.count
				If orderCount > 0 Then
					set obj1 = strObj.lstorder
						for i=0 to obj1.length-1
							ReceiveAddr = ""
							ReceiveAddr1 = ""
							ReceiveAddr2 = ""

							orderCsGbn			= 0
'							rw obj1.get(i).editYn									'수정가능여부(택배사,운송장) | "Y:협력사직송(40) AND 출고진행(30) AND 운송장번호null AND 배송취소여부N"
							beasongNum11st		= obj1.get(i).dlvstNo				'배송지시번호
							reserve01			= obj1.get(i).dlvstPtcSeq			'배송지시상세번호
							OutMallOrderSerial	= obj1.get(i).ordNo					'주문번호
							OrgDetailKey		= obj1.get(i).ordPtcSeq				'주문일련번호
'							rw obj1.get(i).sitmCd									'단축상품코드
							outMallGoodsNo		= obj1.get(i).slitmCd				'판매상품코드
							partnerItemName		= obj1.get(i).slitmNm				'상품명
							outMallOptionNo		= obj1.get(i).uitmCd				'상품속성코드
							partnerOptionName	= obj1.get(i).uitmTotNm				'상품속성명
'							rw obj1.get(i).dlvFormGbcd								'배송형태 | 20:현대홈직택배, 30:협력사직택배, 40:협력사직송
							lastDlvstPrgrGbcd	= obj1.get(i).lastDlvstPrgrGbcd		'최종배송지시진행구분코드 | 25:출고대기, 30:출고진행, 45:출고, 50:배송완료
'							rw obj1.get(i).lastOshpDlineDt							'최종출고마감일자
'							rw obj1.get(i).oshpReqnDt								'출고요청일
							dlvCnclYn			= obj1.get(i).dlvCnclYn				'배송취소여부
'							rw obj1.get(i).dlvCnclNm								'배송취소
'							rw obj1.get(i).dlvstQty									'배송지시수량
'							rw obj1.get(i).unitDlvQty								'출고수량
'							rw obj1.get(i).prrgQty									'대상수량
'							rw obj1.get(i).custCnclQty								'고객취소수량
							SellPrice			= obj1.get(i).sellUprc				'판매단가
							RealSellPrice		= obj1.get(i).sellUprc				'진영Comment : 실 판매가가 안 넘어옴...그냥 판매가로 넣음
'							rw obj1.get(i).sellSum									'판매가 (진영 코맨트 : 판매합계가 오는 듯)
'							rw obj1.get(i).prchUprcSum								'매입가
'							rw obj1.get(i).dsrvDlvcoCd								'택배배송사코드
'							rw obj1.get(i).invcNo									'운송장번호
'							rw obj1.get(i).dsntDtDlvYn								'지정일자배송여부
'							rw obj1.get(i).rsvSellYn								'예약판매여부
'							rw obj1.get(i).venCd									'협력사
'							rw obj1.get(i).ven2Cd									'2차협력사
							dlvTypeGbcd = obj1.get(i).dlvTypeGbcd					'배송유형(주문구분) | 10:주문출력; 40:교환배송
'							rw obj1.get(i).dlvTypeGbcdColor							'배송유형표시색상
'							rw obj1.get(i).dlvcPayGbcd								'배송비지불구분코드 | 00:무료, 10:선결제, 20:착불, 30:설치상품
'							rw obj1.get(i).befDlvstNo								'이전배송지시번호
'							rw obj1.get(i).befDlvstPtcSeq							'이전배송지시상세순번
'							rw obj1.get(i).ordSplpnAplyYn							'주문공급계획적용여부
'							rw obj1.get(i).custDlvHopeDt							'고객배송희망일자
'							rw obj1.get(i).oshpCnfmDtm								'출고확정일시
							ReceiveName			= Left(obj1.get(i).rcvrNm, 28)		'인수자명 | 수취자명
							ReceiveHpNo			= obj1.get(i).rcvrTel				'인수자전화 | 수취자전화번호(Astrisk)
							ReceiveTelNo		= obj1.get(i).rcvrTel				'인수자전화 | 수취자전화번호(Astrisk)
							OrderName			= Left(obj1.get(i).dlvApltNm, 28)	'주문자 | 배송신청자명(Astrisk)
							OrderHpNo			= obj1.get(i).dlvApltTel			'주문자전화 | 배송신청자전화번호(Astrisk)
							OrderTelNo			= obj1.get(i).dlvApltTel			'주문자전화 | 배송신청자전화번호(Astrisk)
							ReceiveZipCode		= obj1.get(i).dstnPostNo			'배송지우편번호 | (astrisk)
							ReceiveAddr			= obj1.get(i).dstnAdr				'배달지 | 배송지주소(astrisk)

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

'							rw obj1.get(i).ordCustNo								'주문고객번호
'							rw obj1.get(i).custDstnSeq								'고객배송지순번
'							rw obj1.get(i).rcvrPaonMsg								'전하는말 | 수취자전달메시지
'							rw obj1.get(i).befInvcNo								'원송장번호
'							rw obj1.get(i).prsnYn									'선물포장여부 | 선물여부
'							rw obj1.get(i).frdlvYn									'해외배송여부
'							rw obj1.get(i).frgnDstnSeq								'해외배송지순번
'							rw obj1.get(i).custVenPaonMsg							'주문시요청사항 | 고객협력사전달메시지
'							rw obj1.get(i).frgnAdr									'해외배송지
'							rw obj1.get(i).hmallGiftRfrNote							'hmall사은품참조사항
'							rw obj1.get(i).dlvNo									'배송번호
'							rw obj1.get(i).dlvPtcSeq								'배송상세순번
							SellDate = LEFT(obj1.get(i).ptcOrdDtm, 10)				'상세주문일자
'							rw obj1.get(i).cvstDlvYn								'편의점배송여부 | 편의점배송여부( Y or N)
'							rw obj1.get(i).inslItemYn								'설치상품여부 | INSL_ITEM_YN
'							rw obj1.get(i).almlBasktNo								'제휴장바구니번호
'							rw obj1.get(i).nshipTypeGbcd							'미출고유형구분코드 | 출고지연:10, 품절취소:20
							deliverymemo		= obj1.get(i).dlvPaonMsg			'배송메세지
'							rw obj1.get(i).webExpsPrmoNm							'웹프로모션 문구
'							rw obj1.get(i).addCmpsItemNm							'추가구성상품명
'							rw obj1.get(i).oshpPrrgNm								'당일출고예정
'							rw obj1.get(i).giftStrtDt								'사은품 시작일 (미사용)
'							rw obj1.get(i).giftEndDt								'사은품 종료일 (미사용)
'							rw obj1.get(i).giftStrtEndDt							'사은품 이벤트 시작/종료일 (미사용)
							matchItemID			= obj1.get(i).venItemCd				'협력사 상품관리코드
							matchItemID = replace(matchItemID, "TEST_", "")
							ItemOrderCount		= obj1.get(i).ordQty				'주문 수량
							matchItemOption		= getTenOptionCode(matchItemID, partnerOptionName)

'							rw "beasongNum11st : " & beasongNum11st
'							rw "reserve01 : " & reserve01
'							rw "OutMallOrderSerial : " & OutMallOrderSerial
'							rw "OrgDetailKey : " & OrgDetailKey
'							rw "outMallGoodsNo : " & outMallGoodsNo
'							rw "partnerItemName : " & partnerItemName
'							rw "outMallOptionNo : " & outMallOptionNo
'							rw "partnerOptionName : " & partnerOptionName
'							rw "lastDlvstPrgrGbcd : " & lastDlvstPrgrGbcd
'							rw "dlvCnclYn : " & dlvCnclYn
'							rw "SellPrice : " & SellPrice
'							rw "ReceiveName : " & ReceiveName
'							rw "ReceiveHpNo : " & ReceiveHpNo
'							rw "ReceiveTelNo : " & ReceiveTelNo
'							rw "OrderName : " & OrderName
'							rw "OrderHpNo : " & OrderHpNo
'							rw "OrderTelNo : " & OrderTelNo
'							rw "ReceiveZipCode : " & ReceiveZipCode
'							rw "ReceiveAddr1 : " & ReceiveAddr1
'							rw "ReceiveAddr2 : " & ReceiveAddr2
'							rw "SellDate : " & SellDate
'							rw "deliverymemo : " & deliverymemo
'							rw "matchItemID : " & matchItemID
'							rw "ItemOrderCount : " & ItemOrderCount
'							rw "matchItemOption : " & matchItemOption

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
					set obj1 = nothing
				End If
			Set strObj = nothing

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
	On Error Goto 0
	Set objXML = nothing
End If
rw searchDate & " Order End"

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
