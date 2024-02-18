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
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
<!-- #include virtual="/outmall/11st/inc11stFunction.asp"-->
<!-- #include virtual="/outmall/incOutMallCommonFunction.asp"-->
<%
function get11StrequiredetailByOPtionStr(iOrgOptionTxtStr,imatchitemid,imatchoptionCode,iordCnt)
    dim ret : ret=""
    dim i, j, sqlStr
    Dim tmpTxtStr : tmpTxtStr = Trim(iOrgOptionTxtStr)
    Dim bufTmpTxtStr
    Dim ArrRows
    Dim foundoptTypename

    If InStr(tmpTxtStr, "텍스트를 입력하세요") > 0 Then
        tmpTxtStr = Trim(split(tmpTxtStr,"-"&iordCnt&"개")(0))
        tmpTxtStr = Trim(replace(tmpTxtStr,"텍스트를 입력하세요:",""))
        if imatchoptionCode="0000" Then
            ret = tmpTxtStr
        ELSE
            If getChrCount(tmpTxtStr, ",") >= 1 Then
                bufTmpTxtStr = Split(tmpTxtStr, ",")
                sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteOrder_OptionTypeNameList] "&imatchitemid
                rsget.CursorLocation = adUseClient
                rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly
                if NOT rsget.Eof then
                    ArrRows =rsget.getRows
                end if
                rsget.Close

                for i=LBound(bufTmpTxtStr) to UBound(bufTmpTxtStr)
                    ''옵션타입명을 가져와야한다..
                    foundoptTypename = FALSE
                    if InStr(bufTmpTxtStr(i),":")>0 then
                        if isArray(ArrRows) then
                            for j=0 To UBound(ArrRows,2)
                                if (inStr(bufTmpTxtStr(i),ArrRows(0,j)&":")>0) then
                                    ''옵션선택임.
                                    foundoptTypename = true
                                    exit for
                                end if
                            next
                            if (NOT foundoptTypename) then ret = ret & bufTmpTxtStr(i)&","
                        end if
                    else
                        ret = ret & bufTmpTxtStr(i)&","
                    end if
                next
                if Right(ret,1)="," then ret=LEFT(ret,LEN(ret)-1)

                '' 포기
                if (ret="") then ret = tmpTxtStr
            Else
                ret = tmpTxtStr
            End If
        end if
    ELSE
        ret = ""
    end if

    get11StrequiredetailByOPtionStr = ret
end function

Function fn11stConfirmOrder(vOrderserial, vOrgDetailKey, vBeasongNum11st)
	Dim objXML, xmlDOM, iRbody, strSql
	On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APISSLURL&"/ordservices/reqpackaging/" & vOrderserial & "/" & vOrgDetailKey & "/N/null/" & vBeasongNum11st
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				If xmlDOM.getElementsByTagName("result_code").item(0).text = "0" Then
					strSql = ""
					strSql = strSql & " UPDATE db_temp.[dbo].[tbl_xSite_TMP11stOrder] SET "
					strSql = strSql & " isbaljuConfirmSend = 'Y' "
					strSql = strSql & " , lastUpdate = getdate() "
					strSql = strSql & " WHERE outmallorderserial = '"&vOrderserial&"'  "
					strSql = strSql & " and beasongNum11st = '"&vBeasongNum11st&"' "
					strSql = strSql & " and orgDetailKey = '"&vOrgDetailKey&"' "
					strSql = strSql & " and mallid = '11st1010' "
					dbget.Execute strSql
					fn11stConfirmOrder= true
				Else
					fn11stConfirmOrder= false
				End If
			Set xmlDOM = Nothing
		Else
			fn11stConfirmOrder= false
		End If
	Set objXML = nothing
	On Error Goto 0
End Function

function saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr, beasongNum11st)
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
		,Array("@partnerOptionName"	,adVarchar, adParamInput,1024, partnerOptionName) _
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
		,Array("@beasongNum11st"	,adVarchar, adParamInput,16, beasongNum11st) _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert_From11stXML"
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

function getLastOrderInputDT()
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr&" from db_temp.dbo.tbl_XSite_TMpOrder"
    sqlStr = sqlStr&" where sellsite='11st1010'"
    sqlStr = sqlStr&" order by selldate desc"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	if (Not rsget.Eof) then
		getLastOrderInputDT = rsget("lastOrdInputDt")
	end if
	rsget.Close

end function

Dim sqlStr, buf, i, mode, sellsite
Dim divcd, yyyymmdd, idx, Nodes, Nodes2, SubNodes, SubNodes2, vOrder
Dim objXML, xmlDOM, retCode, iMessage, reqOrderdate
mode		= requestCheckVar(html2db(request("mode")),32)
sellsite	= requestCheckVar(html2db(request("sellsite")),32)
idx			= requestCheckVar(html2db(request("idx")),32)
reqOrderdate = request("reqOrderdate")

Dim strsql, retVal, deliverymemo, orderCsGbn, errCode, errStr, succCNT, failCNT
Dim dueDate, iRbody, result_text
Dim orderDlvPay, beasongNum11st, sellsiteUserID, ordAmt, SellDate, OrderName, outMallGoodsNo, ordOptWonStl, ordPayAmt, OrgDetailKey, OrderHpNo, ItemOrderCount, OrderTelNo, partnerItemName
Dim OutMallOrderSerial, prdStckNo, ReceiveAddr1, ReceiveAddr2, ReceiveZipCode, ReceiveName, ReceiveHpNo, ReceiveTelNo, selPrc, sellerDscPrc, matchItemID, partnerOptionName, tmallDscPrc, lstTmallDscPrc, lstSellerDscPrc, sellerStockCd
Dim requireDetail, matchItemOption, SellPrice, RealSellPrice
Dim prev7Day, nowDay, lastOrderDate, resultNode

If reqOrderdate = "" Then
	lastOrderDate = getLastOrderInputDT
Else
	lastOrderDate = reqOrderdate
End If
'lastOrderDate = "2017-11-13"
prev7Day = CStr(Replace((lastOrderDate), "-", ""))&"0000"
nowDay	 = CStr(Replace(Date(), "-", ""))&"2359"

If sellsite = "11st1010" Then

	If (CDate(lastOrderDate) > date()) Then
		response.write "날짜 오류 입니다."
		response.end
	End If

	On Error Resume Next
	dueDate = prev7Day &"/"& nowDay
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "" & APISSLURL&"/ordservices/complete/"&dueDate
		objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
		objXML.setRequestHeader "openapikey",""&APIkey&""
		objXML.send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
				xmlDOM.LoadXML iRbody
				Set resultNode = xmlDOM.getElementsByTagName("ns2:result_code")
					If NOT (resultNode Is Nothing)  Then
						result_text = xmlDOM.getElementsByTagName("ns2:result_text").item(0).text
					End If
				Set resultNode = nothing

				If result_text <> "" Then
					response.write result_text
					response.write "<br /><input type='button' value='뒤로' onclick='history.back(-1);'>"
					response.end
				Else
'rw "주문들어옴 response.end"
'response.end
					Set vOrder = xmlDOM.getElementsByTagName("ns2:order")
						For each SubNodes in vOrder
							orderCsGbn			= 0
							orderDlvPay			= Trim(SubNodes.getElementsByTagName("dlvCst").item(0).text)				'배송비
							beasongNum11st		= Trim(SubNodes.getElementsByTagName("dlvNo").item(0).text)					'배송번호
							sellsiteUserID		= Trim(SubNodes.getElementsByTagName("memID").item(0).text)					'회원ID
							ordAmt				= Clng(Trim(SubNodes.getElementsByTagName("ordAmt").item(0).text))			'주문총액 | 판매단가 * 수량(주문 -취소 -반품) + 옵션가
							deliverymemo		= Trim(SubNodes.getElementsByTagName("ordDlvReqCont").item(0).text)			'배송시 요청사항
							SellDate			= Trim(SubNodes.getElementsByTagName("ordDt").item(0).text)					'주문일시
							OrderName			= LEFT(Trim(SubNodes.getElementsByTagName("ordNm").item(0).text), 28)		'구매자 이름
							OutMallOrderSerial	= Trim(SubNodes.getElementsByTagName("ordNo").item(0).text)					'11번가 주문번호
							ordOptWonStl		= Clng(Trim(SubNodes.getElementsByTagName("ordOptWonStl").item(0).text))	'주문상품옵션결제금액
							ordPayAmt			= Clng(Trim(SubNodes.getElementsByTagName("ordPayAmt").item(0).text))		'결제금액 | 주문금액 + 배송비 - 판매자 할인금액 - mo쿠폰
							OrgDetailKey		= Trim(SubNodes.getElementsByTagName("ordPrdSeq").item(0).text)				'주문순번
							OrderHpNo			= Trim(SubNodes.getElementsByTagName("ordPrtblTel").item(0).text)			'구매자 휴대폰번호
							ItemOrderCount		= Trim(SubNodes.getElementsByTagName("ordQty").item(0).text)				'수량
							OrderTelNo			= Trim(SubNodes.getElementsByTagName("ordTlphnNo").item(0).text)			'주문자전화번호
							partnerItemName		= Trim(SubNodes.getElementsByTagName("prdNm").item(0).text)					'상품명
							outMallGoodsNo		= Trim(SubNodes.getElementsByTagName("prdNo").item(0).text)					'11번가상품번호
							prdStckNo			= Trim(SubNodes.getElementsByTagName("prdStckNo").item(0).text)				'주문상품옵션코드
							ReceiveAddr1		= Trim(SubNodes.getElementsByTagName("rcvrBaseAddr").item(0).text)			'배송기본주소
							ReceiveAddr2		= Trim(SubNodes.getElementsByTagName("rcvrDtlsAddr").item(0).text)			'배송상세주소
							ReceiveZipCode		= Trim(SubNodes.getElementsByTagName("rcvrMailNo").item(0).text)			'배송지우편번호
							ReceiveName			= LEFT(Trim(SubNodes.getElementsByTagName("rcvrNm").item(0).text), 28)		'수령자명
							ReceiveHpNo			= Trim(SubNodes.getElementsByTagName("rcvrPrtblNo").item(0).text)			'수령자핸드폰번호
							ReceiveTelNo		= Trim(SubNodes.getElementsByTagName("rcvrTlphn").item(0).text)				'수령자전화번호
							selPrc				= Clng(Trim(SubNodes.getElementsByTagName("selPrc").item(0).text))			'판매가 | 객단가
							sellerDscPrc		= Clng(Trim(SubNodes.getElementsByTagName("sellerDscPrc").item(0).text))	'판매자 할인금액
							matchItemID			= Trim(SubNodes.getElementsByTagName("sellerPrdCd").item(0).text)			'판매자상품번호
							partnerOptionName	= Trim(SubNodes.getElementsByTagName("slctPrdOptNm").item(0).text)			'주문상품옵션명
							tmallDscPrc			= Clng(Trim(SubNodes.getElementsByTagName("tmallDscPrc").item(0).text))		'11번가 할인금액
							lstTmallDscPrc		= Clng(Trim(SubNodes.getElementsByTagName("lstTmallDscPrc").item(0).text))	'11번가 할인금액-각상품별
							lstSellerDscPrc 	= Clng(Trim(SubNodes.getElementsByTagName("lstSellerDscPrc").item(0).text))	'판매자 할인금액-각상품별
							sellerStockCd		= Trim(SubNodes.getElementsByTagName("sellerStockCd").item(0).text)			'판매자 재고번호
							SellPrice			= selPrc + (ordOptWonStl / ItemOrderCount)
							RealSellPrice		= SellPrice - Clng((lstTmallDscPrc + lstSellerDscPrc) / ItemOrderCount)
							'RealSellPrice		= SellPrice - Clng(lstSellerDscPrc / ItemOrderCount)



							If sellerStockCd <> "" Then
								matchItemOption = Split(sellerStockCd, "_")(1)
							Else
								matchItemOption = "0000"
							End If

							' If InStr(partnerOptionName, "텍스트를 입력하세요") > 0 Then
							' 	requireDetail	= Trim(Split(partnerOptionName, "텍스트를 입력하세요:")(1))
							' 	If getChrCount(requireDetail, ",") >= 1 Then
							' 		requireDetail = Trim(Split(requireDetail, ",")(0))
							' 	Else
							' 		requireDetail = Trim(Split(requireDetail, "-")(0))
							' 	End If
							' Else
							' 	requireDetail	= ""
							' End If

							requireDetail = get11StrequiredetailByOPtionStr(partnerOptionName,matchItemID,matchItemOption,ItemOrderCount)

'							rw "orderDlvPay : " & orderDlvPay
'							rw "beasongNum11st : " & beasongNum11st
'							rw "sellsiteUserID : " & sellsiteUserID
'							rw "ordAmt : " & ordAmt
'							rw "deliverymemo : " & deliverymemo
'							rw "SellDate : " & SellDate
'							rw "OrderName : " & OrderName
'							rw "OutMallOrderSerial : " & OutMallOrderSerial
'							rw "ordOptWonStl : " & ordOptWonStl
'							rw "ordPayAmt : " & ordPayAmt
'							rw "OrgDetailKey : " & OrgDetailKey
'							rw "OrderHpNo : " & OrderHpNo
'							rw "ItemOrderCount : " & ItemOrderCount
'							rw "OrderTelNo : " & OrderTelNo
'							rw "partnerItemName : " & partnerItemName
'							rw "outMallGoodsNo : " & outMallGoodsNo
'							rw "prdStckNo : " & prdStckNo
'							rw "ReceiveAddr1 : " & ReceiveAddr1
'							rw "ReceiveAddr2 : " & ReceiveAddr2
'							rw "ReceiveZipCode : " & ReceiveZipCode
'							rw "ReceiveName : " & ReceiveName
'							rw "ReceiveHpNo : " & ReceiveHpNo
'							rw "ReceiveTelNo : " & ReceiveTelNo
'							rw "selPrc : " & selPrc
'							rw "sellerDscPrc : " & sellerDscPrc
'							rw "matchItemID : " & matchItemID
'							rw "partnerOptionName : " & partnerOptionName
'							rw "tmallDscPrc : " & tmallDscPrc
'							rw "lstTmallDscPrc : " & lstTmallDscPrc
'							rw "lstSellerDscPrc : " & lstSellerDscPrc
'							rw "sellerStockCd : " & sellerStockCd
'							rw "requireDetail : " & requireDetail
'							rw "matchItemOption : " & matchItemOption
'							rw "SellPrice : " & SellPrice
'							rw "RealSellPrice : " & RealSellPrice
''							slctPrdOptNm : 텍스트를 입력하세요:11번가_문구테스트,핸드폰기종:아이폰5/5S/SE,색상:하얀,이모티콘:1.노란 달-1개
'							rw "--------------------------------------------------------"

							retVal= saveOrderOneToTmpTable(SellSite, OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName,outMallGoodsNo _
									, OrderName, OrderTelNo, OrderHpNo _
									, ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
									, SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
									, deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
									, errCode, errStr, beasongNum11st )
							If (retVal) Then
								succCNT = succCNT + 1
								strsql = ""
								strsql = strsql & " INSERT INTO db_temp.[dbo].[tbl_xSite_TMP11stOrder] (outmallorderserial, OrgDetailKey, beasongNum11st, isbaljuConfirmSend, regdate, mallid) "
								strsql = strsql & " VALUES ('"&OutMallOrderSerial&"', '"&OrgDetailKey&"', '"&beasongNum11st&"', 'N', getdate(), '11st1010')"
								dbget.Execute strSql
							Else
								failCNT = failCNT + 1
							End If
						Next
					Set vOrder = nothing
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
				strsql = strsql & " and T.mallid = '11st1010' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " update T "
				strsql = strsql & " set T.isbaljuConfirmSend='Y' "
				strsql = strsql & " FROM db_order.dbo.tbl_order_master as M "
				strsql = strsql & " JOIN db_temp.[dbo].[tbl_xSite_TMP11stOrder] as T on M.authcode = T.outmallorderserial "
				strsql = strsql & " WHERE M.cancelyn ='Y' "
				strsql = strsql & " and T.isbaljuConfirmSend <> 'Y' "
				strsql = strsql & " and T.mallid = '11st1010' "
				dbget.Execute strsql

				strsql = ""
				strsql = strsql & " SELECT TOP 3000 outmallorderserial, OrgDetailKey, beasongNum11st FROM db_temp.[dbo].[tbl_xSite_TMP11stOrder] "
				strsql = strsql & " WHERE isbaljuConfirmSend = 'N' "
				strsql = strsql & " and mallid = '11st1010' "
				strsql = strsql & " and regdate > '2021-11-01' "
				strsql = strsql & " ORDER BY regdate DESC "
				rsget.CursorLocation = adUseClient
				rsget.Open strsql, dbget, adOpenForwardOnly, adLockReadOnly
			    if not rsget.Eof then
			        arrList = rsget.getRows()
			    end if
			    rsget.close

				For lp = 0 To Ubound(arrList, 2)
					if (NOT (application("Svr_Info")="Dev")) then
						ret1 = fn11stConfirmOrder(arrList(0, lp), arrList(1, lp), arrList(2, lp))
					end if

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
		Else
			rw "주문연동 실패..잠시 후 시도 요망"
		End If
	On Error Goto 0
	Set objXML = nothing
End If

''품절/가격 오류체크
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->