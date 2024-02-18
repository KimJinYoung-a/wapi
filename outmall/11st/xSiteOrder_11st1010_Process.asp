<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� XML �ֹ�ó��
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

    If InStr(tmpTxtStr, "�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
        tmpTxtStr = Trim(split(tmpTxtStr,"-"&iordCnt&"��")(0))
        tmpTxtStr = Trim(replace(tmpTxtStr,"�ؽ�Ʈ�� �Է��ϼ���:",""))
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
                    ''�ɼ�Ÿ�Ը��� �����;��Ѵ�..
                    foundoptTypename = FALSE
                    if InStr(bufTmpTxtStr(i),":")>0 then
                        if isArray(ArrRows) then
                            for j=0 To UBound(ArrRows,2)
                                if (inStr(bufTmpTxtStr(i),ArrRows(0,j)&":")>0) then
                                    ''�ɼǼ�����.
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

                '' ����
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

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' �����޼���
    else
        ierrCode = -999
        ierrStr = "��ǰ�ڵ� �Ǵ� �ɼ��ڵ�  ��Ī ����" & OrgDetailKey & " ��ǰ�ڵ� =" & matchItemID&" �ɼǸ� = "&partnerOptionName
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
		response.write "��¥ ���� �Դϴ�."
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
					response.write "<br /><input type='button' value='�ڷ�' onclick='history.back(-1);'>"
					response.end
				Else
'rw "�ֹ����� response.end"
'response.end
					Set vOrder = xmlDOM.getElementsByTagName("ns2:order")
						For each SubNodes in vOrder
							orderCsGbn			= 0
							orderDlvPay			= Trim(SubNodes.getElementsByTagName("dlvCst").item(0).text)				'��ۺ�
							beasongNum11st		= Trim(SubNodes.getElementsByTagName("dlvNo").item(0).text)					'��۹�ȣ
							sellsiteUserID		= Trim(SubNodes.getElementsByTagName("memID").item(0).text)					'ȸ��ID
							ordAmt				= Clng(Trim(SubNodes.getElementsByTagName("ordAmt").item(0).text))			'�ֹ��Ѿ� | �ǸŴܰ� * ����(�ֹ� -��� -��ǰ) + �ɼǰ�
							deliverymemo		= Trim(SubNodes.getElementsByTagName("ordDlvReqCont").item(0).text)			'��۽� ��û����
							SellDate			= Trim(SubNodes.getElementsByTagName("ordDt").item(0).text)					'�ֹ��Ͻ�
							OrderName			= LEFT(Trim(SubNodes.getElementsByTagName("ordNm").item(0).text), 28)		'������ �̸�
							OutMallOrderSerial	= Trim(SubNodes.getElementsByTagName("ordNo").item(0).text)					'11���� �ֹ���ȣ
							ordOptWonStl		= Clng(Trim(SubNodes.getElementsByTagName("ordOptWonStl").item(0).text))	'�ֹ���ǰ�ɼǰ����ݾ�
							ordPayAmt			= Clng(Trim(SubNodes.getElementsByTagName("ordPayAmt").item(0).text))		'�����ݾ� | �ֹ��ݾ� + ��ۺ� - �Ǹ��� ���αݾ� - mo����
							OrgDetailKey		= Trim(SubNodes.getElementsByTagName("ordPrdSeq").item(0).text)				'�ֹ�����
							OrderHpNo			= Trim(SubNodes.getElementsByTagName("ordPrtblTel").item(0).text)			'������ �޴�����ȣ
							ItemOrderCount		= Trim(SubNodes.getElementsByTagName("ordQty").item(0).text)				'����
							OrderTelNo			= Trim(SubNodes.getElementsByTagName("ordTlphnNo").item(0).text)			'�ֹ�����ȭ��ȣ
							partnerItemName		= Trim(SubNodes.getElementsByTagName("prdNm").item(0).text)					'��ǰ��
							outMallGoodsNo		= Trim(SubNodes.getElementsByTagName("prdNo").item(0).text)					'11������ǰ��ȣ
							prdStckNo			= Trim(SubNodes.getElementsByTagName("prdStckNo").item(0).text)				'�ֹ���ǰ�ɼ��ڵ�
							ReceiveAddr1		= Trim(SubNodes.getElementsByTagName("rcvrBaseAddr").item(0).text)			'��۱⺻�ּ�
							ReceiveAddr2		= Trim(SubNodes.getElementsByTagName("rcvrDtlsAddr").item(0).text)			'��ۻ��ּ�
							ReceiveZipCode		= Trim(SubNodes.getElementsByTagName("rcvrMailNo").item(0).text)			'����������ȣ
							ReceiveName			= LEFT(Trim(SubNodes.getElementsByTagName("rcvrNm").item(0).text), 28)		'�����ڸ�
							ReceiveHpNo			= Trim(SubNodes.getElementsByTagName("rcvrPrtblNo").item(0).text)			'�������ڵ�����ȣ
							ReceiveTelNo		= Trim(SubNodes.getElementsByTagName("rcvrTlphn").item(0).text)				'��������ȭ��ȣ
							selPrc				= Clng(Trim(SubNodes.getElementsByTagName("selPrc").item(0).text))			'�ǸŰ� | ���ܰ�
							sellerDscPrc		= Clng(Trim(SubNodes.getElementsByTagName("sellerDscPrc").item(0).text))	'�Ǹ��� ���αݾ�
							matchItemID			= Trim(SubNodes.getElementsByTagName("sellerPrdCd").item(0).text)			'�Ǹ��ڻ�ǰ��ȣ
							partnerOptionName	= Trim(SubNodes.getElementsByTagName("slctPrdOptNm").item(0).text)			'�ֹ���ǰ�ɼǸ�
							tmallDscPrc			= Clng(Trim(SubNodes.getElementsByTagName("tmallDscPrc").item(0).text))		'11���� ���αݾ�
							lstTmallDscPrc		= Clng(Trim(SubNodes.getElementsByTagName("lstTmallDscPrc").item(0).text))	'11���� ���αݾ�-����ǰ��
							lstSellerDscPrc 	= Clng(Trim(SubNodes.getElementsByTagName("lstSellerDscPrc").item(0).text))	'�Ǹ��� ���αݾ�-����ǰ��
							sellerStockCd		= Trim(SubNodes.getElementsByTagName("sellerStockCd").item(0).text)			'�Ǹ��� ����ȣ
							SellPrice			= selPrc + (ordOptWonStl / ItemOrderCount)
							RealSellPrice		= SellPrice - Clng((lstTmallDscPrc + lstSellerDscPrc) / ItemOrderCount)
							'RealSellPrice		= SellPrice - Clng(lstSellerDscPrc / ItemOrderCount)



							If sellerStockCd <> "" Then
								matchItemOption = Split(sellerStockCd, "_")(1)
							Else
								matchItemOption = "0000"
							End If

							' If InStr(partnerOptionName, "�ؽ�Ʈ�� �Է��ϼ���") > 0 Then
							' 	requireDetail	= Trim(Split(partnerOptionName, "�ؽ�Ʈ�� �Է��ϼ���:")(1))
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
''							slctPrdOptNm : �ؽ�Ʈ�� �Է��ϼ���:11����_�����׽�Ʈ,�ڵ�������:������5/5S/SE,����:�Ͼ�,�̸�Ƽ��:1.��� ��-1��
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
			    rw "["&failCNT&"] �� ����(�ֹ���ȸ)"
			End if

			If (succCNT <> 0) then
			    rw "["&succCNT&"] �� ����(�ֹ���ȸ)"
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
					rw "["&OKcnt&"] �� ����(����Ȯ��)"
				End If

				If NOcnt <> 0 then
					rw "["&NOcnt&"] �� ����(����Ȯ��)"
				End If
			End If
		Else
			rw "�ֹ����� ����..��� �� �õ� ���"
		End If
	On Error Goto 0
	Set objXML = nothing
End If

''ǰ��/���� ����üũ
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->