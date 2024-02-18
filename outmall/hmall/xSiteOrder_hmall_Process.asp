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
	'ProcGb | P1:�ֹ�Ȯ��, P2:���Ϸ�, P3:��ۿϷ�
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

	If ipartnerOptionName = "���Ͽɼ�" Then
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
		'prgrGb | P0:�����, P1:�������, P2:���, P3:��ۿϷ�
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
'							rw obj1.get(i).editYn									'�������ɿ���(�ù��,�����) | "Y:���»�����(40) AND �������(30) AND ������ȣnull AND �����ҿ���N"
							beasongNum11st		= obj1.get(i).dlvstNo				'������ù�ȣ
							reserve01			= obj1.get(i).dlvstPtcSeq			'������û󼼹�ȣ
							OutMallOrderSerial	= obj1.get(i).ordNo					'�ֹ���ȣ
							OrgDetailKey		= obj1.get(i).ordPtcSeq				'�ֹ��Ϸù�ȣ
'							rw obj1.get(i).sitmCd									'�����ǰ�ڵ�
							outMallGoodsNo		= obj1.get(i).slitmCd				'�ǸŻ�ǰ�ڵ�
							partnerItemName		= obj1.get(i).slitmNm				'��ǰ��
							outMallOptionNo		= obj1.get(i).uitmCd				'��ǰ�Ӽ��ڵ�
							partnerOptionName	= obj1.get(i).uitmTotNm				'��ǰ�Ӽ���
'							rw obj1.get(i).dlvFormGbcd								'������� | 20:����Ȩ���ù�, 30:���»����ù�, 40:���»�����
							lastDlvstPrgrGbcd	= obj1.get(i).lastDlvstPrgrGbcd		'��������������౸���ڵ� | 25:�����, 30:�������, 45:���, 50:��ۿϷ�
'							rw obj1.get(i).lastOshpDlineDt							'�������������
'							rw obj1.get(i).oshpReqnDt								'����û��
							dlvCnclYn			= obj1.get(i).dlvCnclYn				'�����ҿ���
'							rw obj1.get(i).dlvCnclNm								'������
'							rw obj1.get(i).dlvstQty									'������ü���
'							rw obj1.get(i).unitDlvQty								'������
'							rw obj1.get(i).prrgQty									'������
'							rw obj1.get(i).custCnclQty								'����Ҽ���
							SellPrice			= obj1.get(i).sellUprc				'�ǸŴܰ�
							RealSellPrice		= obj1.get(i).sellUprc				'����Comment : �� �ǸŰ��� �� �Ѿ��...�׳� �ǸŰ��� ����
'							rw obj1.get(i).sellSum									'�ǸŰ� (���� �ڸ�Ʈ : �Ǹ��հ谡 ���� ��)
'							rw obj1.get(i).prchUprcSum								'���԰�
'							rw obj1.get(i).dsrvDlvcoCd								'�ù��ۻ��ڵ�
'							rw obj1.get(i).invcNo									'������ȣ
'							rw obj1.get(i).dsntDtDlvYn								'�������ڹ�ۿ���
'							rw obj1.get(i).rsvSellYn								'�����Ǹſ���
'							rw obj1.get(i).venCd									'���»�
'							rw obj1.get(i).ven2Cd									'2�����»�
							dlvTypeGbcd = obj1.get(i).dlvTypeGbcd					'�������(�ֹ�����) | 10:�ֹ����; 40:��ȯ���
'							rw obj1.get(i).dlvTypeGbcdColor							'�������ǥ�û���
'							rw obj1.get(i).dlvcPayGbcd								'��ۺ����ұ����ڵ� | 00:����, 10:������, 20:����, 30:��ġ��ǰ
'							rw obj1.get(i).befDlvstNo								'����������ù�ȣ
'							rw obj1.get(i).befDlvstPtcSeq							'����������û󼼼���
'							rw obj1.get(i).ordSplpnAplyYn							'�ֹ����ް�ȹ���뿩��
'							rw obj1.get(i).custDlvHopeDt							'������������
'							rw obj1.get(i).oshpCnfmDtm								'���Ȯ���Ͻ�
							ReceiveName			= Left(obj1.get(i).rcvrNm, 28)		'�μ��ڸ� | �����ڸ�
							ReceiveHpNo			= obj1.get(i).rcvrTel				'�μ�����ȭ | ��������ȭ��ȣ(Astrisk)
							ReceiveTelNo		= obj1.get(i).rcvrTel				'�μ�����ȭ | ��������ȭ��ȣ(Astrisk)
							OrderName			= Left(obj1.get(i).dlvApltNm, 28)	'�ֹ��� | ��۽�û�ڸ�(Astrisk)
							OrderHpNo			= obj1.get(i).dlvApltTel			'�ֹ�����ȭ | ��۽�û����ȭ��ȣ(Astrisk)
							OrderTelNo			= obj1.get(i).dlvApltTel			'�ֹ�����ȭ | ��۽�û����ȭ��ȣ(Astrisk)
							ReceiveZipCode		= obj1.get(i).dstnPostNo			'����������ȣ | (astrisk)
							ReceiveAddr			= obj1.get(i).dstnAdr				'����� | ������ּ�(astrisk)

							'''�ּҿ� ���ּҰ� ������� 3��° Blank���� ����.
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

'							rw obj1.get(i).ordCustNo								'�ֹ�����ȣ
'							rw obj1.get(i).custDstnSeq								'�����������
'							rw obj1.get(i).rcvrPaonMsg								'���ϴ¸� | ���������޸޽���
'							rw obj1.get(i).befInvcNo								'�������ȣ
'							rw obj1.get(i).prsnYn									'�������忩�� | ��������
'							rw obj1.get(i).frdlvYn									'�ؿܹ�ۿ���
'							rw obj1.get(i).frgnDstnSeq								'�ؿܹ��������
'							rw obj1.get(i).custVenPaonMsg							'�ֹ��ÿ�û���� | �����»����޸޽���
'							rw obj1.get(i).frgnAdr									'�ؿܹ����
'							rw obj1.get(i).hmallGiftRfrNote							'hmall����ǰ��������
'							rw obj1.get(i).dlvNo									'��۹�ȣ
'							rw obj1.get(i).dlvPtcSeq								'��ۻ󼼼���
							SellDate = LEFT(obj1.get(i).ptcOrdDtm, 10)				'���ֹ�����
'							rw obj1.get(i).cvstDlvYn								'��������ۿ��� | ��������ۿ���( Y or N)
'							rw obj1.get(i).inslItemYn								'��ġ��ǰ���� | INSL_ITEM_YN
'							rw obj1.get(i).almlBasktNo								'������ٱ��Ϲ�ȣ
'							rw obj1.get(i).nshipTypeGbcd							'��������������ڵ� | �������:10, ǰ�����:20
							deliverymemo		= obj1.get(i).dlvPaonMsg			'��۸޼���
'							rw obj1.get(i).webExpsPrmoNm							'�����θ�� ����
'							rw obj1.get(i).addCmpsItemNm							'�߰�������ǰ��
'							rw obj1.get(i).oshpPrrgNm								'���������
'							rw obj1.get(i).giftStrtDt								'����ǰ ������ (�̻��)
'							rw obj1.get(i).giftEndDt								'����ǰ ������ (�̻��)
'							rw obj1.get(i).giftStrtEndDt							'����ǰ �̺�Ʈ ����/������ (�̻��)
							matchItemID			= obj1.get(i).venItemCd				'���»� ��ǰ�����ڵ�
							matchItemID = replace(matchItemID, "TEST_", "")
							ItemOrderCount		= obj1.get(i).ordQty				'�ֹ� ����
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

							If (dlvCnclYn <> "Y") AND (dlvTypeGbcd <> "40") Then	'�����ҿ��ΰ� Y�� �ƴϰ�, ��ȯ�ֹ��� �ƴϸ� ����
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
					rw "["&OKcnt&"] �� ����(����Ȯ��)"
				End If

				If NOcnt <> 0 then
					rw "["&NOcnt&"] �� ����(����Ȯ��)"
				End If
			End If
'			response.end
			If (iSelldate < Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, Left(DateAdd("d", 1, CDate(iSellDate)), 10), "N")
			ElseIf (iSellDate = Left(Now(), 10)) then
				Call SetCheckStatus(sellsite, iSellDate, "Y")
			End If
		Else
			rw "�ֹ����� ����..��� �� �õ� ���"
		End If
	On Error Goto 0
	Set objXML = nothing
End If
rw searchDate & " Order End"

''ǰ��/���� ����üũ
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
