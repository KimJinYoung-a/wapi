<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="javascript" runat="server">
var confirmDt = (new Date()).valueOf();
</script>
<style>
body {
  font-size: small;
}
</style>
</head>
<body bgcolor="#F4F4F4" >
<%
function getEzWelDlvCode2Name(idlvCd)
    if isNULL(idlvCd) then Exit function

    SELECT CASE idlvCd
        CASE "1007" : getEzWelDlvCode2Name = "CJ�������"
        CASE "1017" : getEzWelDlvCode2Name = "�Ե��ù�"
        CASE "1016" : getEzWelDlvCode2Name = "�����ù�"
        CASE "1008" : getEzWelDlvCode2Name = "�����ù�"
        CASE "1161" : getEzWelDlvCode2Name = "������"

        CASE "1180" : getEzWelDlvCode2Name = "�Ͼ������"
        CASE "1163" : getEzWelDlvCode2Name = "�̳�����"
        CASE "1200" : getEzWelDlvCode2Name = "����ù�"
        CASE "1082" : getEzWelDlvCode2Name = "��Ÿ�ù�"
        CASE "1001" : getEzWelDlvCode2Name = "DHL"
        CASE "1002" : getEzWelDlvCode2Name = "KGB�ù�"
        CASE "1005" : getEzWelDlvCode2Name = "�浿�ù�"
        CASE "1011" : getEzWelDlvCode2Name = "���ο�ĸ"
        CASE "1012" : getEzWelDlvCode2Name = "��ü���ù�EMS"
        CASE "1014" : getEzWelDlvCode2Name = "õ���ù�"
        CASE "1080" : getEzWelDlvCode2Name = "KG�������ù�"
        CASE "1081" : getEzWelDlvCode2Name = "��ü�����"
        CASE "1260" : getEzWelDlvCode2Name = "GTX������"
        
        CASE "1102" : getEzWelDlvCode2Name = "�յ��ù�"
        CASE "1103" : getEzWelDlvCode2Name = "���ǻ���ù�"
        CASE "1104" : getEzWelDlvCode2Name = "�ٵ帲"
        CASE "1105" : getEzWelDlvCode2Name = "������"
        CASE "1106" : getEzWelDlvCode2Name = "�ǿ��ù�"
        CASE "1107" : getEzWelDlvCode2Name = "ȣ���ù�"
        CASE "1108" : getEzWelDlvCode2Name = "CJ��������Ư��"
        CASE "1109" : getEzWelDlvCode2Name = "EMS"
        CASE "1110" : getEzWelDlvCode2Name = "�ѵ���"
        CASE "1111" : getEzWelDlvCode2Name = "FedEx"
        CASE "1112" : getEzWelDlvCode2Name = "UPS"
        CASE "1113" : getEzWelDlvCode2Name = "TNT"
        CASE "1114" : getEzWelDlvCode2Name = "USPS"
        CASE "1115" : getEzWelDlvCode2Name = "i-parcel"
        CASE "1116" : getEzWelDlvCode2Name = "GSM NtoN"
        CASE "1117" : getEzWelDlvCode2Name = "�����۷ι�"
        CASE "1118" : getEzWelDlvCode2Name = "�������佺"
        CASE "1119" : getEzWelDlvCode2Name = "ACI Express"
        CASE "1121" : getEzWelDlvCode2Name = "���۷ι�"
        CASE "1122" : getEzWelDlvCode2Name = "������ͽ�������"
        CASE "1123" : getEzWelDlvCode2Name = "KGL��Ʈ����"
        CASE "1124" : getEzWelDlvCode2Name = "LineExpress"
        CASE "1125" : getEzWelDlvCode2Name = "2fast�ͽ�������"
        CASE "1126" : getEzWelDlvCode2Name = "GSI�ͽ�������"
        CASE "1240" : getEzWelDlvCode2Name = "�������ù�"
        CASE ELSE : getEzWelDlvCode2Name =""
    END SELECT 
end function

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''���ֵȳ����� (�ֹ����� view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''�ֹ�Ȯ�� Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''�Է� Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE 

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-2,now()),10)

    if eddt<>"" then 
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

 

CONST cspCd		= "10040413"							'CP��ü�ڵ�(������ �߱�)
CONST crtCd		= "8e5a6dbdd27efb49fc600c293884ef47"	'�����ڵ�(������ �߱�)

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ʱ�ȭ�۾�"

'' 1001:�ֹ��Ϸ� / 1002:����غ��� / 1003:����� / 1004:����Ϸ� / 1005:�ֹ���� / 1007:��ǰ��û ....
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)

call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1001","�ֹ���Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1002","�ֹ�Ȯ��")
response.flush
call Get_ezwelOrderListByStatus(istyyyymmdd,iedyyyymmdd,"1003","�����")
response.flush

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ֹ�����"

rw "�Ϸ�"

' �Ϻ������ �̿��� �� ��������.
' call Get_ezwelOrderListByStatus("2019-10-25","2019-10-25","1004","����Ϸ�")
' response.flush
'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

function Get_ezwelOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
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

    dim strSql, bufStr

	Get_ezwelOrderListByStatus = False

	'// =======================================================================
	'// ��¥����
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(�Ⱓ������ �ֹ� ��������)
	xmlURL = "http://api.ezwel.com/if/api/orderListAPI.ez"
	''response.write xmlURL

	postParam = "cspCd=" & cspCd & "&crtCd=" & crtCd
	postParam = postParam & "&startDate=" & Replace(stdate, "-", "") & "000000"
	postParam = postParam & "&endDate=" & Replace(Left(DateAdd("d", 1, CDate(eddate)), 10), "-", "") & "000000"
	postParam = postParam & "&orderStatus="&iorderStatus
	''response.write postParam

    rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" ����:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// ����Ÿ ��������
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "POST", xmlURL
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send(postParam)

	if objXML.Status <> "200" then
		if IsAutoScript then
			response.write "ERROR : ��ſ���"
		else
			response.write "ERROR : ��ſ���" & objXML.Status
			response.write "<script>alert('ERROR : ��ſ���.');</script>"
		end if

		dbget.close : response.end
	end if

    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim  optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim  orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo

	'// =======================================================================
	'// XML DOM ����
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False

    if (application("Svr_Info")="Dev") then

        bufStr = "<?xml version='1.0' encoding='EUC-KR'?>"
        bufStr = bufStr & "<resultSet class='array'><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>8640</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1234873</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016905554</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]���ŰƼ �ٿ뵵��ߴ� ����]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>12100</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>6</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>9600</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053458145</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191018181827</orderDt><orderNum type='string'>1026227679</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��⵵ �Ȼ�� �ܿ��� ���ݷ� 164 (���ܵ�, ���ܷԵ�ĳ�������ũ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[108�� 2102ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-9648-0820</rcvrMobile><rcvrNm type='string'>���?</rcvrNm><rcvrPost type='string'>15347</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-9648-0820</sndMobile><sndNm type='string'>�����</sndNm><sndTelNum type='string'>02-</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>177300</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2376775</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019145570</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]���ڰ��� ���� ���� 800 ���Űſ� ������A ���� NA8744]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>323000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[����:�׷���^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>197000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053491658</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020211957</orderDt><orderNum type='string'>1026244168</orderNum><orderReqContent type='string'><![CDATA[����� �̸� ���� �ٶ��ϴ�. ]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��⵵ �Ȼ�� �ܿ��� ���2�� 29 (���ܵ�, �ְ�8��������Ʈ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[801�� 606ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-7270-9109</rcvrMobile><rcvrNm type='string'>Ȳ����</rcvrNm><rcvrPost type='string'>15338</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-7270-9109</sndMobile><sndNm type='string'>Ȳ����</sndNm><sndTelNum type='string'>02-</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>60300</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1784132</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd class='object' null='true'/><dlvrDt class='object' null='true'/><dlvrFinishDt class='object' null='true'/><dlvrNo class='object' null='true'/><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1015413316</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]��ũ��Ÿ �ڼ� �ϸ�Ŀư]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>67000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>2</orderQty><orderStatus type='string'>1002</orderStatus><salePrice type='number'>67000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053494255</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020230642</orderDt><orderNum type='string'>1026245580</orderNum><orderReqContent type='string'><![CDATA[�����չ�����ּ��� ]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[04355]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[����Ư���� ��걸 ����� 250 (ȿâ��, ����û�������Ʈ) 101�� 803ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-3023-5688</rcvrMobile><rcvrNm type='string'>�����</rcvrNm><rcvrPost type='string'>04355</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-3023-5688</sndMobile><sndNm type='string'>�����</sndNm><sndTelNum type='string'>-</sndTelNum></arrOrderList><orderCnt type='number'>3</orderCnt><resultCode type='string'>200</resultCode><resultMsg type='string'>����</resultMsg></resultSet>"
        
        bufStr = "<?xml version='1.0' encoding='EUC-KR'?>"
        bufStr = bufStr & "<resultSet class='array'><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>128160</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1738833</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191010161501</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625820114020</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1015247678</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]��ȭ�� �� Ȭ�� �̺� �� (Q)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[����:������Ʈ,�ػ����߰�����^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>142400</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053219404</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191007161131</orderDt><orderNum type='string'>1026118893</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[����Ư���� ������  �����19�� 15 (�Ͽ���, ���Ÿ�����Ʈ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[109�� 1403ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-8459-3620</rcvrMobile><rcvrNm type='string'>������</rcvrNm><rcvrPost type='string'>06355</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-8459-3620</sndMobile><sndNm type='string'>������</sndNm><sndTelNum type='string'>02-459-3620</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>7740</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2468588</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191011154502</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625902859056</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019787277</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����][���ý�������] ����� / ���۸������� / 230ml]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[:^]]></optionContent><orderGoodsNum type='number'>2</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>8600</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053227540</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191007215851</orderDt><orderNum type='string'>1026122245</orderNum><orderReqContent type='string'><![CDATA[�����ù��Կ� �־��ּ���.]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��õ������ ������ �۵����з�27���� 55 (�۵���, �Ե�ĳ�� ķ�۽�Ÿ��)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[101-1504]]></rcvrAddr2><rcvrMobile type='string'>010-5573-7371</rcvrMobile><rcvrNm type='string'>������</rcvrNm><rcvrPost type='string'>21982</rcvrPost><rcvrTelNum type='string'>02-0000-0000</rcvrTelNum><sndMobile type='string'>010-5573-7371</sndMobile><sndNm type='string'>������</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>41310</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2400003</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191010141504</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625868717385</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1019261135</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]��� ���� ���ڷ�������(��,ȭ��Ʈ)_(2172003)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[����:ȭ��Ʈ^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>45900</salePrice><useAmt type='string'>2300</useAmt></arrOrderGoods><aspOrderNum type='string'>1053264243</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191009175555</orderDt><orderNum type='string'>1026137775</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[������ ��õ�� �ĸ���24���� 13 (����, ����5������Ʈ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[501�� 414ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-9086-9004</rcvrMobile><rcvrNm type='string'>�輮��</rcvrNm><rcvrPost type='string'>24285</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-9086-9004</sndMobile><sndNm type='string'>�輮��</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>18000</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1392969</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191011104503</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625871366653</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016258682</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]�긮��Ŭ(Brinicle) Ǯ �����η��� 304 ���ô�����/ ���úм��]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>0</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>20000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053267154</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191009203155</orderDt><orderNum type='string'>1026139025</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��õ������ ���� ������840���� 21 (���ϵ�, ǳ�����̿�����Ʈ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[824��1403ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-2602-4682</rcvrMobile><rcvrNm type='string'>���</rcvrNm><rcvrPost type='string'>22682</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-2602-4682</sndMobile><sndNm type='string'>���</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>29700</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2277613</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1008</dlvrCd><dlvrDt type='string'>20191016124501</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>95184514994</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1018868242</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]���� �����ġ��]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>33000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[����:��ũ^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>33000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053289051</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191010181801</orderDt><orderNum type='string'>1026148586</orderNum><orderReqContent type='string'><![CDATA[]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��⵵ ȭ���� ���Ƿ� 99 (��ȵ�, ǳ���Ź��־���Ʈ)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[114-402]]></rcvrAddr2><rcvrMobile type='string'>010-8917-5627</rcvrMobile><rcvrNm type='string'>�ּ���</rcvrNm><rcvrPost type='string'>18342</rcvrPost><rcvrTelNum type='string'>02-8917-5627</rcvrTelNum><sndMobile type='string'>010-8917-5627</sndMobile><sndNm type='string'>�ּ���</sndNm><sndTelNum type='string'>010-8917-5627</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>40500</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>2189217</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1007</dlvrCd><dlvrDt type='string'>20191018202205</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>625927013676</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>2500</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1018127841</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]�����̸� UV���� �ϸ� ����Ŀư 95X210]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>60000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>45000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053346247</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191014000135</orderDt><orderNum type='string'>1026174793</orderNum><orderReqContent type='string'><![CDATA[100cm (����) x 150cm (����) /���Ϸ�]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[����Ư���� ���۱� ����16�� 131-5 (��絿)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[301ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-9948-8313</rcvrMobile><rcvrNm type='string'>�迵��</rcvrNm><rcvrPost type='string'>07022</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-9948-8313</sndMobile><sndNm type='string'>�迵��</sndNm><sndTelNum type='string'>010-9948-8313</sndTelNum></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>16200</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1649904</cspGoodsCd><dccpnPrice class='object' null='true'/><dlvrCd type='string'>1008</dlvrCd><dlvrDt type='string'>20191015164603</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>95202908673</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1016673436</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����]�ָ��� ��� �ǳ�ȭ - ������_(1860391)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>18000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[]]></optionContent><orderGoodsNum type='number'>3</orderGoodsNum><orderQty type='number'>2</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>18000</salePrice><useAmt type='string'>0</useAmt></arrOrderGoods><aspOrderNum type='string'>1053369946</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191014235220</orderDt><orderNum type='string'>1026185859</orderNum><orderReqContent type='string'><![CDATA[����� ���տ� �ΰ��ּ���~]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[��⵵ ���ֽ� �ʿ��� ������ 1023-10]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[���յ� 303ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-8915-9103</rcvrMobile><rcvrNm type='string'>������</rcvrNm><rcvrPost type='string'>12736</rcvrPost><rcvrTelNum class='object' null='true'/><sndMobile type='string'>010-8915-9103</sndMobile><sndNm type='string'>������</sndNm><sndTelNum class='object' null='true'/></arrOrderList><arrOrderList class='object'><arrOrderGoods class='object'><buyPrice type='number'>24120</buyPrice><cancelDt class='object' null='true'/><cspCalUseAmt type='string'>0</cspCalUseAmt><cspCd type='string'>10040413</cspCd><cspDlvrId type='string'>10040413</cspDlvrId><cspGoodsCd type='string'>1997414</cspGoodsCd><dccpnPrice type='string'>0</dccpnPrice><dlvrCd type='string'>1014</dlvrCd><dlvrDt type='string'>20191021144504</dlvrDt><dlvrFinishDt class='object' null='true'/><dlvrNo type='string'>51969716790</dlvrNo><dlvrPayCd type='string'>1001</dlvrPayCd><dlvrPrice type='number'>0</dlvrPrice><exchangeKey class='object' null='true'/><goodsCd type='string'>1017342754</goodsCd><goodsNm type='string'><![CDATA[[�ٹ�����][�̳뼾Ʈ] ���� 300 ���� ��ġ �ſ�_(1237091)]]></goodsNm><marginRate type='number'>10.0</marginRate><normalSalePrice type='number'>29000</normalSalePrice><optionAddPrice type='number'>0</optionAddPrice><optionContent type='string'><![CDATA[����:ȭ��Ʈ^]]></optionContent><orderGoodsNum type='number'>1</orderGoodsNum><orderQty type='number'>1</orderQty><orderStatus type='string'>1003</orderStatus><salePrice type='number'>26800</salePrice><useAmt type='string'>1000</useAmt></arrOrderGoods><aspOrderNum type='string'>1053481532</aspOrderNum><dlvrHopeDt class='object' null='true'/><orderDt type='string'>20191020111041</orderDt><orderNum type='string'>1026239132</orderNum><orderReqContent type='string'><![CDATA[���տ� ���ּ���]]></orderReqContent><rcvrAddr1 type='string'><![CDATA[���󳲵� ��õ�� ǳ������ 22 (ǳ����)]]></rcvrAddr1><rcvrAddr2 type='string'><![CDATA[������ 201ȣ]]></rcvrAddr2><rcvrMobile type='string'>010-6669-6416</rcvrMobile><rcvrNm type='string'>�����</rcvrNm><rcvrPost type='string'>57995</rcvrPost><rcvrTelNum type='string'>02--</rcvrTelNum><sndMobile type='string'>010-6669-6416</sndMobile><sndNm type='string'>�����</sndNm><sndTelNum type='string'>-</sndTelNum></arrOrderList><orderCnt type='number'>8</orderCnt><resultCode type='string'>200</resultCode><resultMsg type='string'>����</resultMsg></resultSet>"


        xmlDOM.LoadXML(bufStr)

    else
	    xmlDOM.LoadXML(objXML.responseText)
    end if
'response.write "<textarea cols='40' rows=10'>"&objXML.responseText & "</textarea><br /><br />"


	if (xmlDOM.getElementsByTagName("resultSet/arrOrderList").length < 1) then
		''if IsAutoScript then
			response.write "�������� : ����" & "<br />"
		''end if

        if (xmlDOM.getElementsByTagName("resultSet/resultMsg").length>0) then
            rw "resultMsg:"&xmlDOM.getElementsByTagName("resultSet/resultMsg")(0).Text
        end if

		Get_ezwelOrderListByStatus = True
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	else
		response.write "�Ǽ�(" & xmlDOM.getElementsByTagName("resultSet/arrOrderList").length & ") " & "<br />"

        ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'ezwel','"&confirmDt&"'"
        ' dbget.Execute strSql
	end if

	set objMasterListXML = xmlDOM.getElementsByTagName("resultSet/arrOrderList")
	For each objMasterOneXML in objMasterListXML
        
        
		set objDetailListXML = objMasterOneXML.getElementsByTagName("arrOrderGoods")
		For each objDetailOneXML in objDetailListXML
            ordNo = objMasterOneXML.getElementsByTagName("orderNum")(0).Text      ''�ֹ���ȣ
            ordItemSeq = objDetailOneXML.getElementsByTagName("orderGoodsNum")(0).Text  ''�ֹ�����
            shppNo = "" ''��۹�ȣ
            shppSeq ="" ''���Seq
            reOrderYn ="N" ''���ֹ����� 
            delayNts  =""  ''�����ϼ�
            cspGoodsCd = objDetailOneXML.getElementsByTagName("cspGoodsCd")(0).Text   ''��ü��ǰ�ڵ�
            goodsCd = objDetailOneXML.getElementsByTagName("goodsCd")(0).Text         ''(����)��ǰ�ڵ�
            uitemId = "" ''���� ��ǰID
            orderQty = objDetailOneXML.getElementsByTagName("orderQty")(0).Text       ''�ֹ�����
            shppDivDtlNm = "" ''��۱��л󼼸� (���/��ȯ���..)
            optionContent = objDetailOneXML.getElementsByTagName("optionContent")(0).Text        ''�ɼǸ� */arrOrderGoods/optionContent
            shppRsvtDt = ""  ''������?
            whoutCritnDt ="" ''��������
            autoShortgYn ="" ''�ڵ���ǰ����

            orderStatus = objDetailOneXML.getElementsByTagName("orderStatus")(0).Text         ''�ֹ�����
            dlvrCd = objDetailOneXML.getElementsByTagName("dlvrCd")(0).Text         ''�ù��
            dlvrNo = objDetailOneXML.getElementsByTagName("dlvrNo")(0).Text         ''�����ȣ
            dlvrDt = objDetailOneXML.getElementsByTagName("dlvrDt")(0).Text         ''�����
            dlvrFinishDt = objDetailOneXML.getElementsByTagName("dlvrFinishDt")(0).Text   ''��ۿϷ���
            cancelDt = objDetailOneXML.getElementsByTagName("cancelDt")(0).Text       ''�����

            bufStr = ""
            bufStr = sellsite&"|"&ordNo
            bufStr = bufStr &"|"&ordItemSeq
            bufStr = bufStr &"|"&cspGoodsCd
            bufStr = bufStr &"|"&goodsCd
            
            bufStr = bufStr &"|"&orderQty

            bufStr = bufStr &"|"&orderStatus
            bufStr = bufStr &"|"&dlvrCd
			bufStr = bufStr &"|"&dlvrNo

            bufStr = bufStr &"|"&dlvrDt
            bufStr = bufStr &"|"&dlvrFinishDt
            bufStr = bufStr &"|"&cancelDt

            shppTypeDtlNm = ""
            delicoVenId   = dlvrCd
            delicoVenNm   = getEzWelDlvCode2Name(dlvrCd)
            wblNo         = dlvrNo

            sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
            paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, sellsite)	_
                ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordItemSeq)) _

                ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim(shppNo)) _
                ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim(shppSeq)) _
                ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(reOrderYn)) _
                ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(delayNts)) _
                ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(cspGoodsCd)) _
                ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(goodsCd)) _
                ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(uitemId)) _
                ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(orderQty)) _
                ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(optionContent)) _
                ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim(shppRsvtDt)) _
                ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim(whoutCritnDt)) _
                ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim(autoShortgYn)) _
                ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(orderStatus)) _

                ,Array("@shppTypeDtlNm"		, adVarchar		, adParamInput		,   16, Trim(shppTypeDtlNm)) _
                ,Array("@delicoVenId"		, adVarchar		, adParamInput		,   16, Trim(delicoVenId)) _
                ,Array("@delicoVenNm"		, adVarchar		, adParamInput		,   32, Trim(delicoVenNm)) _
                ,Array("@wblNo"		        , adVarchar		, adParamInput		,   32, Trim(wblNo)) _

                ,Array("@invoiceUpDt"	    , adVarchar		, adParamInput		,   19, "") _
                ,Array("@outjFixedDt"		, adVarchar		, adParamInput		,   19, Trim(dlvrFinishDt)) _
            )

            'On Error RESUME Next
            retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
            RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
		next

	next

    '' �ֹ���ȣ ����.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'ezwel','"&confirmDt&"'"
    ' dbget.Execute strSql

	''if IsAutoScript then
		'response.write "�ֹ��Է�(" & successCnt & ")" & "<br />"
	''end if

	Get_ezwelOrderListByStatus = True
	Set xmlDOM = Nothing
	Set objXML = Nothing

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->