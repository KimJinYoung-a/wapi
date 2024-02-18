<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
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
function getHmallDlvCode2Name(idlvCd)
    if isNULL(idlvCd) then Exit function

    SELECT CASE idlvCd

        CASE "12" : getHmallDlvCode2Name = "CJ�������"
        CASE "11" : getHmallDlvCode2Name = "�Ե��ù�"
        CASE "13" : getHmallDlvCode2Name = "�����ù�"
        CASE "33" : getHmallDlvCode2Name = "�����ù�"
        CASE "35" : getHmallDlvCode2Name = "��ü��"

        CASE "61" : getHmallDlvCode2Name = "�浿�ù�"
        CASE "38" : getHmallDlvCode2Name = "�Ͼ��ù�"
        CASE "70" : getHmallDlvCode2Name = "�ǿ��ù�"
        CASE "64" : getHmallDlvCode2Name = "õ���ù�"
        CASE "63" : getHmallDlvCode2Name = "ȣ���ù�"
        CASE "69" : getHmallDlvCode2Name = "����ù�"
        CASE "68" : getHmallDlvCode2Name = "�յ��ù�"
        CASE "71" : getHmallDlvCode2Name = "GTX������"
        CASE "65" : getHmallDlvCode2Name = "CU POST"
        CASE "29" : getHmallDlvCode2Name = "KGB�ù�"
        CASE "74" : getHmallDlvCode2Name = "FLF�۷����ù�"
        CASE "60" : getHmallDlvCode2Name = "������"

        CASE "16" : getHmallDlvCode2Name = "�ڰ����"


        ' CASE "1082" : getHmallDlvCode2Name = "��Ÿ�ù�"
        ' CASE "1001" : getHmallDlvCode2Name = "DHL"
        ' CASE "1011" : getHmallDlvCode2Name = "���ο�ĸ"
        ' CASE "1012" : getHmallDlvCode2Name = "��ü���ù�EMS"
        ' CASE "1080" : getHmallDlvCode2Name = "KG�������ù�"
        ' CASE "1081" : getHmallDlvCode2Name = "��ü�����"
        ' CASE "1103" : getHmallDlvCode2Name = "���ǻ���ù�"
        ' CASE "1104" : getHmallDlvCode2Name = "�ٵ帲"
        ' CASE "1105" : getHmallDlvCode2Name = "������"
        ' CASE "1108" : getHmallDlvCode2Name = "CJ��������Ư��"
        ' CASE "1109" : getHmallDlvCode2Name = "EMS"
        ' CASE "1110" : getHmallDlvCode2Name = "�ѵ���"
        ' CASE "1111" : getHmallDlvCode2Name = "FedEx"
        ' CASE "1112" : getHmallDlvCode2Name = "UPS"
        ' CASE "1113" : getHmallDlvCode2Name = "TNT"
        ' CASE "1114" : getHmallDlvCode2Name = "USPS"
        ' CASE "1115" : getHmallDlvCode2Name = "i-parcel"
        ' CASE "1116" : getHmallDlvCode2Name = "GSM NtoN"
        ' CASE "1117" : getHmallDlvCode2Name = "�����۷ι�"
        ' CASE "1118" : getHmallDlvCode2Name = "�������佺"
        ' CASE "1119" : getHmallDlvCode2Name = "ACI Express"
        ' CASE "1121" : getHmallDlvCode2Name = "���۷ι�"
        ' CASE "1122" : getHmallDlvCode2Name = "������ͽ�������"
        ' CASE "1123" : getHmallDlvCode2Name = "KGL��Ʈ����"
        ' CASE "1124" : getHmallDlvCode2Name = "LineExpress"
        ' CASE "1125" : getHmallDlvCode2Name = "2fast�ͽ�������"
        ' CASE "1126" : getHmallDlvCode2Name = "GSI�ͽ�������"

        CASE ELSE : getHmallDlvCode2Name = idlvCd
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
    istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

Dim strSql : strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'hmall1010','"&confirmDt&"'"
dbget.Execute strSql
rw "�ʱ�ȭ�۾�"

'' �ִ� 7�ϰ� �����ϴ�. 7*3 =21 ��

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")  
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-5,iedyyyymmdd),10)

'' P0:�����, P1:�������, P2:���, P3:��ۿϷ�
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P0","�ֹ���Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P1","�ֹ�Ȯ��")
response.flush
call Get_HmallOrderListByStatus(istyyyymmdd,iedyyyymmdd,"P2","���Ϸ�")
response.flush

strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'hmall1010','"&confirmDt&"'"
dbget.Execute strSql
rw "�ֹ�����"

rw "�Ϸ�"
'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

function Get_HmallOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "hmall1010"
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

	Get_HmallOrderListByStatus = False

	'// =======================================================================
	'// ��¥����
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(�Ⱓ������ �ֹ� ��������)
	xmlURL = "http://xapi.10x10.co.kr:8080/Orders/Hmall/output"


	postParam = ""
	postParam = postParam & "startdate=" & Replace(stdate, "-", "")
	postParam = postParam & "&enddate=" & Replace(Left(DateAdd("d", 1, CDate(eddate)), 10), "-", "")
	postParam = postParam & "&prgrGb="&iorderStatus
	''response.write postParam

    rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" ����:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// ����Ÿ ��������


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL&"?"&postParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=EUC-KR"
	objXML.send()

	if objXML.Status <> "200" then
		response.write "ERROR : ��ſ���" & objXML.Status
		dbget.close : response.end
	end if

    Dim iRbody, strObj, orderCount, obj1


    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim  optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim  orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo


	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

        Set strObj = JSON.parse(iRbody)
        orderCount = strObj.count
        If orderCount >0 Then
            'response.write "�Ǽ�(" & orderCount & ") " & "<br />"
            ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] '"&sellsite&"','"&confirmDt&"'"
            ' dbget.Execute strSql

            set obj1 = strObj.lstorder
            	for i=0 to obj1.length-1
                    orOrdNo="": orordItemSeq=""

                    ordNo           = obj1.get(i).ordNo					        '�ֹ���ȣ
                    ordItemSeq      = obj1.get(i).ordPtcSeq				    '�ֹ��Ϸù�ȣ
                    shppNo		    = obj1.get(i).dlvstNo				'������ù�ȣ
                    shppSeq			= obj1.get(i).dlvstPtcSeq			'������û󼼹�ȣ
                    reOrderYn ="N" ''���ֹ�����
                    delayNts  =""  ''�����ϼ�
                    cspGoodsCd      = obj1.get(i).venItemCd				'���»� ��ǰ�����ڵ�
                    goodsCd         = obj1.get(i).slitmCd				    '�ǸŻ�ǰ�ڵ�
                    uitemId         = obj1.get(i).uitmCd				'��ǰ�Ӽ��ڵ�
                    orderQty        = obj1.get(i).ordQty				'�ֹ� ����

                    shppDivDtlNm = ""
                    if (obj1.get(i).dlvTypeGbcd="40") then
                        shppDivDtlNm = "��ȯ���"
                    end if
                    if (obj1.get(i).dlvCnclNm<>"") then
                        shppDivDtlNm = shppDivDtlNm & CHKIIF(shppDivDtlNm<>"","/","") & obj1.get(i).dlvCnclNm             '������ (''��۱��л󼼸� (���/��ȯ���..))  //nshipTypeGbcd ��������������ڵ�	�������:10, ǰ�����:20
                    end if

                    if (shppDivDtlNm = "��ü���") then
                        shppDivDtlNm = "�ֹ����"
                    elseif (shppDivDtlNm = "��ȯ���/��ü���") then
                        shppDivDtlNm = "��ȯ���öȸ"
                    end if

                    optionContent   = obj1.get(i).uitmTotNm				'��ǰ�Ӽ���
                    shppRsvtDt      = ""''������
                    whoutCritnDt    = obj1.get(i).lastOshpDlineDt		    '�������������  (''��������)
                    autoShortgYn    = "" ''�ڵ���ǰ����

                    orderStatus     = obj1.get(i).lastDlvstPrgrGbcd		        '��������������౸���ڵ� | 25:�����, 30:�������, 45:���, 50:��ۿϷ�

                    shppTypeDtlNm = ""
                    delicoVenId     = obj1.get(i).dsrvDlvcoCd								'�ù��ۻ��ڵ�
                    wblNo           = obj1.get(i).invcNo									'������ȣ
                    delicoVenNm     = getHmallDlvCode2Name(delicoVenId)


                    bufStr = ""
                    bufStr = sellsite&"|"&ordNo
                    bufStr = bufStr &"|"&ordItemSeq
                    bufStr = bufStr &"|"&shppNo
                    bufStr = bufStr &"|"&shppSeq
                    bufStr = bufStr &"|"&cspGoodsCd
                    bufStr = bufStr &"|"&goodsCd

                    bufStr = bufStr &"|"&uitemId
                    bufStr = bufStr &"|"&orderQty
                    bufStr = bufStr &"|"&shppDivDtlNm

                    bufStr = bufStr &"|"&optionContent
                    bufStr = bufStr &"|"&whoutCritnDt


                    bufStr = bufStr &"|"&orderStatus
                    bufStr = bufStr &"|"&shppTypeDtlNm
                    bufStr = bufStr &"|"&delicoVenId
                    bufStr = bufStr &"|"&wblNo
                    bufStr = bufStr &"|"&delicoVenNm
'rw bufStr
                    if (whoutCritnDt<>"") then
                        whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                    end if


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

                        ,Array("@shppTypeDtlNm"		    , adVarchar		, adParamInput		,   16, Trim(shppTypeDtlNm)) _
                        ,Array("@delicoVenId"		    , adVarchar		, adParamInput		,   16, Trim(delicoVenId)) _
                        ,Array("@delicoVenNm"		    , adVarchar		, adParamInput		,   32, Trim(delicoVenNm)) _
                        ,Array("@wblNo"		            , adVarchar		, adParamInput		,   32, Trim(wblNo)) _
                        ,Array("@invoiceUpDt"		    , adVarchar		, adParamInput		,   19, Trim("")) _
                        ,Array("@outjFixedDt"		    , adVarchar		, adParamInput		,   19, Trim("")) _

                        ,Array("@OrgOutMallOrderSerial"	, adVarchar		, adParamInput		,   32, Trim(orordNo)) _
                        ,Array("@OrgOrgDetailKey"		, adVarchar		, adParamInput		,   32, Trim(orordItemSeq)) _
                    )

                    'On Error RESUME Next
                    retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
                    RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�

                     successCnt = successCnt+1
                Next
            set obj1 = nothing
        End If
    Set strObj = nothing

    '' �ֹ���ȣ ����.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "�Ǽ�:"&successCnt&"======================================"

	Get_HmallOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
