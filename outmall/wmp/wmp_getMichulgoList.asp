<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
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


Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''���ֵȳ����� (�ֹ����� view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''�ֹ�Ȯ�� Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''�Է� Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-1,now()),10)

    if eddt<>"" then
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-30,iedyyyymmdd),10)
''

'' ��ȸ ���� (NEW:�ű��ֹ� ,CONFIRM:�߼�ó�����, DELIVERY:�����, COMPLETE:��ۿϷ�)

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'WMP','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ʱ�ȭ�۾�"

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k
for k=0 to datelen-1
    thedate=dateadd("d",-1*k,iedyyyymmdd)
    if k<5 then
    call Get_WMPOrderListByStatus(thedate,thedate,"NEW","�ֹ��뺸")
    response.flush
    end if
    call Get_WMPOrderListByStatus(thedate,thedate,"CONFIRM","�ֹ�Ȯ��")
    response.flush
    call Get_WMPOrderListByStatus(thedate,thedate,"DELIVERY","���Ϸ�")
    response.flush

    '' call Get_WMPOrderListByStatus(istyyyymmdd,iedyyyymmdd,"COMPLETE","��ۿϷ�")
    '' response.flush
next

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'WMP','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ֹ�����"

rw "�Ϸ�"
'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

function Get_WMPOrderListByStatus(stdate,eddate,iorderStatus,istatusName)
	dim sellsite : sellsite = "WMP"
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

	Get_WMPOrderListByStatus = False

	'// =======================================================================
	'// ��¥����
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// API URL(�Ⱓ������ �ֹ� ��������)
	xmlURL = "http://110.93.128.100:8090/wemake/Orders/orderlist"


	postParam = ""
	postParam = postParam & "reqdate=" & stdate
	''postParam = postParam & "&enddate=" & Left(DateAdd("d", 1, CDate(eddate)), 10)
	postParam = postParam & "&type="&iorderStatus
    if (iorderStatus="NEW") then
        postParam = postParam & "&DateType=NEW"     ''�����Ϸ���
    else
        postParam = postParam & "&DateType=CONFIRM" ''�ֹ�Ȯ����
    end if
	''response.write postParam

    rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" ����:"&iorderStatus&"("&istatusName&")"
	'// =======================================================================
	'// ����Ÿ ��������


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL&"?"&postParam, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.send()

	if objXML.Status <> "200" then
		response.write "ERROR : ��ſ���" & objXML.Status
		dbget.close : response.end
	end if

    Dim iRbody, strObj, orderCount, obj1, obj2, obj3


    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo
	Dim invoiceUpDt, outjFixedDt

	iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'rw"<textarea cols=80 rows=20>"&iRbody&"</textarea>"
'exit function

    Set strObj = JSON.parse(iRbody)
    if Not isObject(strObj.outPutValue.data.bundle) then
        rw "No outPutValue"
        exit function
    end if

        set obj1 = strObj.outPutValue.data.bundle

        'rw strObj.totalcount & ":" &obj1.length

        If obj1.length >0 Then
            response.write "�ֹ��Ǽ�(" & obj1.length & ") " & "<br />"
            for i=0 to obj1.length-1
                ordNo           = obj1.get(i).bundleNo				'�ֹ���ȣ(��۹�ȣ)

                shppSeq			= ""			'������û󼼹�ȣ
                reOrderYn ="N" ''���ֹ�����
                delayNts  =""  ''�����ϼ�

                shppTypeDtlNm   = obj1.get(i).delivery.shipMethod
                delicoVenId     = ""	           								'�ù��ۻ��ڵ�
                wblNo           = obj1.get(i).delivery.invoiceNo									'������ȣ
                if (shppTypeDtlNm="��Ÿ���") then
                    wblNo = wblNo & obj1.get(i).delivery.shipMethodMessage					'��۹�� �޼��� ��۹���� [��Ÿ���]�� ��� �Է¹޴� �޼���
                end if
                delicoVenNm     = obj1.get(i).delivery.parcelCompany
                orderStatus     = obj1.get(i).delivery.shipStatus              '���ּ����� | ACCEPT/INSTRUCT/DEPARTURE/DELIVERING/FINAL_DELIVERY/NONE_TRACKING

                whoutCritnDt    = obj1.get(i).originShipDate	 '' �߼۱���.
                outjFixedDt     = obj1.get(i).shipCompleteDate ''����Ȯ������  - ��ü�����ΰ�� 7���� �Ϸ�ȴ�. ������ �ȵǸ� ��ü�������� �����ؾ��Ѵ�.



                set obj2 = obj1.get(i).orderProduct
                    For j=0 to obj2.length-1
                        shppNo		    = obj2.get(j).orderNo			    'reserve01(�ֹ���ȣ)

                        cspGoodsCd      = obj2.get(j).sellerProductCode	'��ü��ǰ�ڵ�
                        goodsCd         = obj2.get(j).productNo		                    '�ǸŻ�ǰ�ڵ�
                        uitemId         = obj2.get(j).sellerProductCode				                '��ǰ�Ӽ��ڵ�

                        shppDivDtlNm = ""
                        ' if (obj2.get(j).canceled="true") then
                        '     shppDivDtlNm = "���"
                        ' end if
                        ' if (obj2.get(j).cancelCount<>0) then
                        '     shppDivDtlNm = shppDivDtlNm & CHKIIF(shppDivDtlNm<>"","/","") & obj2.get(j).cancelCount      ''��Ҽ���
                        ' end if

                        shppRsvtDt      = ""''������
                        autoShortgYn    = "" ''�ڵ���ǰ����
                        invoiceUpDt = "" ''������ȣ ���ε� �Ͻ� (�̰� �����ȰŸ� ����(����)�� �ȵȰ� �� �� �ִ�.)


                        set obj3 = obj2.get(j).orderOption
						    For k=0 to obj3.length-1
                                ordItemSeq      = obj3.get(k).orderOptionNo		'�ֹ��ɼǹ�ȣ
                                uitemId		    = obj3.get(k).optionNo			'�ɼǹ�ȣ
                                optionContent	= obj3.get(k).optionName		'�ɼ�
                                orderQty		= obj3.get(k).optionQty			'����

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

                            next



    'rw bufStr
                        ' if (whoutCritnDt<>"") then
                        '     whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                        ' end if


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

                            ,Array("@invoiceUpDt"	    , adVarchar		, adParamInput		,   19, Trim(invoiceUpDt)) _
                            ,Array("@outjFixedDt"		, adVarchar		, adParamInput		,   19, Trim(outjFixedDt)) _

                        )

                        'On Error RESUME Next
                        retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
                        ' If ERR then
                        '     rw invoiceUpDt
                        '     rw outjFixedDt
                        '     response.end
                        ' end if
                        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�

                        successCnt = successCnt+1
                    next
                set obj2 = nothing
            next
            set obj1 = nothing
        End If
    Set strObj = nothing

    '' �ֹ���ȣ ����.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "�󼼰Ǽ�:"&successCnt
    rw "======================================"

	Get_WMPOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->