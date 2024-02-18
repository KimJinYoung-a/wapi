<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/order/lib/xSiteOrderLib.asp"-->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<%
'' TLS 1.2�� �������� �ʴ� ������ �ִµ���..
const Option_TLS12 = 2048
const Option_TLS1 = 512
const Option_TLS = 128

'' 1. ������ø����ȸ
'' 2. �ֹ� Ȯ�� ó��
'' 3. �Ͱ� ����� ��ȸ

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''���ֵȳ����� (�ֹ����� view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''�ֹ�Ȯ�� Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''�Է� Proc

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE '' SaveOrderToDB

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = replace(LEFT(now(),10),"-","")
    istyyyymmdd = replace(dateadd("d",-7,LEFT(iedyyyymmdd,4)&"-"&Mid(iedyyyymmdd,5,2)&"-"&Mid(iedyyyymmdd,7,2)),"-","")

'response.write istyyyymmdd&":"&iedyyyymmdd
'response.end

'istyyyymmdd = "20180226"
'iedyyyymmdd = "20180226"

''if (istyyyymmdd<"20180219") then istyyyymmdd="20180219" ''2018/02/19  ''�� ���޶� ������ ���Ƽ� �Է��� �ȵǴµ��Ͽ� �Ϸ羿 �Է��Ͽ���.


isOnlyTodayBaljuView = (request("rcvtp")="1") '' justView
isDlvConfirmProc = (request("rcvtp")="2") '�ֹ�Ȯ�� ����.
isDlvInputProc   = (request("rcvtp")="3") '�ֹ��Է�

if (request("targetdt")<>"") then
	istyyyymmdd = request("targetdt")
	iedyyyymmdd = istyyyymmdd
end if

if (isOnlyTodayBaljuView) then  ''�ֹ�Ȯ���� ���� View
	call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd)
elseif (isDlvInputProc) then    ''�ֹ�Ȯ���� ������������ �ֹ��Է�
    call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd)
elseif (isDlvConfirmProc) then  ''�ֹ���Ȯ�ΰ� Ȯ��ó��Proc
	call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)
else
    call getSsgDlvReqList(istyyyymmdd,iedyyyymmdd)     '' 7�ϰ� //ActConfirmDlvReq �� �ʹ� ���� ȣ���ϸ� SSL ���� ������ �߻� Ȯ��ó���� �� �� �Ŀ� �̰��� �ּ�ó���� �ֹ��� ���ܿ��� ����..
    call getSsgDlvConfirmList(iedyyyymmdd,iedyyyymmdd) '' �ֹ�Ȯ���� ����.(���� Ȯ���Ѱ����� �ص� ����.)  // �����߻��� iedyyyymmdd ���� ����.
end if

''call getSsgDlvConfirmList(istyyyymmdd,iedyyyymmdd)

''������ø�� ��ȸ
public function getSsgDlvReqList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim ordNo, ordItemSeq, orOrdNo, orordItemSeq, shppNo, shppSeq, ordStatCd, shppStatCd, shppStatNm, itemId, itemNm, splVenItemId
    Dim ordCstId, ordCstOccCd, shppcst, shppcstCodYn, ordRcpDts, ordpeNm, rcptpeNm, rcptpeHpno, rcptpeTelno, shppDivDtlCd, shppProgStatDtlCd, shppRsvtDt
    Dim uitemId, uitemNm, siteNo, rsvtItemYn, frgShppYn, dircItemQty, cnclItemQty, ordQty, splprc, sellprc, ordCmplDts, ordpeHpno
    Dim shpplocAddr, shpplocZipcd, shpplocOldZipcd, ordMemoCntt, ordpeRoadAddr, ordShpplocId, shppTypeDtlCd, reOrderYn, itemDiv, shpplocBascAddr, shpplocDtlAddr, ordItemDivNm

    Dim ArrShppNo, ArrShppSeq, ArrshppStatCd

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listShppDirection.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestShppDirection>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"
    requestBody = requestBoDy&"</requestShppDirection>"

	objXML.send(requestBody)
	''rw objXML.status
'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
'response.end

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/shppDirections/shppDirection")
			If Not (LagrgeNode Is Nothing) Then
			    ''response.write "�Ǽ�:" & LagrgeNode.length
			    redim ArrShppNo(LagrgeNode.length-1)
			    redim ArrShppSeq(LagrgeNode.length-1)
			    redim ArrshppStatCd(LagrgeNode.length-1)

			    For i = 0 To LagrgeNode.length - 1
    			    ordNo="": ordItemSeq="": orOrdNo="": orordItemSeq="": shppNo="": shppSeq="": ordStatCd="": shppStatCd=""
                    shppStatNm="": itemId="": itemNm="": splVenItemId="": ordCstId="": ordCstOccCd="": shppcst="": shppcstCodYn=""
                    ordRcpDts="": ordpeNm="": rcptpeNm="": rcptpeHpno="": rcptpeTelno="": shppDivDtlCd="": shppProgStatDtlCd="": shppRsvtDt=""
                    uitemId="": uitemNm="": siteNo="": rsvtItemYn="": frgShppYn="": dircItemQty="": cnclItemQty="": ordQty="": splprc="": sellprc=""
                    ordCmplDts="": ordpeHpno="": shpplocAddr="": shpplocZipcd="": shpplocOldZipcd="": ordMemoCntt="": ordpeRoadAddr="": ordShpplocId=""
                    shppTypeDtlCd="": reOrderYn="": itemDiv="": shpplocBascAddr="": shpplocDtlAddr="": ordItemDivNm=""

                    shppNo           = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''**��۹�ȣ [D2125835493]
    			    shppSeq         = LagrgeNode(i).SelectSingleNode("shppSeq").Text               ''**��ۼ��� [1]
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*�ֹ���ȣ [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*�ֹ�����? [1]
					If NOT (LagrgeNode(i).SelectSingleNode("orOrdNo") is Nothing) Then
    			    	orOrdNo          = LagrgeNode(i).SelectSingleNode("orOrdNo").Text                ''���ֹ���ȣ
					End If
    			    orordItemSeq     = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text                ''���ֹ�����  orOrdNo
    			    shppStatCd      = LagrgeNode(i).SelectSingleNode("shppStatCd").Text            '' *��ۻ����ڵ� 10 ���� 30 ���[10]

    			    ArrShppNo(i) = shppNo
    			    ArrShppSeq(i) = shppSeq
    			    ArrshppStatCd(i) = shppStatCd

'    			    ordStatCd       = LagrgeNode(i).SelectSingleNode("ordStatCd").Text              ''          [120]
    			    shppStatNm      = LagrgeNode(i).SelectSingleNode("shppStatNm").Text            '' ��ۻ��¸�         [����]
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''��ǰ��ȣ  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''��ǰ��    [���ֶ���ƼĿ]
					If NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) Then
	    			    splVenItemId    = LagrgeNode(i).SelectSingleNode("splVenItemId").Text        ''��ü��ǰ��ȣ [1024019]
					Else
						strSql = ""
						strSql = strSql & " select top 1 itemid "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem"
						strSql = strSql & " where ssgGoodNo = '"& itemId &"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							splVenItemId = rsget("itemid")
						Else
							rw "���� �ֹ���ȣ : " & ordNo
						End If
						rsget.Close

						If session("ssBctID")="kjy8517" Then
							response.write "<textarea cols=100 rows=30>!!"&objXML.responseText&"</textarea>"
						End If

					End If
    			    ordCstId        = LagrgeNode(i).SelectSingleNode("ordCstId").Text                ''�ֹ������̵�
    			    ordCstOccCd     = LagrgeNode(i).SelectSingleNode("ordCstOccCd").Text             ''�ֹ����߻��ڵ� [�ΰ�] :: 01,02�� �ƴ�
    			    shppcst         = LagrgeNode(i).SelectSingleNode("shppcst").Text                 ''��ۺ�?
    			    shppcstCodYn    = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''��ۺ����ҿ��� : Y :����,N :���� [N]
    			    ordRcpDts       = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text            ''�ֹ������Ͻ� [2017-11-27 09:32:31.0]
					ordpeNm			= LagrgeNode(i).SelectSingleNode("ordpeNm").Text				 ''�ֹ���
					rcptpeNm 		= LagrgeNode(i).SelectSingleNode("rcptpeNm").Text				 ''������

    			    rcptpeHpno      = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text                ''������ �޴���
    			    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
        			    rcptpeTelno     = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text                ''������ ��ȭ [--]
        			end if
    			    shppDivDtlCd    = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text               ''��۱��л��ڵ� 11 �Ϲ���� 12 �κ���� 14 ���� 15 ��ȯ��� 16 AS��� [11]
    			    shppProgStatDtlCd = LagrgeNode(i).SelectSingleNode("shppProgStatDtlCd").Text        ' ������ۻ���������ڵ�(��ۻ�ǰ����) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���	[11]
    			    shppRsvtDt      = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text                 ''[20171128]
    			    uitemId         = LagrgeNode(i).SelectSingleNode("uitemId").Text                 ''��ǰID [00000]

    			    siteNo          = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''����Ʈ��ȣ 6001 �̸�Ʈ�� 6002 Ʈ���̴����� 6003 �н��� 6004 �ż���� 6005 S.COM�� 6009 �ż����ȭ����[6004]
    			    rsvtItemYn      = LagrgeNode(i).SelectSingleNode("rsvtItemYn").Text                 ''�����Ǹű��� [N]
					' If NOT (LagrgeNode(i).SelectSingleNode("frgShppYn") is Nothing) then
    			    ' 	frgShppYn       = LagrgeNode(i).SelectSingleNode("frgShppYn").Text                 ''����/�� ���� [N]
					' End If

    			    dircItemQty     = LagrgeNode(i).SelectSingleNode("dircItemQty").Text                 ''���ü��� [2]
    			    cnclItemQty     = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text                 ''��Ҽ��� [2]
    			    ordQty          = LagrgeNode(i).SelectSingleNode("ordQty").Text                 ''�ֹ����� [2]
    			    splprc          = LagrgeNode(i).SelectSingleNode("splprc").Text                 ''���ް� [755]
    			    sellprc         = LagrgeNode(i).SelectSingleNode("sellprc").Text                 ''�ǸŰ� [1000]


    			    if NOT (LagrgeNode(i).SelectSingleNode("ordCmplDts") is Nothing) then
    			        ordCmplDts      = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text                 ''�ֹ��Ϸ��Ͻ� [2017-11-27 09:32:31.0]
    			    end if
    			    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
    			        ordpeHpno       = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text                 ''�ֹ����޴�����ȣ [01091603979]
    			    end if
    			    shpplocAddr     = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text                 ''[���� ������ ������ 561���� ���̾ȿ��������� 104�� 103ȣ]
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocZipcd") is Nothing) then
    			    	shpplocZipcd    = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text                 ''*������ �����ȣ [04733]
					end if
    			    if NOT (LagrgeNode(i).SelectSingleNode("shpplocOldZipcd") is Nothing) then
    			        shpplocOldZipcd = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text                 ''������(��) �����ȣ[133750]
    			    end if

    			    ordpeRoadAddr   = LagrgeNode(i).SelectSingleNode("ordpeRoadAddr").Text                 ''[���� ������ �ź��� 15, 104�� 103ȣ (������, ���̾ȿ���������)]
    			    ordShpplocId    = LagrgeNode(i).SelectSingleNode("ordShpplocId").Text                 ''�ֹ������ID [1102603504]
    			    shppTypeDtlCd   = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text                 ''����������ڵ� 14 ��ü�ڻ��� 22 ��ü�ù��� 25 �ؿ��ù��� 31 ����湮 41 ��� 51 SMS 52 EMAIL 61 �̹߼�  [22]
    			    reOrderYn       = LagrgeNode(i).SelectSingleNode("reOrderYn").Text                 ''*�����ÿ��α���  [N]
    			    itemDiv         = LagrgeNode(i).SelectSingleNode("itemDiv").Text                 ''�ǸźҰ���û���� 10:�Ϲ� 20: ���� GIFT �Ϲ� 30: ���� GIFT ���� 40: ���� GIFT ���� [10]
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocBascAddr") is Nothing) then
						shpplocBascAddr = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text                 '' [���� ������ �ź���]
					End If
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocDtlAddr") is Nothing) then
						shpplocDtlAddr  = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)                 ''[15, 104�� 103ȣ (������, ���̾ȿ���������)]
					End If
    			    ordItemDivNm    = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text                 ''[�ֹ�]

    			    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			        ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")                 ''[[����۸޸�]��۸޼���]
    			    end if

    			    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
    			        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:�ֹ�����1,2:^:asdasdddd:^:�ֹ�����2]
    			    end if

			    Next

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	dim newOrdCnt, succConfirmCnt, irefErrStr
	newOrdCnt = 0 : succConfirmCnt = 0 : irefErrStr = ""
	if IsArray(ArrShppNo) then
	    for i=LBound(ArrShppNo) to UBound(ArrShppNo)
	        response.flush

	        '' response.write ArrShppNo(i)&":"&ArrShppSeq(i)
	        if (ArrShppNo(i)<>"") and (ArrShppSeq(i)<>"") then
	            if (ArrshppStatCd(i)="10") then
    	            if (ActConfirmDlvReq(ArrShppNo(i),ArrShppSeq(i),irefErrStr)) then
    	                succConfirmCnt=succConfirmCnt+1
    	            else
    	                irefErrStr = irefErrStr&":"&ArrShppNo(i)&":"&ArrShppSeq(i)&"::"
    	            end if
    	            newOrdCnt = newOrdCnt+1
    	        else
	                rw "����ֹ�:"&ArrshppStatCd(i)&":"&ArrShppNo(i)
	            end if

	        end if
	    next
	end if

    rw "========================================="
	rw "�ű��ֹ�Ȯ��:"&styyyymmdd&"~"&edyyyymmdd
	rw "�ű��ֹ�:"&newOrdCnt&"(��)"
	rw "����Ȯ��:"&succConfirmCnt&"(��)"
	if (irefErrStr<>"") then
	    rw irefErrStr
	end if
end function

''�ֹ�Ȯ��ó��
public function ActConfirmDlvReq(iShppno, iShppSeq, byref iErrStr)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
    Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/updateOrderSubjectManage.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestOrderSubjectManage>"
    requestBody = requestBoDy&"<shppNo>"&iShppno&"</shppNo>"
    requestBody = requestBoDy&"<shppSeq>"&iShppSeq&"</shppSeq>"
    requestBody = requestBoDy&"</requestOrderSubjectManage>"

	objXML.send(requestBody)

	'rw objXML.status
	'response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
		Set xmlDOM = nothing
	Set objXML = nothing

	'response.write 	ssgresultCode&":"&ssgresultMessage&":"&ssgresultMessage&":"&ssgresultDesc

	if (ssgresultCode<>"00") then
	    iErrStr = "["&ssgresultMessage&"]"&ssgresultDesc
	end if
	ActConfirmDlvReq = (ssgresultCode="00")
end function

''�Ͱ� ����� ��ȸ
public function getSsgDlvConfirmList(styyyymmdd,edyyyymmdd)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode

    Dim ssgresultCode, ssgresultMessage, ssgresultDesc

    Dim shppNo,shppSeq,shppTabProgStatCd,evntSeq,shppDivDtlCd,shppDivDtlNm,reOrderYn,delayNts,ordNo,ordItemSeq,ordCmplDts
    Dim lastShppProgStatDtlNm,lastShppProgStatDtlCd,salestrNo,shppVenId,shppVenNm,shppTypeNm,shppTypeCd,shppTypeDtlCd,shppTypeDtlNm,delicoVenId,boxNo
    Dim shppcst,shppcstCodYn,itemNm,splVenItemId,itemId,uitemId,dircItemQty,cnclItemQty,ordQty,sellprc,frgShppYn
    Dim ordpeNm,rcptpeNm,rcptpeHpno,rcptpeTelno,shpplocAddr,shpplocZipcd,shpplocOldZipcd,shpplocRoadAddr,itemChrctDivCd,shppStatCd,shppStatNm
    Dim orordNo,orordItemSeq,shppMainCd,siteNo,siteNm,shppRsvtDt,splprc,shortgYn,newWblNoData,newRow,itemDiv
    Dim shpplocBascAddr,shpplocDtlAddr,ordItemDivNm
    Dim ordpeHpno, ordMemoCntt, pCus, frebieNm ,shortgProgStatCd, shortgProgStatNm, uitemNm
    Dim iBufrequireDetail

    Dim oMaster, oDetailArr(0)
    Dim successCnt : successCnt=0
    Dim failCnt : failCnt=0

    'On Error Resume Next
	'Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Set objXML = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
    objXML.Option(9) = Option_TLS
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"

	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>"&styyyymmdd&"</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>"&edyyyymmdd&"</perdEndDts>"  ''�Ϸ縦 ���ؾ�?
    requestBody = requestBoDy&"</requestWarehouseOut>"
	objXML.send(requestBody)

'rw objXML.status
if (isOnlyTodayBaljuView) then
    response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
    'response.end
end if
    dim retBody : retBody=objXML.responseText
    retBody = replace(retBody,"&","")
	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(retBody) ''objXML.responseText)

			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text

			Set LagrgeNode = xmlDOM.SelectNodes("/result/warehouseOuts/warehouseOut")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        ''�����ʱ�ȭ.
			        shppNo ="": shppSeq = "": shppTabProgStatCd ="": evntSeq ="": shppDivDtlCd =""
                    shppDivDtlNm ="": reOrderYn ="": delayNts ="": ordNo ="": ordItemSeq =""
                    ordCmplDts ="": lastShppProgStatDtlNm = "": lastShppProgStatDtlCd ="": salestrNo ="": shppVenId =""
                    shppVenNm ="": shppTypeNm ="": shppTypeCd ="": shppTypeDtlCd ="": shppTypeDtlNm =""
                    delicoVenId ="": boxNo ="": shppcst ="": shppcstCodYn ="": itemNm =""
                    splVenItemId ="":itemId ="":uitemId ="": dircItemQty ="": cnclItemQty =""
                    ordQty ="" :sellprc ="": frgShppYn ="": ordpeNm =""
                    rcptpeNm ="" :rcptpeHpno ="": rcptpeTelno ="": shpplocAddr =""
                    shpplocZipcd ="": shpplocOldZipcd ="": shpplocRoadAddr ="": itemChrctDivCd =""
                    shppStatCd ="": shppStatNm ="": orordNo ="": orordItemSeq ="": shppMainCd =""
                    siteNo ="": siteNm ="": shppRsvtDt ="": splprc ="": shortgYn =""
                    newWblNoData ="": newRow ="": itemDiv ="": shpplocBascAddr ="": shpplocDtlAddr ="": ordItemDivNm =""

                    ordpeHpno = "": ordMemoCntt = "": pCus = "": frebieNm = "": shortgProgStatCd ="": shortgProgStatNm ="" : uitemNm=""
                    iBufrequireDetail = ""

			        shppNo              = LagrgeNode(i).SelectSingleNode("shppNo").Text                 ''*��۹�ȣ
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*��ۼ���
                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''������ۻ���������ڵ�(��۴���) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
					if NOT (LagrgeNode(i).SelectSingleNode("evntSeq") is Nothing) then
                    	evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''�̺�Ʈ����
					end if
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*��۱��л��ڵ� 11 �Ϲ���� 12 �κ���� 14 ���� 15 ��ȯ��� 16 AS���
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''��۱��л󼼸�
                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*�����ÿ��α���
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''����Ƚ��
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*�ֹ���ȣ [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*�ֹ�����
                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*�ֹ��Ϸ��Ͻ� [2017-11-23 10:39:42.0]
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''������ۻ�������¸�(��ۻ�ǰ����) [��ŷ�Ϸ�]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ�(��ۻ�ǰ����) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
'                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''���޾�ü���̵� [0000003198]
                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''���޾�ü��
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''���������    [�ù���]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''��������ڵ� 10 �ڻ��� 20 �ù��� 30 ����湮 40 ��� 50 �̹�� 60 �̹߼�
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''����������ڵ� 14 ��ü�ڻ��� 22 ��ü�ù��� 25 �ؿ��ù��� 31 ����湮 41 ��� 51 SMS 52 EMAIL 61 �̹߼�
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''��������󼼸� [��ü�ù���]
					if NOT (LagrgeNode(i).SelectSingleNode("delicoVenId") is Nothing) then
                    	delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''�ù��ID [0000033011]
					End If
                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''�ڽ���ȣ [398327952]
                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' ��ۺ�? [303] ??
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*��ۺ� ���ҿ��� Y: ���� N: ����
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*��ǰ��
					itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*��ǰ��ȣ [1000024811163]
					if NOT (LagrgeNode(i).SelectSingleNode("splVenItemId") is Nothing) then
	                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*��ü��ǰ��ȣ [1024019]
					Else
						strSql = ""
						strSql = strSql & " select top 1 itemid "
						strSql = strSql & " FROM db_etcmall.dbo.tbl_ssg_regitem"
						strSql = strSql & " where ssgGoodNo = '"& itemId &"' "
						rsget.CursorLocation = adUseClient
						rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
						If Not rsget.Eof Then
							splVenItemId = rsget("itemid")
						Else
							rw "���� �ֹ���ȣ : " & ordNo
						End If
						rsget.Close

						If session("ssBctID")="kjy8517" Then
							response.write "<textarea cols=100 rows=30>##"&objXML.responseText&"</textarea>"
						End If
					End If

                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*��ǰID [00000]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''���ü��� [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''��Ҽ��� [0]
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''�ֹ����� [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''�ǸŰ� [1000]
					If NOT (LagrgeNode(i).SelectSingleNode("frgShppYn") is Nothing) then
    			    	frgShppYn       = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''����/�� ���� [����]
					End If
                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*�ֹ���

                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*������
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*������ �޴�����ȣ
                    if NOT (LagrgeNode(i).SelectSingleNode("rcptpeTelno") is Nothing) then
                        rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*������ ����ȭ��ȣ
                    end if
                    shpplocAddr         = LEFT(LagrgeNode(i).SelectSingleNode("shpplocAddr").Text, 500)        ''������ ���ּ�
					if NOT (LagrgeNode(i).SelectSingleNode("shpplocZipcd") is Nothing) then
                    	shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*������ �����ȣ          [04733]
					end if
                    if NOT (LagrgeNode(i).SelectSingleNode("shpplocOldZipcd") is Nothing) then
                        shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*������ �������ȣ(6�ڸ�)  [133750]
                    end if
                    shpplocRoadAddr     = LEFT(LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text, 500)    ''�����ε��θ��ּ�
                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''��ǰƯ�������ڵ� 10 �Ϲ� 20 ���θ� 30 �ؿܱ��Ŵ����ǰ 40 �̰����ͱݼ� 50 ����ϱ���Ʈ 60 ��ǰ�� 70 ���������� 80 ����ϻ�ǰ�� 91 �̺�Ʈ
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*��ۻ����ڵ� 10 ���� 30 ���
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''��ۻ��¸�
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''���ֹ���ȣ [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''���ֹ����� [2]
                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''�����ü�ڵ� 32 ��üâ�� 41 ���¾�ü 42 �귣������  [41]
                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''����Ʈ��ȣ 6001 �̸�Ʈ�� 6002 Ʈ���̴����� 6003 �н��� 6004 �ż���� 6005 S.COM�� 6009 �ż����ȭ����
                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''����Ʈ��
'                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
					if NOT (LagrgeNode(i).SelectSingleNode("splprc") is Nothing) then
                    	splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''���ް�
					end if

                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''�ǸźҰ���û���� 10:�Ϲ� 20: ���� GIFT �Ϲ� 30: ���� GIFT ���� 40: ���� GIFT ����
                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''�������ּ� 20170712
                    shpplocDtlAddr      = LEFT(LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text, 500)         ''�����λ��ּ�	20170712
                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''�ֹ���ǰ����	20170809  // �ֹ�, �κй���ֹ�


                    ''//�ʼ��� �ƴѰ�� .
                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''�ֹ����޴�����ȣ  //���ð�
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
                        ordMemoCntt         = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")            ''����۸޸�  //���ð�
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''�������������ȣ  //���ð�
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''����ǰ  //���ð�
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''�ǸźҰ���û����  //���ð� 11 ��ǰ��� 12 ��ǰCSó���� 13 ��ǰȮ�� 21 ��ǰ����������� 22 ��ǰ��������CSó���� 23 ��ǰ��������Ȯ�� 41 �԰�������� 43 �԰������Ϸ� 51 ����������
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''��ǰ������¸�  //���ð�
                    end if

                    ''�ɼǸ�
                    if NOT (LagrgeNode(i).SelectSingleNode("uitemNm") is Nothing) then
                        uitemNm         = LagrgeNode(i).SelectSingleNode("uitemNm").Text                 ''[,1:^:asdasd:^:�ֹ�����1,2:^:asdasdddd:^:�ֹ�����2]
                    end if

                    if (orordNo<>ordNo) then ''���ֹ���ȣ�� ������Ʈ ''�κ����ó���� �ֹ���ȣ�� �ٲ�
                        ordNo=orordNo
                    end if

                    if (orordItemSeq<>ordItemSeq) then  ''2018/03/05 �߰�  <ordItemDivNm>�κй���ֹ�</ordItemDivNm> 20180305585498
                        ordItemSeq=orordItemSeq
                    end if

                    ''�ֹ��Է�.
                    Set oMaster = new COrderMasterItem
                        oMaster.FSellSite 			= CMALLNAME  ''ssgcom
                        oMaster.FOutMallOrderSerial = ordNo

                        oMaster.FbeasongNum11st     = shppNo     ''��۹�ȣ
                        oMaster.Freserve01          = shppSeq    ''��۹�ȣ-seq

                        oMaster.FSellDate 			= Left(ordCmplDts, 19)
                		oMaster.FPayType			= "50"
                		oMaster.FPaydate			= oMaster.FSellDate
                		oMaster.FOrderUserID		= ""
                		oMaster.FOrderName			= LEFT(html2db(ordpeNm),32) '' �ֹ���ȣ 20180106923841
                		oMaster.FOrderTelNo			= ""
                		oMaster.FOrderHpNo			= ordpeHpno
                		oMaster.FOrderEmail			= ""
                		oMaster.FReceiveName		= LEFT(html2db(rcptpeNm),32)
                		oMaster.FReceiveTelNo		= html2db(rcptpeTelno)
                		oMaster.FReceiveHpNo		= html2db(rcptpeHpno)

                		oMaster.Fdeliverymemo		= html2db(ordMemoCntt)
                		oMaster.FdeliverPay			= shppcst
                        ''if (oMaster.FdeliverPay>0) then oMaster.FdeliverPay=2500 ''��ۺ� �ȺеǾ� �� //�Ʒ��� ó��

                		oMaster.FReceiveZipCode		= shpplocZipcd
		                oMaster.FReceiveAddr1		= html2db(shpplocBascAddr)
			            oMaster.FReceiveAddr2	    = html2db(shpplocDtlAddr)

                		Set oDetailArr(0) = new COrderDetail
                		oDetailArr(0).FdetailSeq = ordItemSeq
                		oDetailArr(0).FItemID = splVenItemId    ''��ü��ǰ��ȣ
                		oDetailArr(0).FItemOption = ""
                		oDetailArr(0).FOutMallItemID = itemId           ''ssg ��ǰ�ڵ�
                		oDetailArr(0).FOutMallItemOption = uitemId      ''ssg ��ǰ�ڵ�
                		oDetailArr(0).FOutMallItemName = html2db(itemNm)
                		oDetailArr(0).FOutMallItemOptionName = (uitemNm)  ''�ɼǸ��� ����? html2db(objMasterOneXML.attributes.GetNamedItem("RequestOption").text)

                		oDetailArr(0).FItemNo = CLng(dircItemQty) ''�ֹ�����-ordQty, ���ü���- dircItemQty

                		oDetailArr(0).Fitemcost = CLNG(sellprc) '' �ܰ� ����. �ܰ����� Ȯ�� Clng(objMasterOneXML.getElementsByTagName("OrderBase")(0).attributes.GetNamedItem("AwardAmount").text) / oDetailArr(0).FItemNo

                		oDetailArr(0).FReducedPrice = oDetailArr(0).Fitemcost  ''���� ���ΰ�(���ǸŰ�) Ȯ��. ToDo :: ���µ���.
                		oDetailArr(0).FOutMallCouponPrice = 0
                		oDetailArr(0).FTenCouponPrice = 0


                		if oDetailArr(0).FOutMallItemOption = "00000" then
                			oDetailArr(0).FItemOption = "0000"
                			''oDetailArr(0).FOutMallItemOption = "0000"  ''�ּ�ó�� 2018/03/06
                		end if

                        '' ToDo : �ɼ��� �ִ°�� : �ɼǸ����� �ɼ��� �����;���..
                		if ((oDetailArr(0).FItemOption <> "0000") or (oDetailArr(0).FOutMallItemOptionName<>"")) then
                		    oDetailArr(0).FItemOption = getOptionCodByOptionNameSSG(oDetailArr(0).FItemID,oDetailArr(0).FOutMallItemOptionName,iBufrequireDetail, uitemId)

							'2020-05-14 ������..&amp;�� replace���� ���� ������ ���� body &�� �������� ó���ؼ���
							iBufrequireDetail = replace(iBufrequireDetail, "amp;", "&")
							iBufrequireDetail = replace(iBufrequireDetail, "quot;", """")
							iBufrequireDetail = replace(iBufrequireDetail, "lt;", "<")
							iBufrequireDetail = replace(iBufrequireDetail, "gt;", ">")

             		        oDetailArr(0).FrequireDetail = iBufrequireDetail
'                			if Not GetCheckItemOptionValid(oDetailArr(0).FItemID, oDetailArr(0).FItemOption) then
'                				'// �߸��� �ɼ�.
'                				tmpOptionSeq = tmpOptionSeq + 1
'                				oDetailArr(0).FItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
'                				oDetailArr(0).FOutMallItemOption = "FF" & Right(CStr(tmpOptionSeq + 100), 2)
'                			end if
                		end if
 ''rw "oDetailArr(0).FItemOption:"&oDetailArr(0).FItemOption&":"&oDetailArr(0).FOutMallItemOptionName
                        ''rw "---------------------------------------"
                        ''�ǵ�� �Է�.


                        IF (isOnlyTodayBaljuView) then
                            response.write oMaster.FOutMallOrderSerial&":"&shppDivDtlCd&":"&oMaster.FbeasongNum11st&":"&oDetailArr(0).FdetailSeq&"<br>"
                        else
                            if NOT ((shppDivDtlCd="11") or (shppDivDtlCd="12")) then
                                rw oMaster.FOutMallOrderSerial&":"&shppDivDtlCd&":CS"
                            else
                                if (SaveOrderToDB(oMaster, oDetailArr) = True) then
                                    successCnt = successCnt + 1

                                    '// siteNo ���� => �Ʒ� ���ν����� ����
        							' strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder] "
        							' strSql = strSql + " set subSellSite = '" & siteNo & "' "
        							' strSql = strSql + " where sellsite = '" & CMALLNAME & "' "
        							' strSql = strSql + " and OutMallOrderSerial = '" + CStr(oMaster.FOutMallOrderSerial) + "' "
        							' strSql = strSql + " and OrgDetailKey = '" + CStr(oDetailArr(0).FdetailSeq) + "' "
        							' ''response.write strSql
        							' dbget.Execute strSql

									''��ۺ� ���ؼ� �ٽ� ����
									strSql = " Exec db_temp.[dbo].[sp_TEN_xSite_TMPOrder_ssg_DlvPayUp] '" + CStr(oMaster.FOutMallOrderSerial) + "','"&CStr(oDetailArr(0).FdetailSeq)&"','"&siteNo&"',"&oMaster.FdeliverPay
									dbget.Execute strSql
                                else
                                    failCnt = failCnt + 1
                                end if
                            end if
                        end if

                        SEt oDetailArr(0) = Nothing
                    SEt oMaster = Nothing
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing

	rw "======================================"
	rw "�ѹ����Ǽ�:"&failCnt+successCnt
	rw "�ֹ��Է°Ǽ�:"&successCnt
end function

''ǰ��/���� ����üũ
sqlStr = "exec [db_temp].[dbo].[usp_TEN_xSiteTmpOrderCHECK_Make]"
dbget.Execute sqlStr
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->