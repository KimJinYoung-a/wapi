<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/ssg/ssgItemcls.asp"-->
<!-- include virtual="/outmall/ssg/incssgFunction.asp"-->
<!-- include virtual="/outmall/incOutmallCommonFunction.asp"-->
<%
'' 1. ������ø����ȸ
'' 2. �ֹ� Ȯ�� ó��
'' 3. �Ͱ� ����� ��ȸ

'CONST ssgAPIURL = "http://eapi.ssgadm.com"
'CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
'CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"

''call getSsgDlvReqList()
''call getSsgDlvConfirmList()

''call getSsgJungsanList()


''���긮��Ʈ
dim ix , istyyyymmdd
istyyyymmdd = "20180301"

    call getSsgJungsanList(istyyyymmdd,"6006")
    call getSsgJungsanList(istyyyymmdd,"6007")

'for ix=1 to 1
'    istyyyymmdd = CStr(CLNG(istyyyymmdd)+1)
'    call getSsgJungsanList(istyyyymmdd,"6006")
'    call getSsgJungsanList(istyyyymmdd,"6007")
'
'    response.write istyyyymmdd&"<br>"
'    response.flush
'next

''������  ��ȸ
public function getSsgJungsanList(yyyymmdd,isiteno)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    
    Dim urlparam : urlparam="critnDt="&yyyymmdd&"&siteNo="&isiteno  '6006/6007
    Dim sqlStr
    
    'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & ssgAPIURL&"/se/alln/getVendorSalesList.ssg"&"?"&urlparam    '''POST �̳� getParam?
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	
	objXML.send()

'	response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
'exit function		
    Dim siteNo,settleDt,orordNo,ordNo,ordItemSeq
    Dim txnDivNm,itemId,itemNm,uitemId,salesQty
    Dim settleAmt,settleVat,salesAmt ,splvenBdnDcAmt 
    Dim owncoBdnDcAmt,netAmt ,mrgrt ,dvShppcstAmt ,dvShppcstVat 
    Dim custBdnShppAmt,splvenBdnShppAmt ,txnDivCd 
    Dim extCommPrice, extVatYN, dvShppcstTTL
    Dim isDataExists : isDataExists = FALSE
    Dim isDlvPayExists 
        Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
	
	        Set LagrgeNode = xmlDOM.SelectNodes("/result/sales/sale")
			If Not (LagrgeNode Is Nothing) Then
			    response.write "<br>�Ǽ�:" & LagrgeNode.length&":"&"<br>"
			    'redim ArrShppNo(LagrgeNode.length-1)
			    'redim ArrShppSeq(LagrgeNode.length-1)
			    
			    IF (LagrgeNode.length>0) then
			        isDataExists = TRUE
			        sqlStr = " delete from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite = 'ssg' "
			        dbget.execute sqlStr
			    end if
			        
			    For i = 0 To LagrgeNode.length - 1
    			    siteNo = "" : settleDt = "" : orordNo ="" : ordNo="" : ordItemSeq =0
    			    txnDivNm = "" : itemId = "" : itemNm = "" : uitemId="" : salesQty=1
    			    settleAmt =0 : settleVat = 0 : salesAmt = 0 : splvenBdnDcAmt = 0
    			    owncoBdnDcAmt = 0 : netAmt =0 : mrgrt = 0.0 : dvShppcstAmt = 0 : dvShppcstVat =0
    			    custBdnShppAmt = 0 : splvenBdnShppAmt = 0 : txnDivCd =""
                    isDlvPayExists = FALSE
                    
                    siteNo           = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''6006: �̸�Ʈ / 6007:�ż���
    			    settleDt         = LagrgeNode(i).SelectSingleNode("settleDt").Text               ''*���� ���� ����	
    			    orordNo          = LagrgeNode(i).SelectSingleNode("orordNo").Text                ''���ֹ���ȣ
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*�ֹ���ȣ [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*�ֹ�����? [1]
    			    
    			    txnDivNm     = LagrgeNode(i).SelectSingleNode("txnDivNm").Text                ''���������ڵ�� ����/�鼼
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''��ǰ��ȣ  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''��ǰ��    [���ֶ���ƼĿ]
    			    uitemId      = LagrgeNode(i).SelectSingleNode("uitemId").Text                   '' ��ǰ ID	
    			    
    			    salesQty       = LagrgeNode(i).SelectSingleNode("salesQty").Text              ''    ���� ����
    			    settleAmt      = LagrgeNode(i).SelectSingleNode("settleAmt").Text            '' ���� �ݾ�(vat ����)
    			    salesAmt      = LagrgeNode(i).SelectSingleNode("salesAmt").Text            '' ���� �ݾ�(vat ����)
    			    splvenBdnDcAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnDcAmt").Text            ''��ü ���� �δ� �ݾ�	
    			    owncoBdnDcAmt    = LagrgeNode(i).SelectSingleNode("owncoBdnDcAmt").Text            ''�ڻ�(SSG) ���� �δ� �ݾ�
    			    netAmt    = LagrgeNode(i).SelectSingleNode("netAmt").Text            ''�� �Ǹ� �ݾ�(vat ����)
    			    mrgrt    = LagrgeNode(i).SelectSingleNode("mrgrt").Text            ''������
    			    dvShppcstAmt    = LagrgeNode(i).SelectSingleNode("dvShppcstAmt").Text            ''��� �ݾ�(vat ����)	
    			    dvShppcstVat    = LagrgeNode(i).SelectSingleNode("dvShppcstVat").Text            ''��� VAT �ݾ�
    			    custBdnShppAmt    = LagrgeNode(i).SelectSingleNode("custBdnShppAmt").Text            ''�� �δ� ��� �ݾ�
    			    splvenBdnShppAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnShppAmt").Text            ''��ü �δ� ��� �ݾ�
    			    
    			    dvShppcstTTL = CLNG(dvShppcstAmt)+CLNG(dvShppcstVat)
    			    isDlvPayExists = (dvShppcstTTL<>0)
    			    extVatYN = "Y"
                    if (txnDivNm="�鼼") then extVatYN="N"
                    
                    if (extVatYN = "Y") then
    			        settleAmt = CLNG(settleAmt*1.1)  ''2018/03/06 �߰�
    			    end if
    			    
    			    if (isDlvPayExists) then
    			        salesAmt = salesAmt - dvShppcstTTL
    			        settleAmt = settleAmt - dvShppcstTTL
    			        netAmt = netAmt - dvShppcstTTL
    			    end if
    			    
    			    
    			    if (salesQty<0) then  ''��ǰ�ΰ��.
    			        ordNo = ordNo&"-"&ordItemSeq
    			    end if
    			    
    			    'if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			    '    ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")                 ''[[����۸޸�]��۸޼���]
    			    'end if
    			    
    			    response.write siteNo&"|"&settleDt&"|"&orordNo&"|"&ordNo&"|"&ordItemSeq&"|"&txnDivNm&"|"&itemId&"|"&itemNm&"|"&uitemId&"|"
    			    response.write salesQty&"|"&settleAmt&"|"&salesAmt&"|"&splvenBdnDcAmt&"|"&owncoBdnDcAmt&"|"&netAmt&"|"&mrgrt&"|"&dvShppcstAmt&"|"&dvShppcstVat&"|"&custBdnShppAmt&"|"&custBdnShppAmt&"|"&splvenBdnShppAmt
    			    response.write "<br>"

                    ''extJungsanType : C-��ǰ, D��ۺ�
                    '' extTenMeachulPrice=extCommPrice+extTenJungsanPrice
                    '' extItemCost=extReducedPrice+extOwnCouponPrice+extTenCouponPrice
                    
                    extCommPrice = salesAmt-settleAmt
                    
                    if (orordNo=ordNo) then orordNo=""
                        
                    if (isDlvPayExists and (salesQty=0) ) then
                        '' ��ǰ��ۺ�
                    elseif ((NOT isDlvPayExists) and (salesQty=0) and (salesAmt=0)and (settleAmt=0))then
                        '' �糦.
                    else 
                            
        			    sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
        			    sqlStr = sqlStr + " (sellsite, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost"
        			    sqlStr = sqlStr + " , extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice"
        			    sqlStr = sqlStr + " , extTenJungsanPrice, extMeachulDate, extJungsanDate"
        			    sqlStr = sqlStr + " , extItemName, extItemOptionName, extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice,extItemID,extItemOption) "
        				sqlStr = sqlStr + " values('ssg', '" + CStr(ordNo) + "', '" + CStr(ordItemSeq)
        				sqlStr = sqlStr + "', '" + CStr(orordNo) + "', '" + CStr(salesQty) + "', '" + CStr(CLNG(salesAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(netAmt/(salesQty))) + "', '" + CStr(CLNG(owncoBdnDcAmt/(salesQty))) + "', '" + CStr(CLNG(splvenBdnDcAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr("C") + "', '" + CStr(CLNG(extCommPrice/(salesQty))) + "', '" + CStr(CLNG(salesAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(settleAmt/(salesQty))) + "', '" + CStr(settleDt) + "', '" + CStr("")
        				sqlStr = sqlStr + "', '" + CStr(itemNm) + "', convert(varchar(128),'" & CStr("") & "'), '" + CStr(extVatYN)
        				sqlStr = sqlStr + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "','"&itemId&"','"&uitemId&"') "
        				
        				dbget.execute sqlStr
    				end if
    				
    				if (isDlvPayExists) then ''��ۺ�
    				    itemNm = "��ۺ�"
    				    ordItemSeq = ordItemSeq+"-D"
    				    
    				    if (salesQty=0) then 
    				        salesQty = 1
    				        itemNm = "��ǰ��"
    				        ordItemSeq = ordItemSeq+"D"
    				    elseif (salesQty<0) then
    				        salesQty = -1
    				        itemNm = "��ǰ��"
    				    else
    				        salesQty = 1
    				        itemNm = "��ۺ�"
    				    end if
    				 
    				    
    				    salesAmt = dvShppcstTTL
    				    netAmt   = dvShppcstTTL
    				    settleAmt = dvShppcstTTL
    				    owncoBdnDcAmt = 0
    				    splvenBdnDcAmt = 0
    				    extCommPrice = 0
    				    uitemId = 0
    				    extVatYN = "Y"
    				    
    				    
    				    '' SSG�� ��ۺ� �Ⱥ��ϴµ� ������ ������ ������ ������.
    				    sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and extOrderserial='"&ordNo&"' and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"'" &vbCRLF
    				    rsget.CursorLocation = adUseClient
                        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                        if Not rsget.Eof then
                            if (rsget("CNT")>0) then
                                ordItemSeq = ordItemSeq + "-"&rsget("CNT")
                            end if
                        end if
                        rsget.Close
                        
    				    sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
        			    sqlStr = sqlStr + " (sellsite, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost"
        			    sqlStr = sqlStr + " , extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice"
        			    sqlStr = sqlStr + " , extTenJungsanPrice, extMeachulDate, extJungsanDate"
        			    sqlStr = sqlStr + " , extItemName, extItemOptionName, extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice,extItemID,extItemOption) "
        				sqlStr = sqlStr + " values('ssg', '" + CStr(ordNo) + "', '" + CStr(ordItemSeq)
        				sqlStr = sqlStr + "', '" + CStr(orordNo) + "', '" + CStr(salesQty) + "', '" + CStr(CLNG(salesAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(netAmt/salesQty)) + "', '" + CStr(CLNG(owncoBdnDcAmt/salesQty)) + "', '" + CStr(CLNG(splvenBdnDcAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr("D") + "', '" + CStr(CLNG(extCommPrice/(salesQty))) + "', '" + CStr(CLNG(salesAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(settleAmt/(salesQty))) + "', '" + CStr(settleDt) + "', '" + CStr("")
        				sqlStr = sqlStr + "', '" + CStr(itemNm) + "', convert(varchar(128),'" & CStr(uitemId) & "'), '" + CStr(extVatYN)
        				sqlStr = sqlStr + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "','"&itemId&"','"&uitemId&"') "
        				
        				dbget.execute sqlStr
    				    
    				end if
			    Next
		    
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
		
	''response.end
	    
	    if (isDataExists) then
	        sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ssg]"
	        'dbget.execute sqlStr
	    end if
	    
		if (FALSE) and (isDataExists) then
    		sqlStr = " update T"  
    		sqlStr = sqlStr + " SET extCommSupplyPrice=(CASE WHEN extVatYN='Y' THEN round(extCommPrice*10/11,0) ELSE extCommPrice END)"
    		sqlStr = sqlStr + " ,extCommSupplyVatPrice=extCommPrice-(CASE WHEN extVatYN='Y' THEN round(extCommPrice*10/11,0) ELSE extCommPrice END)"
    		sqlStr = sqlStr + " ,extTenMeachulSupplyPrice=(CASE WHEN extVatYN='Y' THEN round(extTenMeachulPrice*10/11,0) ELSE extTenMeachulPrice END)"
    		sqlStr = sqlStr + " ,extTenMeachulSupplyVatPrice=extTenMeachulPrice-(CASE WHEN extVatYN='Y' THEN round(extTenMeachulPrice*10/11,0) ELSE extTenMeachulPrice END)"
    		sqlStr = sqlStr + " from db_temp.dbo.tbl_xSite_JungsanTmp T"
    		sqlStr = sqlStr + " where sellsite='ssg'"
    		dbget.execute sqlStr
    		
    		'// ��Ī
    		sqlStr = " update T "
			sqlStr = sqlStr + " set T.OrgOrderserial = o.OrderSerial, T.itemid = o.matchItemID, T.itemoption = o.matchitemoption "
			sqlStr = sqlStr + " from "
			sqlStr = sqlStr + " 	db_temp.dbo.tbl_xSite_JungsanTmp T "
			sqlStr = sqlStr + " 	join db_temp.dbo.tbl_xSite_TMPOrder o "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and T.extOrderserial = o.OutMallOrderSerial "
			sqlStr = sqlStr + " 		and T.extOrderserSeq = o.OrgDetailKey "
			sqlStr = sqlStr + " 		and o.sellsite='ssg'"
			sqlStr = sqlStr + " 		and o.OrderSerial is Not NULL"
			sqlStr = sqlStr + " where T.sellsite = 'ssg' and T.extJungsanType='C' and T.itemid is NULL "  ''and T.extOrgOrderserial = '' 
			sqlStr = sqlStr + " and isNULL(T.orgorderserial,'')=''" ''��Ī �ȵȰ�
			dbget.execute sqlStr
			
			
    		'//��ۺ����
    		sqlStr = " update T "
    		sqlStr = sqlStr + " set "
    		sqlStr = sqlStr + " 	T.OrgOrderserial = o.OrderSerial "
    		sqlStr = sqlStr + " 	, T.itemid = 0 "
    		sqlStr = sqlStr + " 	, T.itemoption = (case "
    		sqlStr = sqlStr + " 						when T.extItemName = '��ǰ��' then '5000' "
    		sqlStr = sqlStr + " 						when T.extItemName <> '��ǰ��' and T.extOrgOrderserial = '' then '1000' "
    		sqlStr = sqlStr + " 						when T.extItemName <> '��ǰ��' and T.extOrgOrderserial <> '' then '5001' "
    		sqlStr = sqlStr + " 	end) "
    		sqlStr = sqlStr + " from "
    		sqlStr = sqlStr + " 	db_temp.dbo.tbl_xSite_JungsanTmp T "
    		sqlStr = sqlStr + " 	join db_temp.dbo.tbl_xSite_TMPOrder o "
    		sqlStr = sqlStr + " 	on "
    		sqlStr = sqlStr + " 		1 = 1 "
    		sqlStr = sqlStr + " 		and T.sellsite = o.sellsite "
    		sqlStr = sqlStr + " 		and ((T.extOrderserial = o.OutMallOrderSerial) or (T.extOrgOrderserial = o.OutMallOrderSerial)) "
    		sqlStr = sqlStr + " where "
    		sqlStr = sqlStr + " 	1 = 1 "
    		sqlStr = sqlStr + " 	and T.sellsite = 'ssg' and T.extJungsanType='D' and T.OrgOrderserial is NULL "
    		sqlStr = sqlStr + " 	and T.extItemName in ('��ۺ�', '��ǰ��') "
    		dbget.execute sqlStr
    		
    		
    		sqlStr = " insert into db_jungsan.dbo.tbl_xSite_JungsanData(sellsite, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost, extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice, extTenJungsanPrice, extMeachulDate, extJungsanDate, OrgOrderserial, itemid, itemoption, extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice) "
        	sqlStr = sqlStr + " select T.sellsite, T.extOrderserial, T.extOrderserSeq, T.extOrgOrderserial, T.extItemNo, T.extItemCost, T.extReducedPrice, T.extOwnCouponPrice, T.extTenCouponPrice, T.extJungsanType, T.extCommPrice, T.extTenMeachulPrice, T.extTenJungsanPrice, T.extMeachulDate, T.extJungsanDate, T.OrgOrderserial, T.itemid, T.itemoption, T.extVatYN, T.extCommSupplyPrice, T.extCommSupplyVatPrice, T.extTenMeachulSupplyPrice, T.extTenMeachulSupplyVatPrice "
        	sqlStr = sqlStr + "  from "
        	sqlStr = sqlStr + "  	db_temp.dbo.tbl_xSite_JungsanTmp T "
        	sqlStr = sqlStr + "  	left join db_jungsan.dbo.tbl_xSite_JungsanData j "
        	sqlStr = sqlStr + "  	on "
        	sqlStr = sqlStr + "  		1 = 1 "
        	sqlStr = sqlStr + "  		and T.sellsite = j.sellsite "
        	sqlStr = sqlStr + "  		and T.extOrderserial = j.extOrderserial "
        	sqlStr = sqlStr + "  		and T.extOrderserSeq = j.extOrderserSeq "
        	sqlStr = sqlStr + "  where 1=1 "
        	sqlStr = sqlStr + "     and T.sellsite='ssg'"
        	sqlStr = sqlStr + "  	and j.sellsite is NULL "
        	dbget.execute sqlStr
		end if
		
	SET objXML=Nothing
	
end function

''�Ͱ� ����� ��ȸ
public function getSsgDlvConfirmList()
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
    Dim ordpeHpno, ordMemoCntt, pCus, frebieNm ,shortgProgStatCd, shortgProgStatNm

    'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listWarehouseOut.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"
	
	requestBody = "<requestWarehouseOut>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>20171122</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>20171124</perdEndDts>"
    requestBody = requestBoDy&"</requestWarehouseOut>"

	objXML.send(requestBody)
	rw objXML.status
	
	response.write "<textarea cols=60 rows=30>"&objXML.responseText&"</textarea>"
    ''response.end

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			
			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			
			Set LagrgeNode = xmlDOM.SelectNodes("/result/warehouseOuts/warehouseOut")
			If Not (LagrgeNode Is Nothing) Then
			    For i = 0 To LagrgeNode.length - 1
			        shppNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''*��۹�ȣ
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*��ۼ���
                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''������ۻ���������ڵ�(��۴���) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���
                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''�̺�Ʈ����
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*��۱��л��ڵ� 11 �Ϲ���� 12 �κ���� 14 ���� 15 ��ȯ��� 16 AS���
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''��۱��л󼼸�
                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*�����ÿ��α��� 
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''����Ƚ�� 
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*�ֹ���ȣ [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*�ֹ�����
                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*�ֹ��Ϸ��Ͻ� [2017-11-23 10:39:42.0]
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''������ۻ�������¸�(��ۻ�ǰ����) [��ŷ�Ϸ�]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''������ۻ���������ڵ�(��ۻ�ǰ����) 11 ������� 21 ��ŷ���� 22 ��ŷ�Ϸ� 31 ��ŷ�Ϸ� 41 ����� 42 ������� 43 ���Ϸ� 51 ��ۿϷ� 52 ��۰���	
                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''���޾�ü���̵� [0000003198]
                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''���޾�ü��     
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''���������    [�ù���]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''��������ڵ� 10 �ڻ��� 20 �ù��� 30 ����湮 40 ��� 50 �̹�� 60 �̹߼�
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''����������ڵ� 14 ��ü�ڻ��� 22 ��ü�ù��� 25 �ؿ��ù��� 31 ����湮 41 ��� 51 SMS 52 EMAIL 61 �̹߼� 
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''��������󼼸� [��ü�ù���]
                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''�ù��ID [0000033011]
                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''�ڽ���ȣ [398327952]
                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' [303] ??
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*��ۺ� ���ҿ��� Y: ���� N: ����
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*��ǰ��
                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*��ü��ǰ��ȣ [1024019]
                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*��ǰ��ȣ [1000024811163]
                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*��ǰID [00000]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''���ü��� [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''��Ҽ��� [0]
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''�ֹ����� [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''�ǸŰ� [1000]
                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''����/�� ���� [����]
                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*�ֹ���

                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*������
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*������ �޴�����ȣ
                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*������ ����ȭ��ȣ
                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''������ ���ּ�
                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*������ �����ȣ          [04733]
                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*������ �������ȣ(6�ڸ�)  [133750]
                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''�����ε��θ��ּ�
                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''��ǰƯ�������ڵ� 10 �Ϲ� 20 ���θ� 30 �ؿܱ��Ŵ����ǰ 40 �̰����ͱݼ� 50 ����ϱ���Ʈ 60 ��ǰ�� 70 ���������� 80 ����ϻ�ǰ�� 91 �̺�Ʈ
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*��ۻ����ڵ� 10 ���� 30 ���
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''��ۻ��¸�
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''���ֹ���ȣ [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''���ֹ����� [2]
                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''�����ü�ڵ� 32 ��üâ�� 41 ���¾�ü 42 �귣������  [41]
                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''����Ʈ��ȣ 6001 �̸�Ʈ�� 6002 Ʈ���̴����� 6003 �н��� 6004 �ż���� 6005 S.COM�� 6009 �ż����ȭ����
                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''����Ʈ��
                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''���ް�
                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''�ǸźҰ���û���� 10:�Ϲ� 20: ���� GIFT �Ϲ� 30: ���� GIFT ���� 40: ���� GIFT ����
                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''�������ּ� 20170712
                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''�����λ��ּ�	20170712
                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''�ֹ���ǰ����	20170809
                    
                    ''//�ʼ��� �ƴѰ�� .
                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''�ֹ����޴�����ȣ  //���ð�
                    end if
                    
                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
                        ordMemoCntt         = LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text           ''����۸޸�  //���ð�
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
                    
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing	
			
	
end function

''������ø�� ��ȸ
public function getSsgDlvReqList()
    Dim objXML, xmlDOM, strSql
    Dim requestBody
    
    'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & ssgSSLAPIURL&"/api/pd/1/listShppDirection.ssg"
	objXML.setRequestHeader "Authorization", ssgApiKey
	objXML.setRequestHeader "Accept", "application/xml"  '' application/xml , application/json
	objXML.setRequestHeader "Content-Type", "application/xml"
	
	requestBody = "<requestShppDirection>"
    requestBody = requestBoDy&"<perdType>01</perdType>"
    requestBody = requestBoDy&"<perdStrDts>20171123</perdStrDts>"
    requestBody = requestBoDy&"<perdEndDts>20171123</perdEndDts>"
    requestBody = requestBoDy&"</requestShppDirection>"

	objXML.send(requestBody)
	rw objXML.status
	
	response.write "<textarea cols=60 rows=30>"&objXML.responseText&"</textarea>"
    response.end

	    Set xmlDOM = Server.CreateObject("MSXML.DOMDocument")
			xmlDOM.async = False
			xmlDOM.loadXML(objXML.responseText)
			
			ssgresultCode = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultMessage = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			ssgresultDesc = xmlDOM.getElementsByTagName("resultCode").Item(0).Text
			
			Set LagrgeNode = xmlDOM.SelectNodes("/result/shppDirections/shppDirection")
			If Not (LagrgeNode Is Nothing) Then
			    
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing	
			
	
end function
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->