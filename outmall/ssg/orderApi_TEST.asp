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
'' 1. 배송지시목록조회
'' 2. 주문 확인 처리
'' 3. 촐고 대상목록 조회

'CONST ssgAPIURL = "http://eapi.ssgadm.com"
'CONST ssgSSLAPIURL = "https://eapi.ssgadm.com"
'CONST ssgApiKey = "18a8d870-12a7-4b36-afaf-1e9d38e2b988"

''call getSsgDlvReqList()
''call getSsgDlvConfirmList()

''call getSsgJungsanList()


''정산리스트
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

''정산목록  조회
public function getSsgJungsanList(yyyymmdd,isiteno)
    Dim objXML, xmlDOM, strSql, i
    Dim requestBody
    Dim LagrgeNode
    
    Dim urlparam : urlparam="critnDt="&yyyymmdd&"&siteNo="&isiteno  '6006/6007
    Dim sqlStr
    
    'On Error Resume Next
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.open "POST", "" & ssgAPIURL&"/se/alln/getVendorSalesList.ssg"&"?"&urlparam    '''POST 이나 getParam?
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
			    response.write "<br>건수:" & LagrgeNode.length&":"&"<br>"
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
                    
                    siteNo           = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''6006: 이마트 / 6007:신세계
    			    settleDt         = LagrgeNode(i).SelectSingleNode("settleDt").Text               ''*정산 매출 일자	
    			    orordNo          = LagrgeNode(i).SelectSingleNode("orordNo").Text                ''원주문번호
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*주문번호 [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*주문순번? [1]
    			    
    			    txnDivNm     = LagrgeNode(i).SelectSingleNode("txnDivNm").Text                ''과세구분코드명 과세/면세
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''상품번호  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''상품명    [문주란스티커]
    			    uitemId      = LagrgeNode(i).SelectSingleNode("uitemId").Text                   '' 단품 ID	
    			    
    			    salesQty       = LagrgeNode(i).SelectSingleNode("salesQty").Text              ''    매출 수량
    			    settleAmt      = LagrgeNode(i).SelectSingleNode("settleAmt").Text            '' 정산 금액(vat 제외)
    			    salesAmt      = LagrgeNode(i).SelectSingleNode("salesAmt").Text            '' 매출 금액(vat 포함)
    			    splvenBdnDcAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnDcAmt").Text            ''업체 할인 부담 금액	
    			    owncoBdnDcAmt    = LagrgeNode(i).SelectSingleNode("owncoBdnDcAmt").Text            ''자사(SSG) 할인 부담 금액
    			    netAmt    = LagrgeNode(i).SelectSingleNode("netAmt").Text            ''순 판매 금액(vat 포함)
    			    mrgrt    = LagrgeNode(i).SelectSingleNode("mrgrt").Text            ''마진율
    			    dvShppcstAmt    = LagrgeNode(i).SelectSingleNode("dvShppcstAmt").Text            ''배송 금액(vat 제외)	
    			    dvShppcstVat    = LagrgeNode(i).SelectSingleNode("dvShppcstVat").Text            ''배송 VAT 금액
    			    custBdnShppAmt    = LagrgeNode(i).SelectSingleNode("custBdnShppAmt").Text            ''고객 부담 배송 금액
    			    splvenBdnShppAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnShppAmt").Text            ''업체 부담 배송 금액
    			    
    			    dvShppcstTTL = CLNG(dvShppcstAmt)+CLNG(dvShppcstVat)
    			    isDlvPayExists = (dvShppcstTTL<>0)
    			    extVatYN = "Y"
                    if (txnDivNm="면세") then extVatYN="N"
                    
                    if (extVatYN = "Y") then
    			        settleAmt = CLNG(settleAmt*1.1)  ''2018/03/06 추가
    			    end if
    			    
    			    if (isDlvPayExists) then
    			        salesAmt = salesAmt - dvShppcstTTL
    			        settleAmt = settleAmt - dvShppcstTTL
    			        netAmt = netAmt - dvShppcstTTL
    			    end if
    			    
    			    
    			    if (salesQty<0) then  ''반품인경우.
    			        ordNo = ordNo&"-"&ordItemSeq
    			    end if
    			    
    			    'if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			    '    ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")                 ''[[고객배송메모]배송메세지]
    			    'end if
    			    
    			    response.write siteNo&"|"&settleDt&"|"&orordNo&"|"&ordNo&"|"&ordItemSeq&"|"&txnDivNm&"|"&itemId&"|"&itemNm&"|"&uitemId&"|"
    			    response.write salesQty&"|"&settleAmt&"|"&salesAmt&"|"&splvenBdnDcAmt&"|"&owncoBdnDcAmt&"|"&netAmt&"|"&mrgrt&"|"&dvShppcstAmt&"|"&dvShppcstVat&"|"&custBdnShppAmt&"|"&custBdnShppAmt&"|"&splvenBdnShppAmt
    			    response.write "<br>"

                    ''extJungsanType : C-상품, D배송비
                    '' extTenMeachulPrice=extCommPrice+extTenJungsanPrice
                    '' extItemCost=extReducedPrice+extOwnCouponPrice+extTenCouponPrice
                    
                    extCommPrice = salesAmt-settleAmt
                    
                    if (orordNo=ordNo) then orordNo=""
                        
                    if (isDlvPayExists and (salesQty=0) ) then
                        '' 반품배송비
                    elseif ((NOT isDlvPayExists) and (salesQty=0) and (salesAmt=0)and (settleAmt=0))then
                        '' 재낌.
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
    				
    				if (isDlvPayExists) then ''배송비
    				    itemNm = "배송비"
    				    ordItemSeq = ordItemSeq+"-D"
    				    
    				    if (salesQty=0) then 
    				        salesQty = 1
    				        itemNm = "반품비"
    				        ordItemSeq = ordItemSeq+"D"
    				    elseif (salesQty<0) then
    				        salesQty = -1
    				        itemNm = "반품비"
    				    else
    				        salesQty = 1
    				        itemNm = "배송비"
    				    end if
    				 
    				    
    				    salesAmt = dvShppcstTTL
    				    netAmt   = dvShppcstTTL
    				    settleAmt = dvShppcstTTL
    				    owncoBdnDcAmt = 0
    				    splvenBdnDcAmt = 0
    				    extCommPrice = 0
    				    uitemId = 0
    				    extVatYN = "Y"
    				    
    				    
    				    '' SSG는 배송비도 안분하는듯 동일한 내역이 있으면 더하자.
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
    		
    		'// 매칭
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
			sqlStr = sqlStr + " and isNULL(T.orgorderserial,'')=''" ''매칭 안된것
			dbget.execute sqlStr
			
			
    		'//배송비매핑
    		sqlStr = " update T "
    		sqlStr = sqlStr + " set "
    		sqlStr = sqlStr + " 	T.OrgOrderserial = o.OrderSerial "
    		sqlStr = sqlStr + " 	, T.itemid = 0 "
    		sqlStr = sqlStr + " 	, T.itemoption = (case "
    		sqlStr = sqlStr + " 						when T.extItemName = '반품비' then '5000' "
    		sqlStr = sqlStr + " 						when T.extItemName <> '반품비' and T.extOrgOrderserial = '' then '1000' "
    		sqlStr = sqlStr + " 						when T.extItemName <> '반품비' and T.extOrgOrderserial <> '' then '5001' "
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
    		sqlStr = sqlStr + " 	and T.extItemName in ('배송비', '반품비') "
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

''촐고 대상목록 조회
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
			        shppNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text                 ''*배송번호
                    shppSeq             = LagrgeNode(i).SelectSingleNode("shppSeq").Text                ''*배송순번
                    shppTabProgStatCd   = LagrgeNode(i).SelectSingleNode("shppTabProgStatCd").Text      ''최종배송상세진행상태코드(배송단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절
                    evntSeq             = LagrgeNode(i).SelectSingleNode("evntSeq").Text                ''이벤트순번
                    shppDivDtlCd        = LagrgeNode(i).SelectSingleNode("shppDivDtlCd").Text           ''*배송구분상세코드 11 일반출고 12 부분출고 14 재배송 15 교환출고 16 AS출고
                    shppDivDtlNm        = LagrgeNode(i).SelectSingleNode("shppDivDtlNm").Text           ''배송구분상세명
                    reOrderYn           = LagrgeNode(i).SelectSingleNode("reOrderYn").Text              ''*재지시여부구분 
                    delayNts            = LagrgeNode(i).SelectSingleNode("delayNts").Text               ''지연횟수 
                    ordNo               = LagrgeNode(i).SelectSingleNode("ordNo").Text                  ''*주문번호 [20171123128379]
                    ordItemSeq          = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text             ''*주문순번
                    ordCmplDts          = LagrgeNode(i).SelectSingleNode("ordCmplDts").Text             ''*주문완료일시 [2017-11-23 10:39:42.0]
                    lastShppProgStatDtlNm   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlNm").Text  ''최종배송상세진행상태명(배송상품단위) [피킹완료]
                    lastShppProgStatDtlCd   = LagrgeNode(i).SelectSingleNode("lastShppProgStatDtlCd").Text  ''최종배송상세진행상태코드(배송상품단위) 11 배송지시 21 피킹지시 22 피킹완료 31 패킹완료 41 출고보류 42 출고지연 43 출고완료 51 배송완료 52 배송거절	
                    salestrNo           = LagrgeNode(i).SelectSingleNode("salestrNo").Text              '' [6004]
                    shppVenId           = LagrgeNode(i).SelectSingleNode("shppVenId").Text      ''공급업체아이디 [0000003198]
                    shppVenNm           = LagrgeNode(i).SelectSingleNode("shppVenNm").Text      ''공급업체명     
                    shppTypeNm          = LagrgeNode(i).SelectSingleNode("shppTypeNm").Text     ''배송유형명    [택배배송]
                    shppTypeCd          = LagrgeNode(i).SelectSingleNode("shppTypeCd").Text     ''배송유형코드 10 자사배송 20 택배배송 30 매장방문 40 등기 50 미배송 60 미발송
                    shppTypeDtlCd       = LagrgeNode(i).SelectSingleNode("shppTypeDtlCd").Text  ''배송유형상세코드 14 업체자사배송 22 업체택배배송 25 해외택배배송 31 매장방문 41 등기 51 SMS 52 EMAIL 61 미발송 
                    shppTypeDtlNm       = LagrgeNode(i).SelectSingleNode("shppTypeDtlNm").Text  ''배송유형상세명 [업체택배배송]
                    delicoVenId         = LagrgeNode(i).SelectSingleNode("delicoVenId").Text    ''택배사ID [0000033011]
                    boxNo               = LagrgeNode(i).SelectSingleNode("boxNo").Text          ''박스번호 [398327952]
                    shppcst             = LagrgeNode(i).SelectSingleNode("shppcst").Text        '' [303] ??
                    shppcstCodYn        = LagrgeNode(i).SelectSingleNode("shppcstCodYn").Text   ''*배송비 착불여부 Y: 착불 N: 선불
                    itemNm              = LagrgeNode(i).SelectSingleNode("itemNm").Text         ''*상품명
                    splVenItemId        = LagrgeNode(i).SelectSingleNode("splVenItemId").Text       ''*업체상품번호 [1024019]
                    itemId              = LagrgeNode(i).SelectSingleNode("itemId").Text             ''*상품번호 [1000024811163]
                    uitemId             = LagrgeNode(i).SelectSingleNode("uitemId").Text            ''*단품ID [00000]
                    dircItemQty         = LagrgeNode(i).SelectSingleNode("dircItemQty").Text        ''지시수량 [2]
                    cnclItemQty         = LagrgeNode(i).SelectSingleNode("cnclItemQty").Text        ''취소수량 [0]
                    ordQty              = LagrgeNode(i).SelectSingleNode("ordQty").Text             ''주문수량 [2]
                    sellprc             = LagrgeNode(i).SelectSingleNode("sellprc").Text            ''판매가 [1000]
                    frgShppYn           = LagrgeNode(i).SelectSingleNode("frgShppYn").Text          ''국내/외 구분 [국내]
                    ordpeNm             = LagrgeNode(i).SelectSingleNode("ordpeNm").Text            ''*주문자

                    rcptpeNm            = LagrgeNode(i).SelectSingleNode("rcptpeNm").Text           ''*수취인
                    rcptpeHpno          = LagrgeNode(i).SelectSingleNode("rcptpeHpno").Text         ''*수취인 휴대폰번호
                    rcptpeTelno         = LagrgeNode(i).SelectSingleNode("rcptpeTelno").Text        ''*수취인 집전화번호
                    shpplocAddr         = LagrgeNode(i).SelectSingleNode("shpplocAddr").Text        ''수취인 상세주소
                    shpplocZipcd        = LagrgeNode(i).SelectSingleNode("shpplocZipcd").Text       ''*수취인 우편번호          [04733]
                    shpplocOldZipcd     = LagrgeNode(i).SelectSingleNode("shpplocOldZipcd").Text    ''*수취인 구우편번호(6자리)  [133750]
                    shpplocRoadAddr     = LagrgeNode(i).SelectSingleNode("shpplocRoadAddr").Text    ''수취인도로명주소
                    itemChrctDivCd      = LagrgeNode(i).SelectSingleNode("itemChrctDivCd").Text     ''상품특성구분코드 10 일반 20 몰인몰 30 해외구매대행상품 40 미가공귀금속 50 모바일기프트 60 상품권 70 쇼핑충전금 80 모바일상품권 91 이벤트
                    shppStatCd          = LagrgeNode(i).SelectSingleNode("shppStatCd").Text         ''*배송상태코드 10 정상 30 대기
                    shppStatNm          = LagrgeNode(i).SelectSingleNode("shppStatNm").Text         ''배송상태명
                    orordNo             = LagrgeNode(i).SelectSingleNode("orordNo").Text            ''원주문번호 [20171123128379]
                    orordItemSeq        = LagrgeNode(i).SelectSingleNode("orordItemSeq").Text       ''원주문순번 [2]
                    shppMainCd          = LagrgeNode(i).SelectSingleNode("shppMainCd").Text         ''배송주체코드 32 업체창고 41 협력업체 42 브랜드직배  [41]
                    siteNo              = LagrgeNode(i).SelectSingleNode("siteNo").Text             ''사이트번호 6001 이마트몰 6002 트레이더스몰 6003 분스몰 6004 신세계몰 6005 S.COM몰 6009 신세계백화점몰
                    siteNm              = LagrgeNode(i).SelectSingleNode("siteNm").Text             ''사이트명
                    shppRsvtDt          = LagrgeNode(i).SelectSingleNode("shppRsvtDt").Text
                    splprc              = LagrgeNode(i).SelectSingleNode("splprc").Text             ''공급가
                    shortgYn            = LagrgeNode(i).SelectSingleNode("shortgYn").Text
                    newWblNoData        = LagrgeNode(i).SelectSingleNode("newWblNoData").Text
                    newRow              = LagrgeNode(i).SelectSingleNode("newRow").Text
                    itemDiv             = LagrgeNode(i).SelectSingleNode("itemDiv").Text                ''판매불가신청상태 10:일반 20: 명절 GIFT 일반 30: 명절 GIFT 센터 40: 명절 GIFT 냉장
                    shpplocBascAddr     = LagrgeNode(i).SelectSingleNode("shpplocBascAddr").Text        ''수취인주소 20170712
                    shpplocDtlAddr      = LagrgeNode(i).SelectSingleNode("shpplocDtlAddr").Text         ''수취인상세주소	20170712
                    ordItemDivNm        = LagrgeNode(i).SelectSingleNode("ordItemDivNm").Text           ''주문상품구분	20170809
                    
                    ''//필수값 아닌경우 .
                    if NOT (LagrgeNode(i).SelectSingleNode("ordpeHpno") is Nothing) then
                        ordpeHpno         = LagrgeNode(i).SelectSingleNode("ordpeHpno").Text           ''주문자휴대폰번호  //선택값
                    end if
                    
                    if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
                        ordMemoCntt         = LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text           ''고객배송메모  //선택값
                    end if
                    
                    if NOT (LagrgeNode(i).SelectSingleNode("pCus") is Nothing) then
                        pCus         = LagrgeNode(i).SelectSingleNode("pCus").Text           ''개인통관고유번호  //선택값
                    end if
                    
                    if NOT (LagrgeNode(i).SelectSingleNode("frebieNm") is Nothing) then
                        frebieNm         = LagrgeNode(i).SelectSingleNode("frebieNm").Text    ''사은품  //선택값
                    end if
                    
                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatCd") is Nothing) then
                        shortgProgStatCd         = LagrgeNode(i).SelectSingleNode("shortgProgStatCd").Text    ''판매불가신청상태  //선택값 11 결품등록 12 결품CS처리중 13 결품확정 21 상품정보오류등록 22 상품정보오류CS처리중 23 상품정보오류확정 41 입고지연등록 43 입고지연완료 51 배송지연등록 
                    end if

                    if NOT (LagrgeNode(i).SelectSingleNode("shortgProgStatNm") is Nothing) then
                        shortgProgStatNm         = LagrgeNode(i).SelectSingleNode("shortgProgStatNm").Text    ''결품진행상태명  //선택값
                    end if
                    
			    Next
			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing
	Set objXML = nothing	
			
	
end function

''배송지시목록 조회
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