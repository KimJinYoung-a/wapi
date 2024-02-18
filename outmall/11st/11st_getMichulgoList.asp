<%@ language=vbscript %>
<% option explicit %>
<!-- include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/11st/11stItemcls.asp"-->
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

function get11stDlvCode2Name(i11stcode)
    select Case i11stcode

        CASE "00034" : get11stDlvCode2Name = "CJ대한통운"    
        CASE "00011" : get11stDlvCode2Name = "한진택배"     
        CASE "00012" : get11stDlvCode2Name = "롯데(현대)택배"     
        CASE "00001" : get11stDlvCode2Name = "KGB택배"     
        CASE "00007" : get11stDlvCode2Name = "우체국택배"     
        CASE "00002" : get11stDlvCode2Name = "로젠택배"     
        CASE "00008" : get11stDlvCode2Name = "우편등기"     
        CASE "00021" : get11stDlvCode2Name = "대신택배"     
        CASE "00022" : get11stDlvCode2Name = "일양로지스"    
        CASE "00023" : get11stDlvCode2Name = "ACI"    
        CASE "00025" : get11stDlvCode2Name = "WIZWA"    
        CASE "00026" : get11stDlvCode2Name = "경동택배"    
        CASE "00027" : get11stDlvCode2Name = "천일택배"    
        CASE "00031" : get11stDlvCode2Name = "OCS Korea"   
        CASE "00035" : get11stDlvCode2Name = "합동택배"   
        CASE "00037" : get11stDlvCode2Name = "건영택배"   
        CASE "00099" : get11stDlvCode2Name = "기타"   
        CASE "00060" : get11stDlvCode2Name = "CVSnet편의점택배"   
        CASE "00061" : get11stDlvCode2Name = "CU편의점택배"   
        CASE "00062" : get11stDlvCode2Name = "호남택배"   
        CASE "00063" : get11stDlvCode2Name = "SLX택배"   
        CASE "00064" : get11stDlvCode2Name = "한의사랑택배"   
        CASE "00065" : get11stDlvCode2Name = "용마로지스"   
        CASE "00066" : get11stDlvCode2Name = "세방택배"   

        CASE "00067" : get11stDlvCode2Name = "농협택배"  
        CASE "00068" : get11stDlvCode2Name = "HI택배"  
        CASE "00069" : get11stDlvCode2Name = "원더스퀵"  
        CASE "00038" : get11stDlvCode2Name = "LG전자물류"  
        CASE "00036" : get11stDlvCode2Name = "삼성전자물류"  
        CASE "00039" : get11stDlvCode2Name = "DHL"  

        CASE  Else
            get11stDlvCode2Name = i11stcode
    end Select
end function

function getDevXNMLSAMPLE(iorderStatus)
    dim ret : ret =""

    if (iorderStatus="complete") then
        ret="<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
        ret=ret&"<ns2:orders xmlns:ns2=""http://skt.tmall.business.openapi.spring.service.client.domain/"">"
        ret=ret&"<ns2:result_code>0</ns2:result_code>"
        ret=ret&"<ns2:result_text>조회된 결과가 없습니다.</ns2:result_text>"
        ret=ret&"</ns2:orders>"
    elseif (iorderStatus="packaging") then
        ret="<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
        ret=ret&"<ns2:orders xmlns:ns2=""http://skt.tmall.business.openapi.spring.service.client.domain/"">"
        ret=ret&"<ns2:order>"
        ret=ret&"<addPrdNo>0</addPrdNo>"
        ret=ret&"<addPrdYn>N</addPrdYn>"
        ret=ret&"<appmtDdDlvDy></appmtDdDlvDy>"
        ret=ret&"<appmtEltRefuseYn></appmtEltRefuseYn>"
        ret=ret&"<appmtselStockCd></appmtselStockCd>"
        ret=ret&"<bmDlvCst>0</bmDlvCst>"
        ret=ret&"<bmDlvCstType>01</bmDlvCstType>"
        ret=ret&"<bndlDlvSeq>0</bndlDlvSeq>"
        ret=ret&"<bndlDlvYN>N</bndlDlvYN>"
        ret=ret&"<custGrdNm>일반고객</custGrdNm>"
        ret=ret&"<delaySendDt>"
        ret=ret&"</delaySendDt>"
        ret=ret&"<dlvCst>2500</dlvCst>"
        ret=ret&"<dlvCstType>01</dlvCstType>"
        ret=ret&"<dlvEtprsCd>"
        ret=ret&"</dlvEtprsCd>"
        ret=ret&"<dlvMthdCd>null</dlvMthdCd>"
        ret=ret&"<dlvNo>1420484760</dlvNo>"
        ret=ret&"<dlvSndDue>2020-01-08 00:00:00</dlvSndDue>"
        ret=ret&"<engNm>"
        ret=ret&"</engNm>"
        ret=ret&"<freeGiftNo>"
        ret=ret&"</freeGiftNo>"
        ret=ret&"<freeGiftQty>"
        ret=ret&"</freeGiftQty>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<giftCd>"
        ret=ret&"</giftCd>"
        ret=ret&"<invcNo>"
        ret=ret&"</invcNo>"
        ret=ret&"<lstDlvCst>2500</lstDlvCst>"
        ret=ret&"<lstSellerDscPrc>0</lstSellerDscPrc>"
        ret=ret&"<lstTmallDscPrc>2490</lstTmallDscPrc>"
        ret=ret&"<memID>xxxxxx</memID>"
        ret=ret&"<memNo>0000</memNo>"
        ret=ret&"<ordAmt>24900</ordAmt>"
        ret=ret&"<ordBaseAddr>서울특별시 종로구 통의동 xxxx</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>"
        ret=ret&"</ordDlvReqCont>"
        ret=ret&"<ordDt>2020-01-05 21:44:21</ordDt>"
        ret=ret&"<ordDtlsAddr>xxxxxxxx</ordDtlsAddr>"
        ret=ret&"<ordMailNo>null</ordMailNo>"
        ret=ret&"<ordNm>0000</ordNm>"
        ret=ret&"<ordNo>202001050486555</ordNo>"
        ret=ret&"<ordOptWonStl>0</ordOptWonStl>"
        ret=ret&"<ordPayAmt>24910</ordPayAmt>"
        ret=ret&"<ordPrdSeq>7</ordPrdSeq>"
        ret=ret&"<ordPrdStat>301</ordPrdStat>"
        ret=ret&"<ordPrtblTel>0000-0000-0000</ordPrtblTel>"
        ret=ret&"<ordQty>1</ordQty>"
        ret=ret&"<ordStlEndDt>2020-01-05 21:44:21</ordStlEndDt>"
        ret=ret&"<ordTlphnNo>000-000-0000</ordTlphnNo>"
        ret=ret&"<plcodrCnfDt>2020-01-06 08:23:18</plcodrCnfDt>"
        ret=ret&"<prdNm>[텐바이텐] Dot cherry bear-ipad pouch(아이패드 파우치)</prdNm>"
        ret=ret&"<prdNo>2627560696</prdNo>"
        ret=ret&"<prdStckNo>9783289854</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>경기도 안산시 단원구 </rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>000000</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>000000</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>"
        ret=ret&"</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>0000</rcvrNm>"
        ret=ret&"<rcvrPrtblNo>0000-0000-0000</rcvrPrtblNo>"
        ret=ret&"<rcvrTlphn>000-0000-0000</rcvrTlphn>"
        ret=ret&"<referSeq>"
        ret=ret&"</referSeq>"
        ret=ret&"<selPrc>24900</selPrc>"
        ret=ret&"<sellerDscPrc>0</sellerDscPrc>"
        ret=ret&"<sellerPrdCd>2583443</sellerPrdCd>"
        ret=ret&"<sellerStockCd>2583443_0010</sellerStockCd>"
        ret=ret&"<slctPrdOptNm>선택:FREE-1개</slctPrdOptNm>"
        ret=ret&"<sndEndDt>"
        ret=ret&"</sndEndDt>"
        ret=ret&"<tmallDscPrc>2490</tmallDscPrc>"
        ret=ret&"<typeAdd>02</typeAdd>"
        ret=ret&"<typeBilNo>4127310100106120000004114</typeBilNo>"
        ret=ret&"<visitDlvYn>N</visitDlvYn>"
        ret=ret&"</ns2:order>"
        ret=ret&"<ns2:order>"
        ret=ret&"<addPrdNo>0</addPrdNo>"
        ret=ret&"<addPrdYn>N</addPrdYn>"
        ret=ret&"<appmtDdDlvDy>"
        ret=ret&"</appmtDdDlvDy>"
        ret=ret&"<appmtEltRefuseYn>"
        ret=ret&"</appmtEltRefuseYn>"
        ret=ret&"<appmtselStockCd>"
        ret=ret&"</appmtselStockCd>"
        ret=ret&"<bmDlvCst>0</bmDlvCst>"
        ret=ret&"<bmDlvCstType>01</bmDlvCstType>"
        ret=ret&"<bndlDlvSeq>0</bndlDlvSeq>"
        ret=ret&"<bndlDlvYN>N</bndlDlvYN>"
        ret=ret&"<custGrdNm>일반고객</custGrdNm>"
        ret=ret&"<delaySendDt>"
        ret=ret&"</delaySendDt>"
        ret=ret&"<dlvCst>2500</dlvCst>"
        ret=ret&"<dlvCstType>01</dlvCstType>"
        ret=ret&"<dlvEtprsCd>"
        ret=ret&"</dlvEtprsCd>"
        ret=ret&"<dlvMthdCd>null</dlvMthdCd>"
        ret=ret&"<dlvNo>1420497378</dlvNo>"
        ret=ret&"<dlvSndDue>2020-01-08 00:00:00</dlvSndDue>"
        ret=ret&"<engNm>"
        ret=ret&"</engNm>"
        ret=ret&"<freeGiftNo>"
        ret=ret&"</freeGiftNo>"
        ret=ret&"<freeGiftQty>"
        ret=ret&"</freeGiftQty>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<giftCd>"
        ret=ret&"</giftCd>"
        ret=ret&"<invcNo>"
        ret=ret&"</invcNo>"
        ret=ret&"<lstDlvCst>2500</lstDlvCst>"
        ret=ret&"<lstSellerDscPrc>0</lstSellerDscPrc>"
        ret=ret&"<lstTmallDscPrc>3490</lstTmallDscPrc>"
        ret=ret&"<memID>00000</memID>"
        ret=ret&"<memNo>0000</memNo>"
        ret=ret&"<ordAmt>21670</ordAmt>"
        ret=ret&"<ordBaseAddr>경상남도 합천군 합천읍 000</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>"
        ret=ret&"</ordDlvReqCont>"
        ret=ret&"<ordDt>2020-01-05 22:10:57</ordDt>"
        ret=ret&"<ordDtlsAddr>709호</ordDtlsAddr>"
        ret=ret&"<ordMailNo>50237</ordMailNo>"
        ret=ret&"<ordNm>000</ordNm>"
        ret=ret&"<ordNo>202001050503265</ordNo>"
        ret=ret&"<ordOptWonStl>0</ordOptWonStl>"
        ret=ret&"<ordPayAmt>20680</ordPayAmt>"
        ret=ret&"<ordPrdSeq>1</ordPrdSeq>"
        ret=ret&"<ordPrdStat>301</ordPrdStat>"
        ret=ret&"<ordPrtblTel>000-0000-0000</ordPrtblTel>"
        ret=ret&"<ordQty>1</ordQty>"
        ret=ret&"<ordStlEndDt>2020-01-05 22:10:57</ordStlEndDt>"
        ret=ret&"<ordTlphnNo>000-0000-0000</ordTlphnNo>"
        ret=ret&"<plcodrCnfDt>2020-01-06 08:23:18</plcodrCnfDt>"
        ret=ret&"<prdNm>[텐바이텐] 돌핀웨일 플라잉날개니트슈트(60-100cm)</prdNm>"
        ret=ret&"<prdNo>2667099845</prdNo>"
        ret=ret&"<prdStckNo>9910872978</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>경상남도 합천군 합천읍 0</rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>000</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>50237</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>"
        ret=ret&"</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>000</rcvrNm>"
        ret=ret&"<rcvrPrtblNo>0000-0000-0000</rcvrPrtblNo>"
        ret=ret&"<rcvrTlphn>"
        ret=ret&"</rcvrTlphn>"
        ret=ret&"<referSeq>"
        ret=ret&"</referSeq>"
        ret=ret&"<selPrc>21670</selPrc>"
        ret=ret&"<sellerDscPrc>0</sellerDscPrc>"
        ret=ret&"<sellerPrdCd>2630664</sellerPrdCd>"
        ret=ret&"<sellerStockCd>2630664_0020</sellerStockCd>"
        ret=ret&"<slctPrdOptNm>색상/사이즈:아이보리/100cm-1개</slctPrdOptNm>"
        ret=ret&"<sndEndDt>"
        ret=ret&"</sndEndDt>"
        ret=ret&"<tmallDscPrc>3490</tmallDscPrc>"
        ret=ret&"<typeAdd>02</typeAdd>"
        ret=ret&"<typeBilNo>4889025021108690012014281</typeBilNo>"
        ret=ret&"<visitDlvYn>N</visitDlvYn>"
        ret=ret&"</ns2:order>"
        ret=ret&"</ns2:orders>"
    elseif (iorderStatus="delaydelivery/packagings") then
        ret="<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
        ret=ret&"<ns2:orders xmlns:ns2=""http://skt.tmall.business.openapi.spring.service.client.domain/"">"
        ret=ret&"<ns2:order>"
        ret=ret&"<addPrdNo>0</addPrdNo>"
        ret=ret&"<addPrdYn>N</addPrdYn>"
        ret=ret&"<appmtDdDlvDy>"
        ret=ret&"</appmtDdDlvDy>"
        ret=ret&"<appmtEltRefuseYn>"
        ret=ret&"</appmtEltRefuseYn>"
        ret=ret&"<appmtselStockCd>"
        ret=ret&"</appmtselStockCd>"
        ret=ret&"<bmDlvCst>0</bmDlvCst>"
        ret=ret&"<bmDlvCstType>01</bmDlvCstType>"
        ret=ret&"<bndlDlvSeq>0</bndlDlvSeq>"
        ret=ret&"<bndlDlvYN>N</bndlDlvYN>"
        ret=ret&"<delaySendDt>2019-12-25 00:00:00</delaySendDt>"
        ret=ret&"<dlvCst>0</dlvCst>"
        ret=ret&"<dlvCstType>03</dlvCstType>"
        ret=ret&"<dlvEtprsCd>"
        ret=ret&"</dlvEtprsCd>"
        ret=ret&"<dlvMthdCd>null</dlvMthdCd>"
        ret=ret&"<dlvNo>1392369104</dlvNo>"
        ret=ret&"<dlvSndDue>2019-12-25 00:00:00</dlvSndDue>"
        ret=ret&"<engNm>"
        ret=ret&"</engNm>"
        ret=ret&"<freeGiftNo>"
        ret=ret&"</freeGiftNo>"
        ret=ret&"<freeGiftQty>"
        ret=ret&"</freeGiftQty>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<giftCd>"
        ret=ret&"</giftCd>"
        ret=ret&"<invcNo>"
        ret=ret&"</invcNo>"
        ret=ret&"<lstDlvCst>0</lstDlvCst>"
        ret=ret&"<lstSellerDscPrc>0</lstSellerDscPrc>"
        ret=ret&"<lstTmallDscPrc>0</lstTmallDscPrc>"
        ret=ret&"<memID>xxxx</memID>"
        ret=ret&"<memNo>40842722</memNo>"
        ret=ret&"<ordAmt>765400</ordAmt>"
        ret=ret&"<ordBaseAddr>서울특별시 마포구 xx</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>"
        ret=ret&"</ordDlvReqCont>"
        ret=ret&"<ordDt>2019-11-14 15:59:16</ordDt>"
        ret=ret&"<ordDtlsAddr>xx-xx xx</ordDtlsAddr>"
        ret=ret&"<ordMailNo>04043</ordMailNo>"
        ret=ret&"<ordNm>xxx</ordNm>"
        ret=ret&"<ordNo>201911140444378</ordNo>"
        ret=ret&"<ordOptWonStl>0</ordOptWonStl>"
        ret=ret&"<ordPayAmt>765400</ordPayAmt>"
        ret=ret&"<ordPrdSeq>3</ordPrdSeq>"
        ret=ret&"<ordPrdStat>301</ordPrdStat>"
        ret=ret&"<ordPrtblTel>010-0000-0000</ordPrtblTel>"
        ret=ret&"<ordQty>10</ordQty>"
        ret=ret&"<ordStlEndDt>2019-11-14 16:29:32</ordStlEndDt>"
        ret=ret&"<ordTlphnNo>02-000-0000</ordTlphnNo>"
        ret=ret&"<plcodrCnfDt>2019-11-14 16:59:06</plcodrCnfDt>"
        ret=ret&"<prdNm>[텐바이텐] 코비 1등 인테리어벽등(B타입)</prdNm>"
        ret=ret&"<prdNo>2366108527</prdNo>"
        ret=ret&"<prdStckNo>8915846197</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>서울특별시 마포구 양화로 xx</rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>xx-xx</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>04043</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>"
        ret=ret&"</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>xx</rcvrNm>"
        ret=ret&"<rcvrPrtblNo>010-0000-0000</rcvrPrtblNo>"
        ret=ret&"<rcvrTlphn>00-000-0000</rcvrTlphn>"
        ret=ret&"<referSeq>"
        ret=ret&"</referSeq>"
        ret=ret&"<selPrc>76540</selPrc>"
        ret=ret&"<sellerDscPrc>0</sellerDscPrc>"
        ret=ret&"<sellerPrdCd>2305398</sellerPrdCd>"
        ret=ret&"<sellerStockCd>2305398_0011</sellerStockCd>"
        ret=ret&"<slctPrdOptNm>전구옵션 1등(26B):타입/전구 선택 안함-10개</slctPrdOptNm>"
        ret=ret&"<sndEndDt>"
        ret=ret&"</sndEndDt>"
        ret=ret&"<tmallDscPrc>0</tmallDscPrc>"
        ret=ret&"<typeAdd>02</typeAdd>"
        ret=ret&"<typeBilNo>1144012000103730009000001</typeBilNo>"
        ret=ret&"</ns2:order>"
        ret=ret&"</ns2:orders>"
    elseif (iorderStatus="dlvcompleted") then
        ret="<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
        ret=ret&"<ns2:orders xmlns:ns2=""http://skt.tmall.business.openapi.spring.service.client.domain/"">"
        ret=ret&"<ns2:order>"
        ret=ret&"<bmDlvCst>0</bmDlvCst>"
        ret=ret&"<bmDlvCstType>01</bmDlvCstType>"
        ret=ret&"<dlvCst>0</dlvCst>"
        ret=ret&"<dlvCstType>01</dlvCstType>"
        ret=ret&"<dlvEndDt>2019-12-31 15:35:00</dlvEndDt>"
        ret=ret&"<dlvEtprsCd>00034</dlvEtprsCd>"
        ret=ret&"<dlvMthd>택배</dlvMthd>"
        ret=ret&"<dlvMthdCd>01</dlvMthdCd>"
        ret=ret&"<dlvNo>1417396121</dlvNo>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<invcNo>627614443992</invcNo>"
        ret=ret&"<memID>xxxx</memID>"
        ret=ret&"<memNo>xxxx</memNo>"
        ret=ret&"<ordAmt>65000</ordAmt>"
        ret=ret&"<ordBaseAddr>"
        ret=ret&"</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>파손 위험이 있는 상품이니 조심히 다뤄주세요.</ordDlvReqCont>"
        ret=ret&"<ordDt>2019-11-20 23:41:11</ordDt>"
        ret=ret&"<ordDtlsAddr>"
        ret=ret&"</ordDtlsAddr>"
        ret=ret&"<ordId>nhanghee1</ordId>"
        ret=ret&"<ordMailNo>null</ordMailNo>"
        ret=ret&"<ordNm>xxx</ordNm>"
        ret=ret&"<ordNo>201911205844685</ordNo>"
        ret=ret&"<ordPrdSeq>1</ordPrdSeq>"
        ret=ret&"<ordPrdStat>501</ordPrdStat>"
        ret=ret&"<ordPrtblTel>000-000-000</ordPrtblTel>"
        ret=ret&"<ordQty>1</ordQty>"
        ret=ret&"<ordStlEndDt>2019-11-20 23:41:11</ordStlEndDt>"
        ret=ret&"<prdNm>[텐바이텐] 미음 써니 원형원목 거실 거실테이블 800</prdNm>"
        ret=ret&"<prdNo>1887862763</prdNo>"
        ret=ret&"<prdStckNo>7654810427</prdStckNo>"
        ret=ret&"<rcvrBaseAddr>서울특별시 관악구 </rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>xxx (xxx</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>08704</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>0</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>xxx</rcvrNm>"
        ret=ret&"<rcvrPrtblNo>000-000-000</rcvrPrtblNo>"
        ret=ret&"<rcvrTlphn>"
        ret=ret&"</rcvrTlphn>"
        ret=ret&"<referSeq>"
        ret=ret&"</referSeq>"
        ret=ret&"<selPrc>65000</selPrc>"
        ret=ret&"<sellerPrdCd>1804833</sellerPrdCd>"
        ret=ret&"<sellerStockCd>1804833_0011</sellerStockCd>"
        ret=ret&"<slctPrdOptNm>색상:화이트-1개</slctPrdOptNm>"
        ret=ret&"<sndEndDt>2019-12-30 12:27:03</sndEndDt>"
        ret=ret&"<typeBilNo>1162010200104620008014187</typeBilNo>"
        ret=ret&"</ns2:order>"
        ret=ret&"<ns2:order>"
        ret=ret&"<bmDlvCst>0</bmDlvCst>"
        ret=ret&"<bmDlvCstType>01</bmDlvCstType>"
        ret=ret&"<dlvCst>2500</dlvCst>"
        ret=ret&"<dlvCstType>01</dlvCstType>"
        ret=ret&"<dlvEndDt>null</dlvEndDt>"
        ret=ret&"<dlvEtprsCd>00012</dlvEtprsCd>"
        ret=ret&"<dlvMthd>택배</dlvMthd>"
        ret=ret&"<dlvMthdCd>01</dlvMthdCd>"
        ret=ret&"<dlvNo>1397825366</dlvNo>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<invcNo>230417660605</invcNo>"
        ret=ret&"<memID>cxxxxxx</memID>"
        ret=ret&"<memNo>17471736</memNo>"
        ret=ret&"<ordAmt>24900</ordAmt>"
        ret=ret&"<ordBaseAddr>서울특별시 </ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>xxxx</ordDlvReqCont>"
        ret=ret&"<ordDt>2019-11-22 13:24:08</ordDt>"
        ret=ret&"<ordDtlsAddr>xx xx</ordDtlsAddr>"
        ret=ret&"<ordId>xxx</ordId>"
        ret=ret&"<ordMailNo>06106</ordMailNo>"
        ret=ret&"<ordNm>xxx xxx</ordNm>"
        ret=ret&"<ordNo>201911226994211</ordNo>"
        ret=ret&"<ordPrdSeq>1</ordPrdSeq>"
        ret=ret&"<ordPrdStat>501</ordPrdStat>"
        ret=ret&"<ordPrtblTel>000-0000-0000</ordPrtblTel>"
        ret=ret&"<ordQty>1</ordQty>"
        ret=ret&"<ordStlEndDt>2019-11-22 13:24:08</ordStlEndDt>"
        ret=ret&"<prdNm>[텐바이텐] 월넛 원목 스탠드 옷걸이</prdNm>"
        ret=ret&"<prdNo>1765797936</prdNo>"
        ret=ret&"<prdStckNo>6770726036</prdStckNo>"
        ret=ret&"<rcvrBaseAddr>서울특별시 </rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>xxxx xxx</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>06106</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>"
        ret=ret&"</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>유근원</rcvrNm>"
        ret=ret&"<rcvrPrtblNo>000-0000-0000</rcvrPrtblNo>"
        ret=ret&"<rcvrTlphn>02-000-0000</rcvrTlphn>"
        ret=ret&"<referSeq>"
        ret=ret&"</referSeq>"
        ret=ret&"<selPrc>24900</selPrc>"
        ret=ret&"<sellerPrdCd>1551550</sellerPrdCd>"
        ret=ret&"<sellerStockCd>"
        ret=ret&"</sellerStockCd>"
        ret=ret&"<slctPrdOptNm>"
        ret=ret&"</slctPrdOptNm>"
        ret=ret&"<sndEndDt>2019-11-25 12:57:04</sndEndDt>"
        ret=ret&"<typeBilNo>1168010800102160013007809</typeBilNo>"
        ret=ret&"</ns2:order>"
        ret=ret&"</ns2:orders>"
    end if
    getDevXNMLSAMPLE = ret
end function 

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE 
 
Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-1,now()),10)

    if eddt<>"" then 
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)
'' 

'' 조회 유형 
'' 발주통보(:결제완료)(https://api.11st.co.kr/rest/ordservices/complete/[startTime]/[endTime])
'' 배송준비중        (https://api.11st.co.kr/rest/ordservices/packaging/[startTime]/[endTime])
''   발송기한경과 요청내역 (배송준비중 목록조회) https://api.11st.co.kr/rest/ordservices/delaydelivery/packagings
'' 배송완료          (https://api.11st.co.kr/rest/ordservices/dlvcompleted/[startTime]/[endTime])
'' 미도착조회        (https://api.11st.co.kr/rest/nondeliverys/nondeliverylist/[shDateType]/[shDateFrom]/[shDateTo])



' Dim lastconfirmDT
' sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_LastInputConfirmDT] 'interpark','"&confirmDt&"'"
' dbget.CursorLocation = adUseClient
' rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
' if NOT rsget.Eof then
'     lastconfirmDT = rsget("lastconfirmDT")
' end if
' rsget.close()

''----------------------------------------------------------------------------------------------------
sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] '11st1010','"&confirmDt&"'"
dbget.Execute sqlStr
rw "초기화작업"

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k

call Get_11st1010OrderListByStatus(LEFT(dateadd("d",-7,iedyyyymmdd),10),iedyyyymmdd,"complete","주문통보",0)  ''최대7일 가능
response.flush
call Get_11st1010OrderListByStatus(LEFT(dateadd("d",-7,iedyyyymmdd),10),iedyyyymmdd,"packaging","주문확인",0)
response.flush

for k=0 to datelen -1
    thedate=dateadd("d",1*k,istyyyymmdd)

    call Get_11st1010OrderListByStatus(thedate,thedate,"dlvcompleted","배송완료",1) ''최근 배송완료건만. dlvEndDt==null ??
    response.flush
    call Get_11st1010OrderListByStatus(thedate,thedate,"dlvcompleted","배송완료",2) ''최근 배송완료건만. dlvEndDt==null ??
next
 
iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","주문확인",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","주문확인",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","주문확인",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","주문확인",0)
response.flush


call Get_11st1010OrderListByStatus("","","delaydelivery/packagings","출고지연",0)


sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '11st1010','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_11st1010OrderListByStatus(stdate,eddate,iorderStatus,istatusName,ipartial)
	dim sellsite : sellsite = "11st1010"
	dim xmlURL, xmlSelldate
	dim objXML, xmlDOM, objData
	dim masterCnt, detailCnt, resultcode, obj
	dim objMasterListXML, SubNodes
	dim oMaster, oDetail, oDetailArr
	dim i, j, k
	dim tmpStr, pos
	dim successCnt : successCnt = 0
	dim strRst
	dim tmpOptionSeq : tmpOptionSeq = 0
	dim postParam
	dim tmpXML, oSql

    dim strSql, bufStr

	Get_11st1010OrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// 주문내역조회2 , 발주확인된 리스트
	xmlURL = APISSLURL&"/"
	xmlURL = xmlURL&"ordservices/"&iorderStatus
    if (stdate<>"") then
        if (ipartial=1) then
            xmlURL = xmlURL&"/" + replace(stdate,"-","") + "0000" + "/" + replace(eddate,"-","") + "1659"
        elseif (ipartial=2) then
            xmlURL = xmlURL&"/" + replace(stdate,"-","") + "1700" + "/" + replace(eddate,"-","") + "2359"
        else
            xmlURL = xmlURL&"/" + replace(stdate,"-","") + "0000" + "/" + replace(eddate,"-","") + "2359"
        end if
    end if

    if (ipartial=1) then
        rw "기간검색:"&stdate&"~"&eddate&" 16:59 상태:"&iorderStatus&"("&istatusName&")"
    elseif (ipartial=2) then
        rw "기간검색:"&stdate&" 17:00~"&eddate&" 23:59 상태:"&iorderStatus&"("&istatusName&")"
    else
        rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
    end if
	'// =======================================================================
	'// 데이타 가져오기


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
	objXML.setRequestHeader "openapikey",""&APIkey&""
    
    if (application("Svr_Info")<>"Dev") then
	    objXML.send()

        if objXML.Status <> "200" then
            response.write "ERROR : 통신오류" & objXML.Status
            dbget.close : response.end
        end if
    end if

	

    Dim iRbody, strObj, orderCount, obj1, obj2, obj3


    Dim ordNo, ordItemSeq, shppNo, shppSeq, reOrderYn, delayNts
    Dim cspGoodsCd, goodsCd, uitemId, orderQty, shppDivDtlNm
    Dim optionContent, shppRsvtDt, whoutCritnDt, autoShortgYn
    Dim orderStatus, dlvrCd, dlvrNo, dlvrDt, dlvrFinishDt, cancelDt
    Dim paramInfo, retParamInfo, RetErr

    Dim shppTypeDtlNm, delicoVenId, delicoVenNm, wblNo
	Dim invoiceUpDt, outjFixedDt
    Dim ORDCLM_STAT_DTS, DELV_DTS, DELV_COMPLETE_DT, dlvEndDt
    Dim resultNode, result_text

    if (application("Svr_Info")<>"Dev") then
	    iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
    else
        iRbody = getDevXNMLSAMPLE(iorderStatus)
    end if

if (request("showrow")<>"") then
    rw "<textarea cols=80 rows=20>"&iRbody&"</textarea>"
end if
' exit function

    Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML iRbody

    Set resultNode = xmlDOM.selectSingleNode("ns2:result_code")
        If NOT (resultNode Is Nothing)  Then
            result_text = xmlDOM.getElementsByTagName("ns2:result_text").item(0).text
        End If
    Set resultNode = nothing

    If result_text <> "" Then
        rw result_text
        Set xmlDOM = Nothing
        Set objXML = Nothing
        exit function
    end if
    

	set objMasterListXML = xmlDOM.getElementsByTagName("ns2:order")  ''selectNodes

	if objMasterListXML is Nothing then
		rw "No outPutValue"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if
    

    
	orderCount = objMasterListXML.length

	response.write "주문건수(" & orderCount & ") " & "<br />"

	For each SubNodes in objMasterListXML
        
        shppNo = Trim(SubNodes.getElementsByTagName("dlvNo").item(0).text)            '배송번호
        shppSeq	  = ""	'배송지시상세번호

        ordNo  = Trim(SubNodes.getElementsByTagName("ordNo").item(0).text)					'11번가 주문번호
        ordItemSeq = Trim(SubNodes.getElementsByTagName("ordPrdSeq").item(0).text)				'주문순번
        orderQty =  Trim(SubNodes.getElementsByTagName("ordQty").item(0).text)				'수량
        goodsCd = Trim(SubNodes.getElementsByTagName("prdNo").item(0).text)					'11번가상품번호
        uitemId = Trim(SubNodes.getElementsByTagName("prdStckNo").item(0).text)				'주문상품옵션코드
        cspGoodsCd = Trim(SubNodes.getElementsByTagName("sellerPrdCd").item(0).text)			'판매자상품번호
        optionContent = Trim(SubNodes.getElementsByTagName("slctPrdOptNm").item(0).text)			'주문상품옵션명
        optionContent = LEFT(optionContent,60)
        '' Trim(SubNodes.getElementsByTagName("sellerStockCd").item(0).text)			'판매자 재고번호

        whoutCritnDt = ""
        if (NOT (SubNodes.selectSingleNode("dlvSndDue") is Nothing)) then
            whoutCritnDt  = LEFT(Trim(SubNodes.getElementsByTagName("dlvSndDue").item(0).text),10)			'발송마감일자.  '' 발송기한.
        end if

        ORDCLM_STAT_DTS = "" : DELV_DTS="" : DELV_COMPLETE_DT ="" : dlvEndDt=""
        if (NOT (SubNodes.selectSingleNode("plcodrCnfDt") is Nothing)) then
            ORDCLM_STAT_DTS = Trim(SubNodes.getElementsByTagName("plcodrCnfDt").item(0).text) ''발주확인일시
        end if

        if (NOT (SubNodes.selectSingleNode("sndEndDt") is Nothing)) then
            DELV_DTS        = Trim(SubNodes.getElementsByTagName("sndEndDt").item(0).text)    ''발송처리일
        end if

        if (NOT (SubNodes.selectSingleNode("dlvEndDt") is Nothing)) then
            dlvEndDt        = Trim(SubNodes.getElementsByTagName("dlvEndDt").item(0).text)    ''배송완료일
        end if

        if (NOT (SubNodes.selectSingleNode("pocnfrmDt") is Nothing)) then
            DELV_COMPLETE_DT= SubNodes.selectSingleNode("pocnfrmDt").text         ''수취확인일
        end if

        invoiceUpDt = "" ''운송장번호 업로드 일시 (이게 오래된거면 추적(집하)이 안된거 일 수 있다.) = 출고일
        if (DELV_DTS<>"") then
            invoiceUpDt = DELV_DTS
        end if

        shppTypeDtlNm   = ""
        if (NOT (SubNodes.selectSingleNode("pocnfrmDt") is Nothing)) then
            shppTypeDtlNm = SubNodes.selectSingleNode("dlvMthdCd").text ''배송방식 01 : 택배, 04 : 우편(소포/등기),  05 : 직접전달(화물배달),  06 : 퀵서비스,  99 : 배송없음
            if (shppTypeDtlNm="01") then
                shppTypeDtlNm="택배"
            elseif (shppTypeDtlNm="04") then
                shppTypeDtlNm="우편"
            elseif (shppTypeDtlNm="05") then
                shppTypeDtlNm="직접전달"
            elseif (shppTypeDtlNm="06") then
                shppTypeDtlNm="퀵서비스"
            end if
        end if

        shppDivDtlNm = "일반출고"
        shppRsvtDt      = "" ''예정일
        autoShortgYn    = "" ''자동결품여부
        
        reOrderYn ="N" ''재주문여부 
        delayNts  =""  ''지연일수

        

        delicoVenId = "" : wblNo="" : delicoVenNm=""
        if (NOT (SubNodes.selectSingleNode("dlvEtprsCd") is Nothing)) then 
            delicoVenId     = SubNodes.selectSingleNode("dlvEtprsCd").text '택배배송사코드
        end if
        if (NOT (SubNodes.selectSingleNode("invcNo") is Nothing)) then 
            wblNo           = SubNodes.selectSingleNode("invcNo").text   '운송장번호
        end if

        delicoVenNm     = get11stDlvCode2Name(delicoVenId) '택배사명     

        orderStatus     = ""  
        if (ORDCLM_STAT_DTS<>"") then   
            orderStatus  = "주문확인" 
        end if 
        if (DELV_DTS<>"") then   
            orderStatus  = "출고완료" 
        end if 
        if (dlvEndDt<>"") and (dlvEndDt<>"null") then   '' null 이 있음 => 아직 정산 안된것? (먼가 CS건이 있는것 같음.)
            orderStatus  = "배송완료" 
        end if 
        if (DELV_COMPLETE_DT<>"") then   
            orderStatus  = "배송완료" ''구매확정.
        end if 
        
        if (istatusName="출고지연") then
            orderStatus="출고지연"
        end if

        outjFixedDt = ""
        outjFixedDt = DELV_COMPLETE_DT '' 


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
        bufStr = bufStr &"|"&invoiceUpDt
        bufStr = bufStr &"|"&outjFixedDt
        

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

        On Error RESUME Next
        retParamInfo = fnExecSPOutput(sqlStr, paramInfo)
          If ERR then
              rw invoiceUpDt 
              rw outjFixedDt
              response.end
          end if
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        
        successCnt = successCnt+1
        
    Next

    set objMasterListXML = Nothing
    Set xmlDOM = Nothing
	Set objXML = Nothing

    rw "상세건수:"&successCnt
    rw "======================================"

	Get_11st1010OrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->