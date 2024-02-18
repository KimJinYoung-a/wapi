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

        CASE "00034" : get11stDlvCode2Name = "CJ�������"    
        CASE "00011" : get11stDlvCode2Name = "�����ù�"     
        CASE "00012" : get11stDlvCode2Name = "�Ե�(����)�ù�"     
        CASE "00001" : get11stDlvCode2Name = "KGB�ù�"     
        CASE "00007" : get11stDlvCode2Name = "��ü���ù�"     
        CASE "00002" : get11stDlvCode2Name = "�����ù�"     
        CASE "00008" : get11stDlvCode2Name = "������"     
        CASE "00021" : get11stDlvCode2Name = "����ù�"     
        CASE "00022" : get11stDlvCode2Name = "�Ͼ������"    
        CASE "00023" : get11stDlvCode2Name = "ACI"    
        CASE "00025" : get11stDlvCode2Name = "WIZWA"    
        CASE "00026" : get11stDlvCode2Name = "�浿�ù�"    
        CASE "00027" : get11stDlvCode2Name = "õ���ù�"    
        CASE "00031" : get11stDlvCode2Name = "OCS Korea"   
        CASE "00035" : get11stDlvCode2Name = "�յ��ù�"   
        CASE "00037" : get11stDlvCode2Name = "�ǿ��ù�"   
        CASE "00099" : get11stDlvCode2Name = "��Ÿ"   
        CASE "00060" : get11stDlvCode2Name = "CVSnet�������ù�"   
        CASE "00061" : get11stDlvCode2Name = "CU�������ù�"   
        CASE "00062" : get11stDlvCode2Name = "ȣ���ù�"   
        CASE "00063" : get11stDlvCode2Name = "SLX�ù�"   
        CASE "00064" : get11stDlvCode2Name = "���ǻ���ù�"   
        CASE "00065" : get11stDlvCode2Name = "�븶������"   
        CASE "00066" : get11stDlvCode2Name = "�����ù�"   

        CASE "00067" : get11stDlvCode2Name = "�����ù�"  
        CASE "00068" : get11stDlvCode2Name = "HI�ù�"  
        CASE "00069" : get11stDlvCode2Name = "��������"  
        CASE "00038" : get11stDlvCode2Name = "LG���ڹ���"  
        CASE "00036" : get11stDlvCode2Name = "�Ｚ���ڹ���"  
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
        ret=ret&"<ns2:result_text>��ȸ�� ����� �����ϴ�.</ns2:result_text>"
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
        ret=ret&"<custGrdNm>�Ϲݰ�</custGrdNm>"
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
        ret=ret&"<ordBaseAddr>����Ư���� ���α� ���ǵ� xxxx</ordBaseAddr>"
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
        ret=ret&"<prdNm>[�ٹ�����] Dot cherry bear-ipad pouch(�����е� �Ŀ�ġ)</prdNm>"
        ret=ret&"<prdNo>2627560696</prdNo>"
        ret=ret&"<prdStckNo>9783289854</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>��⵵ �Ȼ�� �ܿ��� </rcvrBaseAddr>"
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
        ret=ret&"<slctPrdOptNm>����:FREE-1��</slctPrdOptNm>"
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
        ret=ret&"<custGrdNm>�Ϲݰ�</custGrdNm>"
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
        ret=ret&"<ordBaseAddr>��󳲵� ��õ�� ��õ�� 000</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>"
        ret=ret&"</ordDlvReqCont>"
        ret=ret&"<ordDt>2020-01-05 22:10:57</ordDt>"
        ret=ret&"<ordDtlsAddr>709ȣ</ordDtlsAddr>"
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
        ret=ret&"<prdNm>[�ٹ�����] ���ɿ��� �ö��׳�����Ʈ��Ʈ(60-100cm)</prdNm>"
        ret=ret&"<prdNo>2667099845</prdNo>"
        ret=ret&"<prdStckNo>9910872978</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>��󳲵� ��õ�� ��õ�� 0</rcvrBaseAddr>"
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
        ret=ret&"<slctPrdOptNm>����/������:���̺���/100cm-1��</slctPrdOptNm>"
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
        ret=ret&"<ordBaseAddr>����Ư���� ������ xx</ordBaseAddr>"
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
        ret=ret&"<prdNm>[�ٹ�����] �ں� 1�� ���׸����(BŸ��)</prdNm>"
        ret=ret&"<prdNo>2366108527</prdNo>"
        ret=ret&"<prdStckNo>8915846197</prdStckNo>"
        ret=ret&"<psnCscUniqNo>"
        ret=ret&"</psnCscUniqNo>"
        ret=ret&"<rcvrBaseAddr>����Ư���� ������ ��ȭ�� xx</rcvrBaseAddr>"
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
        ret=ret&"<slctPrdOptNm>�����ɼ� 1��(26B):Ÿ��/���� ���� ����-10��</slctPrdOptNm>"
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
        ret=ret&"<dlvMthd>�ù�</dlvMthd>"
        ret=ret&"<dlvMthdCd>01</dlvMthdCd>"
        ret=ret&"<dlvNo>1417396121</dlvNo>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<invcNo>627614443992</invcNo>"
        ret=ret&"<memID>xxxx</memID>"
        ret=ret&"<memNo>xxxx</memNo>"
        ret=ret&"<ordAmt>65000</ordAmt>"
        ret=ret&"<ordBaseAddr>"
        ret=ret&"</ordBaseAddr>"
        ret=ret&"<ordDlvReqCont>�ļ� ������ �ִ� ��ǰ�̴� ������ �ٷ��ּ���.</ordDlvReqCont>"
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
        ret=ret&"<prdNm>[�ٹ�����] ���� ��� �������� �Ž� �Ž����̺� 800</prdNm>"
        ret=ret&"<prdNo>1887862763</prdNo>"
        ret=ret&"<prdStckNo>7654810427</prdStckNo>"
        ret=ret&"<rcvrBaseAddr>����Ư���� ���Ǳ� </rcvrBaseAddr>"
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
        ret=ret&"<slctPrdOptNm>����:ȭ��Ʈ-1��</slctPrdOptNm>"
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
        ret=ret&"<dlvMthd>�ù�</dlvMthd>"
        ret=ret&"<dlvMthdCd>01</dlvMthdCd>"
        ret=ret&"<dlvNo>1397825366</dlvNo>"
        ret=ret&"<gblDlvYn>N</gblDlvYn>"
        ret=ret&"<invcNo>230417660605</invcNo>"
        ret=ret&"<memID>cxxxxxx</memID>"
        ret=ret&"<memNo>17471736</memNo>"
        ret=ret&"<ordAmt>24900</ordAmt>"
        ret=ret&"<ordBaseAddr>����Ư���� </ordBaseAddr>"
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
        ret=ret&"<prdNm>[�ٹ�����] ���� ���� ���ĵ� �ʰ���</prdNm>"
        ret=ret&"<prdNo>1765797936</prdNo>"
        ret=ret&"<prdStckNo>6770726036</prdStckNo>"
        ret=ret&"<rcvrBaseAddr>����Ư���� </rcvrBaseAddr>"
        ret=ret&"<rcvrDtlsAddr>xxxx xxx</rcvrDtlsAddr>"
        ret=ret&"<rcvrMailNo>06106</rcvrMailNo>"
        ret=ret&"<rcvrMailNoSeq>"
        ret=ret&"</rcvrMailNoSeq>"
        ret=ret&"<rcvrNm>���ٿ�</rcvrNm>"
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
    istyyyymmdd = LEFT(dateadd("d",-14,iedyyyymmdd),10)
'' 

'' ��ȸ ���� 
'' �����뺸(:�����Ϸ�)(https://api.11st.co.kr/rest/ordservices/complete/[startTime]/[endTime])
'' ����غ���        (https://api.11st.co.kr/rest/ordservices/packaging/[startTime]/[endTime])
''   �߼۱��Ѱ�� ��û���� (����غ��� �����ȸ) https://api.11st.co.kr/rest/ordservices/delaydelivery/packagings
'' ��ۿϷ�          (https://api.11st.co.kr/rest/ordservices/dlvcompleted/[startTime]/[endTime])
'' �̵�����ȸ        (https://api.11st.co.kr/rest/nondeliverys/nondeliverylist/[shDateType]/[shDateFrom]/[shDateTo])



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
rw "�ʱ�ȭ�۾�"

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k

call Get_11st1010OrderListByStatus(LEFT(dateadd("d",-7,iedyyyymmdd),10),iedyyyymmdd,"complete","�ֹ��뺸",0)  ''�ִ�7�� ����
response.flush
call Get_11st1010OrderListByStatus(LEFT(dateadd("d",-7,iedyyyymmdd),10),iedyyyymmdd,"packaging","�ֹ�Ȯ��",0)
response.flush

for k=0 to datelen -1
    thedate=dateadd("d",1*k,istyyyymmdd)

    call Get_11st1010OrderListByStatus(thedate,thedate,"dlvcompleted","��ۿϷ�",1) ''�ֱ� ��ۿϷ�Ǹ�. dlvEndDt==null ??
    response.flush
    call Get_11st1010OrderListByStatus(thedate,thedate,"dlvcompleted","��ۿϷ�",2) ''�ֱ� ��ۿϷ�Ǹ�. dlvEndDt==null ??
next
 
iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","�ֹ�Ȯ��",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","�ֹ�Ȯ��",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","�ֹ�Ȯ��",0)
response.flush

iedyyyymmdd = LEFT(dateadd("d",-1,istyyyymmdd),10)
istyyyymmdd = LEFT(dateadd("d",-7,iedyyyymmdd),10)
call Get_11st1010OrderListByStatus(istyyyymmdd,iedyyyymmdd,"packaging","�ֹ�Ȯ��",0)
response.flush


call Get_11st1010OrderListByStatus("","","delaydelivery/packagings","�������",0)


sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '11st1010','"&confirmDt&"'"
dbget.Execute sqlStr
rw "�ֹ�����"

rw "�Ϸ�"
'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

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
	'// ��¥����
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// �ֹ�������ȸ2 , ����Ȯ�ε� ����Ʈ
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
        rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" 16:59 ����:"&iorderStatus&"("&istatusName&")"
    elseif (ipartial=2) then
        rw "�Ⱓ�˻�:"&stdate&" 17:00~"&eddate&" 23:59 ����:"&iorderStatus&"("&istatusName&")"
    else
        rw "�Ⱓ�˻�:"&stdate&"~"&eddate&" ����:"&iorderStatus&"("&istatusName&")"
    end if
	'// =======================================================================
	'// ����Ÿ ��������


	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "text/xml; charset=euc-kr"
	objXML.setRequestHeader "openapikey",""&APIkey&""
    
    if (application("Svr_Info")<>"Dev") then
	    objXML.send()

        if objXML.Status <> "200" then
            response.write "ERROR : ��ſ���" & objXML.Status
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

	response.write "�ֹ��Ǽ�(" & orderCount & ") " & "<br />"

	For each SubNodes in objMasterListXML
        
        shppNo = Trim(SubNodes.getElementsByTagName("dlvNo").item(0).text)            '��۹�ȣ
        shppSeq	  = ""	'������û󼼹�ȣ

        ordNo  = Trim(SubNodes.getElementsByTagName("ordNo").item(0).text)					'11���� �ֹ���ȣ
        ordItemSeq = Trim(SubNodes.getElementsByTagName("ordPrdSeq").item(0).text)				'�ֹ�����
        orderQty =  Trim(SubNodes.getElementsByTagName("ordQty").item(0).text)				'����
        goodsCd = Trim(SubNodes.getElementsByTagName("prdNo").item(0).text)					'11������ǰ��ȣ
        uitemId = Trim(SubNodes.getElementsByTagName("prdStckNo").item(0).text)				'�ֹ���ǰ�ɼ��ڵ�
        cspGoodsCd = Trim(SubNodes.getElementsByTagName("sellerPrdCd").item(0).text)			'�Ǹ��ڻ�ǰ��ȣ
        optionContent = Trim(SubNodes.getElementsByTagName("slctPrdOptNm").item(0).text)			'�ֹ���ǰ�ɼǸ�
        optionContent = LEFT(optionContent,60)
        '' Trim(SubNodes.getElementsByTagName("sellerStockCd").item(0).text)			'�Ǹ��� ����ȣ

        whoutCritnDt = ""
        if (NOT (SubNodes.selectSingleNode("dlvSndDue") is Nothing)) then
            whoutCritnDt  = LEFT(Trim(SubNodes.getElementsByTagName("dlvSndDue").item(0).text),10)			'�߼۸�������.  '' �߼۱���.
        end if

        ORDCLM_STAT_DTS = "" : DELV_DTS="" : DELV_COMPLETE_DT ="" : dlvEndDt=""
        if (NOT (SubNodes.selectSingleNode("plcodrCnfDt") is Nothing)) then
            ORDCLM_STAT_DTS = Trim(SubNodes.getElementsByTagName("plcodrCnfDt").item(0).text) ''����Ȯ���Ͻ�
        end if

        if (NOT (SubNodes.selectSingleNode("sndEndDt") is Nothing)) then
            DELV_DTS        = Trim(SubNodes.getElementsByTagName("sndEndDt").item(0).text)    ''�߼�ó����
        end if

        if (NOT (SubNodes.selectSingleNode("dlvEndDt") is Nothing)) then
            dlvEndDt        = Trim(SubNodes.getElementsByTagName("dlvEndDt").item(0).text)    ''��ۿϷ���
        end if

        if (NOT (SubNodes.selectSingleNode("pocnfrmDt") is Nothing)) then
            DELV_COMPLETE_DT= SubNodes.selectSingleNode("pocnfrmDt").text         ''����Ȯ����
        end if

        invoiceUpDt = "" ''������ȣ ���ε� �Ͻ� (�̰� �����ȰŸ� ����(����)�� �ȵȰ� �� �� �ִ�.) = �����
        if (DELV_DTS<>"") then
            invoiceUpDt = DELV_DTS
        end if

        shppTypeDtlNm   = ""
        if (NOT (SubNodes.selectSingleNode("pocnfrmDt") is Nothing)) then
            shppTypeDtlNm = SubNodes.selectSingleNode("dlvMthdCd").text ''��۹�� 01 : �ù�, 04 : ����(����/���),  05 : ��������(ȭ�����),  06 : ������,  99 : ��۾���
            if (shppTypeDtlNm="01") then
                shppTypeDtlNm="�ù�"
            elseif (shppTypeDtlNm="04") then
                shppTypeDtlNm="����"
            elseif (shppTypeDtlNm="05") then
                shppTypeDtlNm="��������"
            elseif (shppTypeDtlNm="06") then
                shppTypeDtlNm="������"
            end if
        end if

        shppDivDtlNm = "�Ϲ����"
        shppRsvtDt      = "" ''������
        autoShortgYn    = "" ''�ڵ���ǰ����
        
        reOrderYn ="N" ''���ֹ����� 
        delayNts  =""  ''�����ϼ�

        

        delicoVenId = "" : wblNo="" : delicoVenNm=""
        if (NOT (SubNodes.selectSingleNode("dlvEtprsCd") is Nothing)) then 
            delicoVenId     = SubNodes.selectSingleNode("dlvEtprsCd").text '�ù��ۻ��ڵ�
        end if
        if (NOT (SubNodes.selectSingleNode("invcNo") is Nothing)) then 
            wblNo           = SubNodes.selectSingleNode("invcNo").text   '������ȣ
        end if

        delicoVenNm     = get11stDlvCode2Name(delicoVenId) '�ù���     

        orderStatus     = ""  
        if (ORDCLM_STAT_DTS<>"") then   
            orderStatus  = "�ֹ�Ȯ��" 
        end if 
        if (DELV_DTS<>"") then   
            orderStatus  = "���Ϸ�" 
        end if 
        if (dlvEndDt<>"") and (dlvEndDt<>"null") then   '' null �� ���� => ���� ���� �ȵȰ�? (�հ� CS���� �ִ°� ����.)
            orderStatus  = "��ۿϷ�" 
        end if 
        if (DELV_COMPLETE_DT<>"") then   
            orderStatus  = "��ۿϷ�" ''����Ȯ��.
        end if 
        
        if (istatusName="�������") then
            orderStatus="�������"
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
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
        
        successCnt = successCnt+1
        
    Next

    set objMasterListXML = Nothing
    Set xmlDOM = Nothing
	Set objXML = Nothing

    rw "�󼼰Ǽ�:"&successCnt
    rw "======================================"

	Get_11st1010OrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->