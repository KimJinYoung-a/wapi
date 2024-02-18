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

function getiparkDlvCode2Name(iiparkcode)
    select Case iiparkcode
        CASE "169178" : getiparkDlvCode2Name = "한진"     ''
        CASE "169198" : getiparkDlvCode2Name = "롯데"     ''
        CASE "169177" : getiparkDlvCode2Name = "CJ GLS"     ''
        CASE "169168" : getiparkDlvCode2Name = "CJ GLS"     '''
        CASE "169199" : getiparkDlvCode2Name = "우체국택배"     ''
        CASE "169187" : getiparkDlvCode2Name = "KGB택배"     ''
        CASE "169194" : getiparkDlvCode2Name = "아주택배"     '' / 로엑스(구 아주)
        CASE "169200" : getiparkDlvCode2Name = "옐로우캡"     ''
        CASE "169182" : getiparkDlvCode2Name = "로젠택배"     ''
        CASE "303978" : getiparkDlvCode2Name = "경동택배"     ''
        CASE "169526" : getiparkDlvCode2Name = "고려택배"     ''
        CASE "236288" : getiparkDlvCode2Name = "쎄덱스택배"     '' 신세계
        CASE "229381" : getiparkDlvCode2Name = "하나로택배"     ''
        CASE "263792" : getiparkDlvCode2Name = "일양택배"     ''
        CASE "169194" : getiparkDlvCode2Name = "LOEX택배"     ''
        CASE "231145" : getiparkDlvCode2Name = "동부익스프레스"     ''
        CASE "231194" : getiparkDlvCode2Name = "건영택배"     ''
        CASE "266237" : getiparkDlvCode2Name = "이노지스"     ''
        CASE "230175" : getiparkDlvCode2Name = "천일택배"     ''
        CASE "250701" : getiparkDlvCode2Name = "호남택배"     ''
        CASE "258064" : getiparkDlvCode2Name = "대신택배"     ''
        CASE "169172" : getiparkDlvCode2Name = "CVSnet택배"     ''
        CASE "2641054" : getiparkDlvCode2Name = "합동택배"     ''
        CASE "2964976" : getiparkDlvCode2Name = "드림택배"     ''(동부택배,옐로우캡)  ''2018/02/13
		CASE "169177" : getiparkDlvCode2Name = "대한통운"     ''CU Post => 
        CASE "169316" : getiparkDlvCode2Name = "직배송"     ''퀵서비스->직배송
        CASE "169167" : getiparkDlvCode2Name = "기타"     ''
        CASE  Else
            getiparkDlvCode2Name = iiparkcode
    end Select
end function

function getDevXNMLSAMPLE0()
    dim ret : ret =""
    ret="<?xml version=""1.0"" encoding=""euc-kr""?>"
    ret=ret&"<ORDER_LIST>"
    ret=ret&"    <ORDER ID=""1"">"
    ret=ret&"        <ORD_ENM/>"
    ret=ret&"        <DELI_ADDR1_DORO>경기도 성남시 중원구 성남대로 xxxx (성남동,성남 메트로칸)</DELI_ADDR1_DORO>"
    ret=ret&"        <ORDER_DT>20191126</ORDER_DT>"
    ret=ret&"        <PAY_DTS>20191201224341</PAY_DTS>"
    ret=ret&"        <PRODUCT>"
    ret=ret&"            <PRD ID=""1"">"
    ret=ret&"                <ORD_ENGNM/>"
    ret=ret&"                <ENTR_DC_COUPON_AMT>950</ENTR_DC_COUPON_AMT>"
    ret=ret&"                <DC_COUPON_AMT>1950</DC_COUPON_AMT>"
    ret=ret&"                <ENTR_PRD_NO>2038410</ENTR_PRD_NO>"
    ret=ret&"                <SALE_UNITCOST>16100</SALE_UNITCOST>"
    ret=ret&"                <CURRENT_STATE>80</CURRENT_STATE>"
    ret=ret&"                <TOT_DC_COUPON_AMT>2900</TOT_DC_COUPON_AMT>"
    ret=ret&"                <IS_COLLECTED>N</IS_COLLECTED>"
    ret=ret&"                <IPOINT_DC_UNITCOST>0</IPOINT_DC_UNITCOST>"
    ret=ret&"                <OPT_PARENT_SEQ>1</OPT_PARENT_SEQ>"
    ret=ret&"                <ORD_AMT>32200</ORD_AMT>"
    ret=ret&"                <ENTR_DIS_UNIT_COST>0</ENTR_DIS_UNIT_COST>"
    ret=ret&"                <ORD_QTY>2</ORD_QTY>"
    ret=ret&"                <OPT_PRD_NO>5915510279</OPT_PRD_NO>"
    ret=ret&"                <OPT_PRD_TP>01</OPT_PRD_TP>"
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"
    ret=ret&"                <REAL_SALE_UNITCOST>16100</REAL_SALE_UNITCOST>"
    ret=ret&"                <PRE_USE_UNITCOST>0</PRE_USE_UNITCOST>"
    ret=ret&"                <ORD_SEQ>1</ORD_SEQ>"
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"
    ret=ret&"                <PRD_NO>5915510275</PRD_NO>"
    ret=ret&"                <IN_OPT_NM/>"
    ret=ret&"                <DELVSETL_SEQ>1</DELVSETL_SEQ>"
    ret=ret&"                <ORDCLM_STAT_DTS>20191202083713</ORDCLM_STAT_DTS>"
    ret=ret&"                <OPT_NO>Z110</OPT_NO>"
    ret=ret&"                <PRD_NM>텐바이텐 미스티코티타 YARN 미스티코티타 패브릭얀티셔츠얀  19 color</PRD_NM>"
    ret=ret&"                <SEL_OPT_NM>복합옵션 / 시크블랙.없음</SEL_OPT_NM>"
    ret=ret&"                <OPT_NM>복합옵션 / 시크블랙.없음</OPT_NM>"
    ret=ret&"                <ABROAD_BS_YN>N</ABROAD_BS_YN>"
    ret=ret&"                <PRE_USE_AMT>0</PRE_USE_AMT>"
    ret=ret&"            </PRD>"
    ret=ret&"        </PRODUCT>"
    ret=ret&"        <DELI_ADDR2_DORO>0000</DELI_ADDR2_DORO>"
    ret=ret&"        <DELI_MOBILE>000-000-0000</DELI_MOBILE>"
    ret=ret&"        <DELIVERY>"
    ret=ret&"            <DELV ID=""1"">"
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"
    ret=ret&"                <DEL_AMT>2500</DEL_AMT>"
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"
    ret=ret&"                <ADD_DEL_AMT>0</ADD_DEL_AMT>"
    ret=ret&"                <INITIAL_DELV_AMT>0</INITIAL_DELV_AMT>"
    ret=ret&"            </DELV>"
    ret=ret&"        </DELIVERY>"
    ret=ret&"        <EMAIL>000@hotmail.com</EMAIL>"
    ret=ret&"        <DELI_COMMENT/>"
    ret=ret&"        <TEL>00-000-0000</TEL>"
    ret=ret&"        <DEL_ZIP_DORO>13376</DEL_ZIP_DORO>"
    ret=ret&"        <DELIVERY_DETAIL>"
    ret=ret&"            <PRD_DELV ID=""1"">"
    ret=ret&"                <DELV_AMT>2500</DELV_AMT>"
    ret=ret&"                <ADD_DELV_AMT>0</ADD_DELV_AMT>"
    ret=ret&"                <DELVSETL_SEQ>1</DELVSETL_SEQ>"
    ret=ret&"            </PRD_DELV>"
    ret=ret&"        </DELIVERY_DETAIL>"
    ret=ret&"        <MOBILE_TEL>000-0000-0000</MOBILE_TEL>"
    ret=ret&"        <GIFT_MSG/>"
    ret=ret&"        <PAY_REF_MTHD_TP>02</PAY_REF_MTHD_TP>"
    ret=ret&"        <ORDER_DTS>20191126143039</ORDER_DTS>"
    ret=ret&"        <ORDCLM_CRT_TP>01</ORDCLM_CRT_TP>"
    ret=ret&"        <RESIDENT_NO/>"
    ret=ret&"        <DELI_ADDR2>0000</DELI_ADDR2>"
    ret=ret&"        <DELI_ADDR1>경기도 성남시 중원구 성남동  000 000</DELI_ADDR1>"
    ret=ret&"        <ORD_NO>20191126143039146839</ORD_NO>"
    ret=ret&"        <ORD_NM>0000</ORD_NM>"
    ret=ret&"        <DELI_TEL>000-0000-0000</DELI_TEL>"
    ret=ret&"        <DEL_ZIP>462829</DEL_ZIP>"
    ret=ret&"        <RCVR_ENM/>"
    ret=ret&"        <RCVR_NM>0000</RCVR_NM>"
    ret=ret&"    </ORDER>"
    ret=ret&"    <RESULT>"
    ret=ret&"        <CODE>000</CODE>"
    ret=ret&"        <MESSAGE>success</MESSAGE>"
    ret=ret&"        <LOG_SEQ>572428065</LOG_SEQ>"
    ret=ret&"    </RESULT>"
    ret=ret&"</ORDER_LIST>"
    getDevXNMLSAMPLE0 = ret
end function 

function getDevXNMLSAMPLE()
    dim ret : ret =""
    ret="<?xml version=""1.0"" encoding=""euc-kr""?>"&vbCRLF
    ret=ret&"<ORDER_LIST>"&vbCRLF
    ret=ret&"    <ORDER ID=""1"">"&vbCRLF
    ret=ret&"        <ORD_ENM/>"&vbCRLF
    ret=ret&"        <DELI_ADDR1_DORO>서울특별시 양천구 000 00-00 (신정동,)</DELI_ADDR1_DORO>"&vbCRLF
    ret=ret&"        <ORDER_DT>20191220</ORDER_DT>"&vbCRLF
    ret=ret&"        <PAY_DTS>20191220121247</PAY_DTS>"&vbCRLF
    ret=ret&"        <PRODUCT>"&vbCRLF
    ret=ret&"            <PRD ID=""1"">"&vbCRLF
    ret=ret&"                <ORD_ENGNM/>"&vbCRLF
    ret=ret&"                <ENTR_PRD_NO>1860114</ENTR_PRD_NO>"&vbCRLF
    ret=ret&"                <SALE_UNITCOST>31880</SALE_UNITCOST>"&vbCRLF
    ret=ret&"                <IS_COLLECTED>N</IS_COLLECTED>"&vbCRLF
    ret=ret&"                <OPT_PARENT_SEQ>1</OPT_PARENT_SEQ>"&vbCRLF
    ret=ret&"                <ORD_AMT>31880</ORD_AMT>"&vbCRLF
    ret=ret&"                <ENTR_DIS_UNIT_COST>0</ENTR_DIS_UNIT_COST>"&vbCRLF
    ret=ret&"                <ORD_QTY>1</ORD_QTY>"&vbCRLF
    ret=ret&"                <OPT_PRD_NO>5335443220</OPT_PRD_NO>"&vbCRLF
    ret=ret&"                <OPT_PRD_TP>01</OPT_PRD_TP>"&vbCRLF
    ret=ret&"                <DELV_DT>20191231</DELV_DT>"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <PRE_USE_UNITCOST>0</PRE_USE_UNITCOST>"&vbCRLF
    ret=ret&"                <ORD_SEQ>1</ORD_SEQ>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <PRD_NO>5335443186</PRD_NO>"&vbCRLF
    ret=ret&"                <IN_OPT_NM/>"&vbCRLF
    ret=ret&"                <DELVSETL_SEQ/>"&vbCRLF
    ret=ret&"                <ORDCLM_STAT_DTS>20191231105102</ORDCLM_STAT_DTS>"&vbCRLF
    ret=ret&"                <OPT_NO>0011</OPT_NO>"&vbCRLF
    ret=ret&"                <PRD_NM/>"&vbCRLF
    ret=ret&"                <SEL_OPT_NM>color / mint</SEL_OPT_NM>"&vbCRLF
    ret=ret&"                <DELV_DTS>20191231105102</DELV_DTS>"&vbCRLF
    ret=ret&"                <DELV_COMPLETE_DT>20200102</DELV_COMPLETE_DT>"&vbCRLF
    ret=ret&"                <OPT_NM>color / mint</OPT_NM>"&vbCRLF
    ret=ret&"                <ABROAD_BS_YN>N</ABROAD_BS_YN>"&vbCRLF
    ret=ret&"                <PRE_USE_AMT>0</PRE_USE_AMT>"&vbCRLF
    ret=ret&"                <DELV_COMP>169177</DELV_COMP>"&vbCRLF
    ret=ret&"                <DELV_NO>627559070632</DELV_NO>"&vbCRLF
    ret=ret&"            </PRD>"&vbCRLF
    ret=ret&"        </PRODUCT>"&vbCRLF
    ret=ret&"        <DELI_ADDR2_DORO>000 000</DELI_ADDR2_DORO>"&vbCRLF
    ret=ret&"        <DELI_MOBILE>000-0000-0000</DELI_MOBILE>"&vbCRLF
    ret=ret&"        <DELIVERY>"&vbCRLF
    ret=ret&"            <DELV ID=""1"">"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <DEL_AMT>2500</DEL_AMT>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <ADD_DEL_AMT>0</ADD_DEL_AMT>"&vbCRLF
    ret=ret&"                <INITIAL_DELV_AMT>0</INITIAL_DELV_AMT>"&vbCRLF
    ret=ret&"            </DELV>"&vbCRLF
    ret=ret&"        </DELIVERY>"&vbCRLF
    ret=ret&"        <EMAIL>00000@naver.com</EMAIL>"&vbCRLF
    ret=ret&"        <DELI_COMMENT>부재 시 집 앞에 놓고 가주세요 감사합니다</DELI_COMMENT>"&vbCRLF
    ret=ret&"        <TEL/>"&vbCRLF
    ret=ret&"        <DEL_ZIP_DORO>0000</DEL_ZIP_DORO>"&vbCRLF
    ret=ret&"        <MOBILE_TEL>000-0000-0000</MOBILE_TEL>"&vbCRLF
    ret=ret&"        <GIFT_MSG/>"&vbCRLF
    ret=ret&"        <PAY_REF_MTHD_TP>01</PAY_REF_MTHD_TP>"&vbCRLF
    ret=ret&"        <ORDER_DTS>20191220121156</ORDER_DTS>"&vbCRLF
    ret=ret&"        <ORDCLM_CRT_TP>15</ORDCLM_CRT_TP>"&vbCRLF
    ret=ret&"        <RESIDENT_NO/>"&vbCRLF
    ret=ret&"        <DELI_ADDR2>00 000</DELI_ADDR2>"&vbCRLF
    ret=ret&"        <DELI_ADDR1>서울특별시 양천구 신정동 000-000</DELI_ADDR1>"&vbCRLF
    ret=ret&"        <ORD_NO>20191220121156120473</ORD_NO>"&vbCRLF
    ret=ret&"        <ORD_NM>0000</ORD_NM>"&vbCRLF
    ret=ret&"        <DELI_TEL>000-000-0000</DELI_TEL>"&vbCRLF
    ret=ret&"        <DEL_ZIP>158871</DEL_ZIP>"&vbCRLF
    ret=ret&"        <RCVR_ENM/>"&vbCRLF
    ret=ret&"        <RCVR_NM>0000</RCVR_NM>"&vbCRLF
    ret=ret&"    </ORDER>"&vbCRLF
    ret=ret&"    <ORDER ID=""2"">"&vbCRLF
    ret=ret&"        <ORD_ENM/>"&vbCRLF
    ret=ret&"        <DELI_ADDR1_DORO>충청남도 논산시 대림길 000 (부창동,대림아파트)</DELI_ADDR1_DORO>"&vbCRLF
    ret=ret&"        <ORDER_DT>20191222</ORDER_DT>"&vbCRLF
    ret=ret&"        <PAY_DTS>20191222104524</PAY_DTS>"&vbCRLF
    ret=ret&"        <PRODUCT>"&vbCRLF
    ret=ret&"            <PRD ID=""1"">"&vbCRLF
    ret=ret&"                <ORD_ENGNM/>"&vbCRLF
    ret=ret&"                <ENTR_PRD_NO>2440993</ENTR_PRD_NO>"&vbCRLF
    ret=ret&"                <SALE_UNITCOST>17230</SALE_UNITCOST>"&vbCRLF
    ret=ret&"                <IS_COLLECTED>N</IS_COLLECTED>"&vbCRLF
    ret=ret&"                <OPT_PARENT_SEQ>1</OPT_PARENT_SEQ>"&vbCRLF
    ret=ret&"                <ORD_AMT>17230</ORD_AMT>"&vbCRLF
    ret=ret&"                <ENTR_DIS_UNIT_COST>0</ENTR_DIS_UNIT_COST>"&vbCRLF
    ret=ret&"                <ORD_QTY>1</ORD_QTY>"&vbCRLF
    ret=ret&"                <OPT_PRD_NO>6685878256</OPT_PRD_NO>"&vbCRLF
    ret=ret&"                <OPT_PRD_TP>01</OPT_PRD_TP>"&vbCRLF
    ret=ret&"                <DELV_DT>20191231</DELV_DT>"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <PRE_USE_UNITCOST>0</PRE_USE_UNITCOST>"&vbCRLF
    ret=ret&"                <ORD_SEQ>1</ORD_SEQ>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <PRD_NO>6685878250</PRD_NO>"&vbCRLF
    ret=ret&"                <IN_OPT_NM/>"&vbCRLF
    ret=ret&"                <DELVSETL_SEQ/>"&vbCRLF
    ret=ret&"                <ORDCLM_STAT_DTS>20191231165101</ORDCLM_STAT_DTS>"&vbCRLF
    ret=ret&"                <OPT_NO>Z132</OPT_NO>"&vbCRLF
    ret=ret&"                <PRD_NM/>"&vbCRLF
    ret=ret&"                <SEL_OPT_NM>복합옵션 / 그레이.L(80-85).L(100-105)</SEL_OPT_NM>"&vbCRLF
    ret=ret&"                <DELV_DTS>20191231165101</DELV_DTS>"&vbCRLF
    ret=ret&"                <DELV_COMPLETE_DT>20200102</DELV_COMPLETE_DT>"&vbCRLF
    ret=ret&"                <OPT_NM>복합옵션 / 그레이.L(80-85).L(100-105)</OPT_NM>"&vbCRLF
    ret=ret&"                <ABROAD_BS_YN>N</ABROAD_BS_YN>"&vbCRLF
    ret=ret&"                <PRE_USE_AMT>0</PRE_USE_AMT>"&vbCRLF
    ret=ret&"                <DELV_COMP>169198</DELV_COMP>"&vbCRLF
    ret=ret&"                <DELV_NO>401964188220</DELV_NO>"&vbCRLF
    ret=ret&"            </PRD>"&vbCRLF
    ret=ret&"            <PRD ID=""2"">"&vbCRLF
    ret=ret&"                <ORD_ENGNM/>"&vbCRLF
    ret=ret&"                <ENTR_PRD_NO>2440993</ENTR_PRD_NO>"&vbCRLF
    ret=ret&"                <SALE_UNITCOST>17230</SALE_UNITCOST>"&vbCRLF
    ret=ret&"                <IS_COLLECTED>N</IS_COLLECTED>"&vbCRLF
    ret=ret&"                <OPT_PARENT_SEQ>2</OPT_PARENT_SEQ>"&vbCRLF
    ret=ret&"                <ORD_AMT>17230</ORD_AMT>"&vbCRLF
    ret=ret&"                <ENTR_DIS_UNIT_COST>0</ENTR_DIS_UNIT_COST>"&vbCRLF
    ret=ret&"                <ORD_QTY>1</ORD_QTY>"&vbCRLF
    ret=ret&"                <OPT_PRD_NO>6685878265</OPT_PRD_NO>"&vbCRLF
    ret=ret&"                <OPT_PRD_TP>01</OPT_PRD_TP>"&vbCRLF
    ret=ret&"                <DELV_DT>20191231</DELV_DT>"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <PRE_USE_UNITCOST>0</PRE_USE_UNITCOST>"&vbCRLF
    ret=ret&"                <ORD_SEQ>2</ORD_SEQ>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <PRD_NO>6685878250</PRD_NO>"&vbCRLF
    ret=ret&"                <IN_OPT_NM/>"&vbCRLF
    ret=ret&"                <DELVSETL_SEQ/>"&vbCRLF
    ret=ret&"                <ORDCLM_STAT_DTS>20191231165100</ORDCLM_STAT_DTS>"&vbCRLF
    ret=ret&"                <OPT_NO>Z232</OPT_NO>"&vbCRLF
    ret=ret&"                <PRD_NM/>"&vbCRLF
    ret=ret&"                <SEL_OPT_NM>복합옵션 / 블루.L(80-85).L(100-105)</SEL_OPT_NM>"&vbCRLF
    ret=ret&"                <DELV_DTS>20191231165100</DELV_DTS>"&vbCRLF
    ret=ret&"                <DELV_COMPLETE_DT>20200102</DELV_COMPLETE_DT>"&vbCRLF
    ret=ret&"                <OPT_NM>복합옵션 / 블루.L(80-85).L(100-105)</OPT_NM>"&vbCRLF
    ret=ret&"                <ABROAD_BS_YN>N</ABROAD_BS_YN>"&vbCRLF
    ret=ret&"                <PRE_USE_AMT>0</PRE_USE_AMT>"&vbCRLF
    ret=ret&"                <DELV_COMP>169198</DELV_COMP>"&vbCRLF
    ret=ret&"                <DELV_NO>401964188220</DELV_NO>"&vbCRLF
    ret=ret&"            </PRD>"&vbCRLF
    ret=ret&"        </PRODUCT>"&vbCRLF
    ret=ret&"        <DELI_ADDR2_DORO>000 000</DELI_ADDR2_DORO>"&vbCRLF
    ret=ret&"        <DELI_MOBILE>010-0000-0000</DELI_MOBILE>"&vbCRLF
    ret=ret&"        <DELIVERY>"&vbCRLF
    ret=ret&"            <DELV ID=""1"">"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <DEL_AMT>2500</DEL_AMT>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <ADD_DEL_AMT>0</ADD_DEL_AMT>"&vbCRLF
    ret=ret&"                <INITIAL_DELV_AMT>0</INITIAL_DELV_AMT>"&vbCRLF
    ret=ret&"            </DELV>"&vbCRLF
    ret=ret&"        </DELIVERY>"&vbCRLF
    ret=ret&"        <EMAIL>00000@hanmail.net</EMAIL>"&vbCRLF
    ret=ret&"        <DELI_COMMENT/>"&vbCRLF
    ret=ret&"        <TEL>000-000-000</TEL>"&vbCRLF
    ret=ret&"        <DEL_ZIP_DORO>32955</DEL_ZIP_DORO>"&vbCRLF
    ret=ret&"        <MOBILE_TEL>000-000-0000</MOBILE_TEL>"&vbCRLF
    ret=ret&"        <GIFT_MSG/>"&vbCRLF
    ret=ret&"        <PAY_REF_MTHD_TP>01</PAY_REF_MTHD_TP>"&vbCRLF
    ret=ret&"        <ORDER_DTS>20191222104442</ORDER_DTS>"&vbCRLF
    ret=ret&"        <ORDCLM_CRT_TP>15</ORDCLM_CRT_TP>"&vbCRLF
    ret=ret&"        <RESIDENT_NO/>"&vbCRLF
    ret=ret&"        <DELI_ADDR2>100동 000호</DELI_ADDR2>"&vbCRLF
    ret=ret&"        <DELI_ADDR1>충청남도 논산시 부창동 000-0 000</DELI_ADDR1>"&vbCRLF
    ret=ret&"        <ORD_NO>20191222104442420745</ORD_NO>"&vbCRLF
    ret=ret&"        <ORD_NM>aaa</ORD_NM>"&vbCRLF
    ret=ret&"        <DELI_TEL>000-000-000</DELI_TEL>"&vbCRLF
    ret=ret&"        <DEL_ZIP>320753</DEL_ZIP>"&vbCRLF
    ret=ret&"        <RCVR_ENM/>"&vbCRLF
    ret=ret&"        <RCVR_NM>aaa</RCVR_NM>"&vbCRLF
    ret=ret&"    </ORDER>"&vbCRLF
    ret=ret&"    <ORDER ID=""207"">"&vbCRLF
    ret=ret&"        <ORD_ENM/>"&vbCRLF
    ret=ret&"        <DELI_ADDR1_DORO>서울특별시 동작구 신대방0길 0 (신대방동,)</DELI_ADDR1_DORO>"&vbCRLF
    ret=ret&"        <ORDER_DT>20191231</ORDER_DT>"&vbCRLF
    ret=ret&"        <PAY_DTS>20191231135441</PAY_DTS>"&vbCRLF
    ret=ret&"        <PRODUCT>"&vbCRLF
    ret=ret&"            <PRD ID=""1"">"&vbCRLF
    ret=ret&"                <ORD_ENGNM/>"&vbCRLF
    ret=ret&"                <ENTR_PRD_NO>2420186</ENTR_PRD_NO>"&vbCRLF
    ret=ret&"                <SALE_UNITCOST>18800</SALE_UNITCOST>"&vbCRLF
    ret=ret&"                <IS_COLLECTED>N</IS_COLLECTED>"&vbCRLF
    ret=ret&"                <OPT_PARENT_SEQ>1</OPT_PARENT_SEQ>"&vbCRLF
    ret=ret&"                <ORD_AMT>18800</ORD_AMT>"&vbCRLF
    ret=ret&"                <ENTR_DIS_UNIT_COST>0</ENTR_DIS_UNIT_COST>"&vbCRLF
    ret=ret&"                <ORD_QTY>1</ORD_QTY>"&vbCRLF
    ret=ret&"                <OPT_PRD_NO>6874405709</OPT_PRD_NO>"&vbCRLF
    ret=ret&"                <OPT_PRD_TP>01</OPT_PRD_TP>"&vbCRLF
    ret=ret&"                <DELV_DT>20191231</DELV_DT>"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <PRE_USE_UNITCOST>0</PRE_USE_UNITCOST>"&vbCRLF
    ret=ret&"                <ORD_SEQ>1</ORD_SEQ>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <PRD_NO>6874405709</PRD_NO>"&vbCRLF
    ret=ret&"                <IN_OPT_NM/>"&vbCRLF
    ret=ret&"                <DELVSETL_SEQ/>"&vbCRLF
    ret=ret&"                <ORDCLM_STAT_DTS>20191231165305</ORDCLM_STAT_DTS>"&vbCRLF
    ret=ret&"                <OPT_NO>2420186</OPT_NO>"&vbCRLF
    ret=ret&"                <PRD_NM/>"&vbCRLF
    ret=ret&"                <SEL_OPT_NM/>"&vbCRLF
    ret=ret&"                <DELV_DTS>20191231165305</DELV_DTS>"&vbCRLF
    ret=ret&"                <DELV_COMPLETE_DT>20200102</DELV_COMPLETE_DT>"&vbCRLF
    ret=ret&"                <OPT_NM/>"&vbCRLF
    ret=ret&"                <ABROAD_BS_YN>N</ABROAD_BS_YN>"&vbCRLF
    ret=ret&"                <PRE_USE_AMT>0</PRE_USE_AMT>"&vbCRLF
    ret=ret&"                <DELV_COMP>169168</DELV_COMP>"&vbCRLF
    ret=ret&"                <DELV_NO>356800507976</DELV_NO>"&vbCRLF
    ret=ret&"            </PRD>"&vbCRLF
    ret=ret&"        </PRODUCT>"&vbCRLF
    ret=ret&"        <DELI_ADDR2_DORO>000호</DELI_ADDR2_DORO>"&vbCRLF
    ret=ret&"        <DELI_MOBILE>000-0000-0000</DELI_MOBILE>"&vbCRLF
    ret=ret&"        <DELIVERY>"&vbCRLF
    ret=ret&"            <DELV ID=""1"">"&vbCRLF
    ret=ret&"                <SUPPLY_ENTR_NO>3000010614</SUPPLY_ENTR_NO>"&vbCRLF
    ret=ret&"                <DEL_AMT>2500</DEL_AMT>"&vbCRLF
    ret=ret&"                <SUPPLY_CTRT_SEQ>2</SUPPLY_CTRT_SEQ>"&vbCRLF
    ret=ret&"                <ADD_DEL_AMT>0</ADD_DEL_AMT>"&vbCRLF
    ret=ret&"                <INITIAL_DELV_AMT>0</INITIAL_DELV_AMT>"&vbCRLF
    ret=ret&"            </DELV>"&vbCRLF
    ret=ret&"        </DELIVERY>"&vbCRLF
    ret=ret&"        <EMAIL>aaa@naver.com</EMAIL>"&vbCRLF
    ret=ret&"        <DELI_COMMENT/>"&vbCRLF
    ret=ret&"        <TEL/>"&vbCRLF
    ret=ret&"        <DEL_ZIP_DORO>0000</DEL_ZIP_DORO>"&vbCRLF
    ret=ret&"        <MOBILE_TEL/>"&vbCRLF
    ret=ret&"        <GIFT_MSG/>"&vbCRLF
    ret=ret&"        <PAY_REF_MTHD_TP>99</PAY_REF_MTHD_TP>"&vbCRLF
    ret=ret&"        <ORDER_DTS>20191231135346</ORDER_DTS>"&vbCRLF
    ret=ret&"        <ORDCLM_CRT_TP>15</ORDCLM_CRT_TP>"&vbCRLF
    ret=ret&"        <RESIDENT_NO/>"&vbCRLF
    ret=ret&"        <DELI_ADDR2>aaa호</DELI_ADDR2>"&vbCRLF
    ret=ret&"        <DELI_ADDR1>서울특별시 동작구 신대방동 aaa-aaa</DELI_ADDR1>"&vbCRLF
    ret=ret&"        <ORD_NO>20191231135346000076</ORD_NO>"&vbCRLF
    ret=ret&"        <ORD_NM>aaaaaa</ORD_NM>"&vbCRLF
    ret=ret&"        <DELI_TEL>000-0000-0000</DELI_TEL>"&vbCRLF
    ret=ret&"        <DEL_ZIP>156853</DEL_ZIP>"&vbCRLF
    ret=ret&"        <RCVR_ENM/>"&vbCRLF
    ret=ret&"        <RCVR_NM>aaaa</RCVR_NM>"&vbCRLF
    ret=ret&"    </ORDER>"&vbCRLF
    ret=ret&"    <RESULT>"&vbCRLF
    ret=ret&"        <CODE>000</CODE>"&vbCRLF
    ret=ret&"        <MESSAGE>success</MESSAGE>"&vbCRLF
    ret=ret&"        <LOG_SEQ>572491389</LOG_SEQ>"&vbCRLF
    ret=ret&"    </RESULT>"&vbCRLF
    ret=ret&"</ORDER_LIST>"&vbCRLF
    getDevXNMLSAMPLE = ret
end function 

Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE 

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-0,now()),10)

    if eddt<>"" then 
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-3,iedyyyymmdd),10)
'' 

'' 조회 유형 orderListDelvForSingle / delvCompListForSingle
'' 다른몰과 다르게 한번 조회하면 더이상 조회 할 필요가 없음.. (최근것만 조회하면된다.. D-1)
'' 주문확인 리스트는 한번만 하면됨.. 배송정보는 기존걸 할 필요가 좀 있음 (송장변경건등을 확인하기 위해..)
if (request("retry")<>"") then
    ''confirmDt ="2020010101"
    confirmDt =request("confirmDt")
end if

Dim lastconfirmDT
sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_LastInputConfirmDT] 'interpark','"&confirmDt&"'"
dbget.CursorLocation = adUseClient
rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
if NOT rsget.Eof then
    lastconfirmDT = rsget("lastconfirmDT")
    rw "lastconfirmDT:"&lastconfirmDT
end if
rsget.close()

if (request("retry")="") then
    sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'interpark','"&confirmDt&"'"
    dbget.Execute sqlStr
    rw "초기화작업"
end if

if (request("thedate")<>"") then
    call Get_InterParkOrderListByStatus(request("thedate"),request("thedate"),"delvCompListForSingle","배송정보",request("thetime")) ''배송정보리스트
end if

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k

for k=0 to datelen ''-1
    thedate=dateadd("d",1*k,istyyyymmdd)
''rw ":"&thedate&":"&(k<datelen)&":"&(CDate(thedate)>=CDate(lastconfirmDT))&":"&thedate&":"&lastconfirmDT
    if k<datelen then
        if (CDate(thedate)>=CDate(lastconfirmDT)) then
            call Get_InterParkOrderListByStatus(thedate,thedate,"orderListDelvForSingle","주문확인",0)
            response.flush
        end if
    end if
    call Get_InterParkOrderListByStatus(thedate,thedate,"delvCompListForSingle","배송정보",1) ''배송정보리스트
    response.flush
    call Get_InterParkOrderListByStatus(thedate,thedate,"delvCompListForSingle","배송정보",2) ''배송정보리스트
    response.flush
    call Get_InterParkOrderListByStatus(thedate,thedate,"delvCompListForSingle","배송정보",3) ''배송정보리스트
    response.flush
    call Get_InterParkOrderListByStatus(thedate,thedate,"delvCompListForSingle","배송정보",4) ''배송정보리스트
    response.flush
next

dim retryCnt : retryCnt = request("retry")
if (retryCnt="") then retryCnt=0
if (retryCnt<3) then
    response.write "<script>location.href='?retry="&retryCnt+1&"&confirmDt="&confirmDt&"&eddt="&LEFT(dateadd("d",-1,istyyyymmdd),10)&"'</script>"
    dbget.Close:response.end
end if

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'interpark','"&confirmDt&"'"
dbget.Execute sqlStr
rw "주문매핑"

rw "완료"
'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function Get_InterParkOrderListByStatus(stdate,eddate,iorderStatus,istatusName,ipartial)
	dim sellsite : sellsite = "interpark"
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

	Get_InterParkOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	''xmlSelldate = Replace(selldate, "-", "")

	'// 주문내역조회2 , 발주확인된 리스트
	xmlURL = "https://joinapi.interpark.com"
	xmlURL = xmlURL&"/order/OrderClmAPI.do?_method="&iorderStatus&"&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2"
    if (ipartial=1) then
        xmlURL = xmlURL&"&sc.strDate=" + replace(stdate,"-","") + "000000" + "&sc.endDate=" + replace(eddate,"-","") + "115959"
    elseif (ipartial=2) then
        xmlURL = xmlURL&"&sc.strDate=" + replace(stdate,"-","") + "120000" + "&sc.endDate=" + replace(eddate,"-","") + "145959"
    elseif (ipartial=3) then
        xmlURL = xmlURL&"&sc.strDate=" + replace(stdate,"-","") + "150000" + "&sc.endDate=" + replace(eddate,"-","") + "165959"
    elseif (ipartial=4) then
        xmlURL = xmlURL&"&sc.strDate=" + replace(stdate,"-","") + "170000" + "&sc.endDate=" + replace(eddate,"-","") + "235959"
    else
        xmlURL = xmlURL&"&sc.strDate=" + replace(stdate,"-","") + "000000" + "&sc.endDate=" + replace(eddate,"-","") + "235959"
    end if


''GetXMLURL = FRectAPIURL + "/order/OrderClmAPI.do?_method=orderListDelvForSingle&sc.entrId=10X10&sc.supplyEntrNo=3000010614&sc.supplyCtrtSeq=2&sc.strDate=" + CStr(GetxSiteDateFormat(FRectStartYYYYMMDD)) + "000000" + "&sc.endDate=" + CStr(GetxSiteDateFormat(FRectEndYYYYMMDD)) + "235959"
	
    if (ipartial=1) then
        rw "기간검색:"&stdate&"~"&eddate&" 11:59 상태:"&iorderStatus&"("&istatusName&")"
    elseif (ipartial=2) then
        rw "기간검색:"&stdate&" 12:00~"&eddate&" 14:59 상태:"&iorderStatus&"("&istatusName&")"
    elseif (ipartial=3) then
        rw "기간검색:"&stdate&" 15:00~"&eddate&" 16:59 상태:"&iorderStatus&"("&istatusName&")"
    elseif (ipartial=4) then
        rw "기간검색:"&stdate&" 17:00~"&eddate&" 23:59 상태:"&iorderStatus&"("&istatusName&")"
    else
        rw "기간검색:"&stdate&"~"&eddate&" 상태:"&iorderStatus&"("&istatusName&")"
    end if
    
	'// =======================================================================
	'// 데이타 가져오기
    

	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", xmlURL, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
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
    Dim ORDCLM_STAT_DTS, DELV_DTS, DELV_COMPLETE_DT

    if (application("Svr_Info")<>"Dev") then
	    iRbody = BinaryToText(objXML.ResponseBody, "euc-kr")
    else
        iRbody = getDevXNMLSAMPLE0() ''getDevXNMLSAMPLE()
    end if

if (request("thedate")<>"") then
    rw "<textarea cols=80 rows=20>"&iRbody&"</textarea>"
    dbget.close : response.end
end if

    Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
	xmlDOM.async = False
	xmlDOM.LoadXML replace(iRbody,"&","＆")

	Set obj = xmlDOM.selectNodes("/ORDER_LIST/ORDER")

	if obj is Nothing then

		rw "No outPutValue"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if
    
    orderCount = (xmlDOM.selectNodes("/ORDER_LIST/ORDER").length)
	''response.write masterCnt

	if orderCount = 0 then

		rw "list - 0"
		Set xmlDOM = Nothing
		Set objXML = Nothing
		exit function
	end if


    set objMasterListXML = xmlDOM.selectNodes("/ORDER_LIST/ORDER")
	orderCount = objMasterListXML.length

	response.write "주문건수(" & orderCount & ") " & "<br />"

	for i = 0 to orderCount - 1
        set objMasterOneXML = objMasterListXML.item(i)
        ordNo           = objMasterOneXML.selectSingleNode("ORD_NO").text
        ''objMasterOneXML.selectSingleNode("PAY_DTS").text                            ''입금일자
        set objDetailListXML = objMasterOneXML.selectNodes("PRODUCT/PRD")
		detailCnt = objDetailListXML.length
        For j = 0 to detailCnt - 1
            Set objDetailOneXML = objDetailListXML.item(j)
            ordItemSeq      = objDetailOneXML.selectSingleNode("ORD_SEQ").text
            cspGoodsCd      = objDetailOneXML.selectSingleNode("ENTR_PRD_NO").text   ''제휴업체 상품 코드
            ''objDetailOneXML.selectSingleNode("OPT_NO").text                        ''제휴업체 옵션코드
            goodsCd = objDetailOneXML.selectSingleNode("PRD_NO").text                ''인터파크상품코드
			uitemId = objDetailOneXML.selectSingleNode("OPT_PRD_NO").text            ''옵션코드
            optionContent = (objDetailOneXML.selectSingleNode("OPT_NM").text)
            orderQty = objDetailOneXML.selectSingleNode("ORD_QTY").text

             
            ORDCLM_STAT_DTS = "" : DELV_DTS="" : DELV_COMPLETE_DT =""
            if (NOT (objDetailOneXML.selectSingleNode("ORDCLM_STAT_DTS") is Nothing)) then
                ORDCLM_STAT_DTS = objDetailOneXML.selectSingleNode("ORDCLM_STAT_DTS").text          ''발주확인일자
            end if

            if (NOT (objDetailOneXML.selectSingleNode("DELV_DTS") is Nothing)) then
                DELV_DTS        = objDetailOneXML.selectSingleNode("DELV_DTS").text                 ''출고일자(20191231082100)
            end if

            if (NOT (objDetailOneXML.selectSingleNode("DELV_COMPLETE_DT") is Nothing)) then
                DELV_COMPLETE_DT= objDetailOneXML.selectSingleNode("DELV_COMPLETE_DT").text         ''배송완료일(20200101)
            end if

            shppNo		    = "" 'obj2.get(j).orderNo			    'reserve01(주문번호)
            
            shppDivDtlNm = "일반출고"
            shppRsvtDt      = "" ''예정일
            autoShortgYn    = "" ''자동결품여부
            invoiceUpDt = "" ''운송장번호 업로드 일시 (이게 오래된거면 추적(집하)이 안된거 일 수 있다.) = 출고일
            if (DELV_DTS<>"") then
                invoiceUpDt = LEFT(DELV_DTS,4)+"-"+MID(DELV_DTS,5,2)+"-"+MID(DELV_DTS,7,2)
            end if
            shppSeq	  = ""	'배송지시상세번호
            reOrderYn ="N" ''재주문여부 
            delayNts  =""  ''지연일수

            shppTypeDtlNm   = ""

            delicoVenId = "" : wblNo="" : delicoVenNm=""
            if (NOT (objDetailOneXML.selectSingleNode("DELV_COMP") is Nothing)) then 
                delicoVenId     = objDetailOneXML.selectSingleNode("DELV_COMP").text '택배배송사코드
            end if
            if (NOT (objDetailOneXML.selectSingleNode("DELV_NO") is Nothing)) then 
                wblNo           = objDetailOneXML.selectSingleNode("DELV_NO").text   '운송장번호
            end if

            ' if (shppTypeDtlNm="기타배송") then 
            '     wblNo = wblNo & obj1.get(i).delivery.shipMethodMessage					'배송방법 메세지 배송방법이 [기타배송]일 경우 입력받는 메세지
            ' end if
            delicoVenNm     = getiparkDlvCode2Name(delicoVenId) '택배사명     

            orderStatus     = ""  
            if (ORDCLM_STAT_DTS<>"") then   
                orderStatus  = "주문확인" 
            end if 
            if (DELV_DTS<>"") then   
                orderStatus  = "출고완료" 
            end if 
            if (DELV_COMPLETE_DT<>"") then   
                orderStatus  = "배송완료" 
            end if 

            whoutCritnDt    = "" '' 발송기한.
            outjFixedDt     = "" ''구매확정일자  - 업체직송인경우 7일후 완료된다. 정산이 안되면 업체직송으로 수정해야한다.


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

            'if NOT (Trim(orderStatus)="배송완료") then
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
                    ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, LEFT(Trim(optionContent),60)) _
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
                '  If ERR then
                '      rw invoiceUpDt 
                '      rw outjFixedDt
                '      response.end
                '  end if
                RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
                
                successCnt = successCnt+1
            'end if
        next
        
    Next

    set objMasterOneXML = Nothing
    set objDetailListXML = Nothing


    rw "상세건수:"&successCnt
    rw "======================================"

	Get_InterParkOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->