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


Dim sqlStr
Dim isOnlyTodayBaljuView : isOnlyTodayBaljuView=false  ''발주된내역만 (주문수신 view)
Dim isDlvConfirmProc 	 : isDlvConfirmProc=false  ''주문확인 Proc
Dim isDlvInputProc 	 	 : isDlvInputProc=false    ''입력 Proc
Dim eddt : eddt=requestCheckvar(request("eddt"),10)

Dim IS_TEST_MODE : IS_TEST_MODE=FALSE 

Dim istyyyymmdd, iedyyyymmdd
    iedyyyymmdd = LEFT(dateadd("d",-2,now()),10)

    if eddt<>"" then 
        if isDate(eddt) then
            iedyyyymmdd=eddt  '''yyyy-mm-dd
        end if
    end if
    istyyyymmdd = LEFT(dateadd("d",-26,iedyyyymmdd),10)
'' 


sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_FinFlagDefaultSET] 'cjmall','"&confirmDt&"'"
dbget.Execute sqlStr

dim datelen : datelen=datediff("d",istyyyymmdd, iedyyyymmdd)
dim thedate , k

for k=0 to datelen-1
    thedate=dateadd("d",-1*k,iedyyyymmdd)
    
    call Get_CjmallOrderListByStatus(thedate)
    response.flush
    
next

sqlStr = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] 'cjmall','"&confirmDt&"'"
dbget.Execute sqlStr

rw "완료"


' for k=0 to datelen-1
'     thedate=dateadd("d",-1*k,iedyyyymmdd)
'     if k<5 then
'     call Get_WMPOrderListByStatus(thedate,thedate,"NEW","주문통보")
'     response.flush
'     end if
'     call Get_WMPOrderListByStatus(thedate,thedate,"CONFIRM","주문확인")
'     response.flush
'     call Get_WMPOrderListByStatus(thedate,thedate,"DELIVERY","출고완료") 
'     response.flush

'     '' call Get_WMPOrderListByStatus(istyyyymmdd,iedyyyymmdd,"COMPLETE","배송완료")
'     '' response.flush
' next


'response.write("<script>setTimeout(alert('완료'),1000);self.close();</script>")

function getTESTSamplXML()
Dim ret
ret=ret&"<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
ret=ret&"<ns1:ifResponse ns1:ifId=""IF_04_01"" xmlns:ns1=""http://www.example.org/ifpa"">"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191203160279</ns1:ordNo>"
ret=ret&"        <ns1:custNm>이장*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0503)6337-1280</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>이장*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>650829</ns1:zipno>"
ret=ret&"            <ns1:addr_1>경남 통영시 광도면 죽림리 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1567-2번지 해피데이 401호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0503)6337-1280</ns1:telno>"
ret=ret&"            <ns1:cellno>0503)6337-1280</ns1:cellno>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56828733</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12091758553</ns1:unitCd>"
ret=ret&"            <ns1:itemName>[스위스밀리터리] OKK 보온병 4종세트사은품빅백</ns1:itemName>"
ret=ret&"            <ns1:unitNm>[스위스밀리터리] OKK 보온병 4종...</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2353929_0000</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791401894</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>81400.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>75710.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-03 23:48:42</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204001419</ns1:ordNo>"
ret=ret&"        <ns1:custNm>최제*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1561-2886</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>최제*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>446715</ns1:zipno>"
ret=ret&"            <ns1:addr_1>경기 용인시 기흥구 중동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1050번지 어은목마을코아루아파트 4307동 1801호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1561-2886</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1561-2886</ns1:cellno>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>54205903</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12077600723</ns1:unitCd>"
ret=ret&"            <ns1:itemName>시네마 레터링 라이트박스 빅 무드램프 36x18cm</ns1:itemName>"
ret=ret&"            <ns1:unitNm>시네마 레터링 라이트박스 빅 무드램프...</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2207717_0000</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791397627</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>14900.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>13860.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 00:10:41</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204006842</ns1:ordNo>"
ret=ret&"        <ns1:custNm>쿠*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0503)6337-4710</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>INTERNET</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>김정*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>314080</ns1:zipno>"
ret=ret&"            <ns1:addr_1>충남 공주시 금학동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>389번지 금학e-편한세상 108동 904호 금학E-편한세상)</ns1:addr_2>"
ret=ret&"            <ns1:telno>0503)6337-4710</ns1:telno>"
ret=ret&"            <ns1:cellno>0503)6337-4710</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>문 앞</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>60574396</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12111573809</ns1:unitCd>"
ret=ret&"            <ns1:itemName>제로넥카라</ns1:itemName>"
ret=ret&"            <ns1:unitNm>XL,민트</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2588673_Z430</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791401656</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>24000.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>24000.0</ns1:outwAmt>"
ret=ret&"            <ns1:delivInfo>문 앞</ns1:delivInfo>"
ret=ret&"            <ns1:ordDtm>2019-12-04 01:24:27</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204007609</ns1:ordNo>"
ret=ret&"        <ns1:custNm>김연*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1558-5787</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>김현*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>200931</ns1:zipno>"
ret=ret&"            <ns1:addr_1>강원 춘천시 근화동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>267-1번지 춘천L-타워 1차 619호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1558-5787</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1558-5787</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>부재 시 현관앞에 놓아주세요. 감사합니다!</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>55961289</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12087167253</ns1:unitCd>"
ret=ret&"            <ns1:itemName>멜로즈 투명커버 빅사이즈 2단분리 화장품 정리함</ns1:itemName>"
ret=ret&"            <ns1:unitNm>그레이</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2300791_0012</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791397372</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>37000.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>33904.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 01:44:37</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204007896</ns1:ordNo>"
ret=ret&"        <ns1:custNm>이은*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1567-9290</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>이은*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>151778</ns1:zipno>"
ret=ret&"            <ns1:addr_1>서울 관악구 은천동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1718번지 벽산블루밍아파트 103동 102호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1567-9290</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1567-9290</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>부재시 문앞에놓아주세요</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>55480625</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12084571791</ns1:unitCd>"
ret=ret&"            <ns1:itemName>[한일카페트] 리디아 순면 워싱 카페트 200x230</ns1:itemName>"
ret=ret&"            <ns1:unitNm>핑크민트</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2274759_0010</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791397190</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>50900.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>47340.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 01:52:11</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204019311</ns1:ordNo>"
ret=ret&"        <ns1:custNm>장미*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1920-3414</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>장미*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>415808</ns1:zipno>"
ret=ret&"            <ns1:addr_1>경기 김포시 풍무동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>354-1번지 풍무푸르지오 108동 2401호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1920-3414</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1920-3414</ns1:cellno>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56293607</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12088910442</ns1:unitCd>"
ret=ret&"            <ns1:itemName>스트링선반-화이트</ns1:itemName>"
ret=ret&"            <ns1:unitNm>스트링선반-화이트</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>1285464_0000</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791417845</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>30400.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>28280.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 07:31:18</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204021748</ns1:ordNo>"
ret=ret&"        <ns1:custNm>이지*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1919-1079</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>이지*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>426785</ns1:zipno>"
ret=ret&"            <ns1:addr_1>경기 안산시 상록구 월피동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>444번지 주공1단지아파트 113동 1303호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1919-1079</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1919-1079</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>오시기전 꼭 핸드폰 연락주세요. 감사합니다.</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56879787</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12091978787</ns1:unitCd>"
ret=ret&"            <ns1:itemName>까사니 아이스크림만들기 세트_2056004</ns1:itemName>"
ret=ret&"            <ns1:unitNm>까사니 아이스크림만들기 세트_(205...</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2343193_0000</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791421538</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>23500.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>20770.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 08:09:25</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204022222</ns1:ordNo>"
ret=ret&"        <ns1:custNm>윤귀*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1922-4336</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>주문 </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-06+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>윤귀*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>791755</ns1:zipno>"
ret=ret&"            <ns1:addr_1>경북 포항시 북구 용흥동 </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>464번지 한라타워맨션 302동 1803호</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1922-4336</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1922-4336</ns1:cellno>"
ret=ret&"            <ns1:packYn>일반</ns1:packYn>"
ret=ret&"            <ns1:itemCd>57250966</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12093622037</ns1:unitCd>"
ret=ret&"            <ns1:itemName>제이라이프 안티 곰팡이 샤워커튼</ns1:itemName>"
ret=ret&"            <ns1:unitNm>제이라이프 안티 곰팡이 샤워커튼</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2380820_0000</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791421346</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>16900.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>14940.0</ns1:outwAmt>"
ret=ret&"            <ns1:ordDtm>2019-12-04 08:16:09</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"

ret=ret&"</ns1:ifResponse>"
getTESTSamplXML = ret
end function

function Get_CjmallOrderListByStatus(thedate)
	dim sellsite : sellsite = "cjmall"
	dim cjMallAPIURL, xmlSelldate
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

	Get_CjmallOrderListByStatus = False

	'// =======================================================================
	'// 날짜형식
	''selldate = "2017-11-10"
	

    IF application("Svr_Info")="Dev" THEN
        'cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' 테스트서버
        cjMallAPIURL = "http://210.122.101.154:8210/IFPAServerAction.action"	'' 개편될 CJ QA서버 URL
    Else
        'cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' 실서버
        cjMallAPIURL = "https://api.cjmall.com/IFPAServerAction.action"			'' 실서버
    End if

    xmlSelldate = ""
	xmlSelldate = xmlSelldate &"<?xml version=""1.0"" encoding=""UTF-8""?>"
    xmlSelldate = xmlSelldate &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
    xmlSelldate = xmlSelldate &"<tns:vendorId>411378</tns:vendorId>"
    xmlSelldate = xmlSelldate &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    xmlSelldate = xmlSelldate &"<tns:contents>"
    xmlSelldate = xmlSelldate &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"  ''1 :출고, 2: 취소
    xmlSelldate = xmlSelldate &"	<tns:wbCrtDt>"&thedate&"</tns:wbCrtDt>" ''조회날짜 yyyy-mm-dd
    xmlSelldate = xmlSelldate &"</tns:contents>"
    xmlSelldate = xmlSelldate &"</tns:ifRequest>"

    rw "기간검색:"&thedate&"~"&thedate&" 상태:전체"
	'// =======================================================================
	'// 데이타 가져오기
    
    Dim SendDoc, retDoc
    Set SendDoc = server.createobject("MSXML2.DomDocument.3.0")
		SendDoc.async = False
		SendDoc.LoadXML(xmlSelldate)

    IF application("Svr_Info")="Dev" THEN
        Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
            retDoc.async = False
            retDoc.LoadXML(getTESTSamplXML)

        rw"<textarea cols=80 rows=20>"&getTESTSamplXML&"</textarea>"
    Else

        Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
        objXML.Open "POST", cjMallAPIURL, false
        objXML.setRequestHeader "Content-Type", "text/xml"
        objXML.send SendDoc

        if objXML.Status <> "200" then
            response.write "ERROR : 통신오류" & objXML.Status
            dbget.close : response.end
        end if

        Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
            retDoc.async = False
            retDoc.LoadXML(objXML.responseTEXT)

        'rw"<textarea cols=80 rows=20>"&objXML.responseTEXT&"</textarea>"
    end if 

    
	Set SendDoc = Nothing
	Set objXML = Nothing




    
    Dim paramInfo, retParamInfo, RetErr

    '''---------------------------------------------------------
    Dim errorMsg
	Dim Nodes, isErrExists, masterSubNodes, detailSubNodes, ErrNode, detailSubNodeItem

    Dim ordNo, ordGSeq, ordDSeq, ordWSeq
    Dim ordDtlCls, ordDtlClsCd, wbCrtDt
    Dim outwConfDt, delivDtm, cnclInsDtm, oldordNo
    Dim toutYn, itemCd, unitCd, itemName, unitNm, contItemCd, outwQty, orderStatus
    Dim delicoVenNm, wblNo, invoiceUpDt, outjFixedDt
    Dim shppDivDtlNm


    Set Nodes = retDoc.getElementsByTagName("ns1:errorMsg")
    If (Not (Nodes is Nothing)) Then
        For each ErrNode in Nodes
            errorMsg = Nodes.item(0).text
            isErrExists = true
            rw "["&thedate&"]"&errorMsg
        next
    end if

    if (isErrExists) then
        Exit function
    end if


        Set Nodes = retDoc.getElementsByTagName("ns1:instruction")

        If (Not (Nodes is Nothing)) Then
            'response.write "주문건수(" & obj1.length & ") " & "<br />"

            For each masterSubNodes in Nodes
            
                ordNo           = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '주문번호
                ''custDeliveryCost = masterSubNodes.getElementsByTagName("ns1:custDeliveryCost")(0).Text	'배송비
                
                ' shppSeq			= ""			'배송지시상세번호
                ' reOrderYn ="N" ''재주문여부 
                ' delayNts  =""  ''지연일수

                ' shppTypeDtlNm   = obj1.get(i).delivery.shipMethod
                ' delicoVenId     = ""	           								'택배배송사코드
                ' wblNo           = obj1.get(i).delivery.invoiceNo									'운송장번호
                ' if (shppTypeDtlNm="기타배송") then 
                '     wblNo = wblNo & obj1.get(i).delivery.shipMethodMessage					'배송방법 메세지 배송방법이 [기타배송]일 경우 입력받는 메세지
                ' end if
                ' delicoVenNm     = obj1.get(i).delivery.parcelCompany
                ' orderStatus     = obj1.get(i).delivery.shipStatus              '발주서상태 | ACCEPT/INSTRUCT/DEPARTURE/DELIVERING/FINAL_DELIVERY/NONE_TRACKING

                ' whoutCritnDt    = obj1.get(i).originShipDate	 '' 발송기한.
                ' outjFixedDt     = obj1.get(i).shipCompleteDate ''구매확정일자  - 업체직송인경우 7일후 완료된다. 정산이 안되면 업체직송으로 수정해야한다.

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:주문상품순번], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:주문상세순번], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:주문처리순번], 001

                        ''다음의 두 속성을 조합하여 문자열 구성함
                        ''1. 주문구분코드(J007)의 코드명 (주문, 취소,  교환, 교환취소, 주문(기출하))
                        ''2. 기출하(운송장M의 배송방법코드:B005의 93)
                        ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text	        ' 주문정보 - 주문구분, 주문

                        ''10=주문,20=취소,30=반품,31=반품취소,40=교환배송,41=교환배송취소,45=교환회수,46=교환회수취소
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' 주문정보 - 주문구분코드, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' 주문정보 - 지시일자, 2013-05-22+09:00

                        
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '주문정보 - 기출하구분(Y-기출하,N-정상출하), N
                        'chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '주문정보 - 채널구분, INTERNET
                        
                        'receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '주문정보 - 인수자, 채현아
                        
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '상품정보 - 판매코드, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '상품정보 - 단품코드, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '상품정보 - 판매상품명, 24K Gold 전자파차단스티커
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '상품정보 - 단품상세, ES-01 잘될꺼야
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '상품정보 - 협력사상품코드, 279751_0011

                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '상품정보 - 수량, 1.0

                        ''필수로 안넘어오는정보들.
                        outwConfDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0) Is Nothing)) Then
                            outwConfDt = detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0).Text       '주문정보 - 출고확정일자
                        end if
                        delivDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text        '주문정보 - 배송완료일
                        end if
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '주문정보 - 취소일자
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '주문정보 - 원주문번호
                        end if
                        
                        
                        delicoVenNm = ""
                        wblNo = ""
                        invoiceUpDt = ""
                        outjFixedDt = ""

                        orderStatus = "주문확인" ''"주문통보"  '' cjmall 주문확인처리를 따로 하지는 않음
                        if (outwConfDt<>"") then
                            orderStatus = "출고완료" ''출고일자
                        end if
                        if (delivDtm<>"") then ''인수등록일자.
                            orderStatus = "배송완료"
                        end if

                        'shppDivDtlNm = ""
                        shppDivDtlNm = ordDtlCls
                        if (shppDivDtlNm="주문") then shppDivDtlNm="일반출고"
                        if (cnclInsDtm<>"") then
                            shppDivDtlNm = "주문취소"
                        end if
                        if (toutYn="Y") then
                            shppDivDtlNm = CHKIIF(shppDivDtlNm<>"","/","")&"기출하"
                        end if
                        

                        bufStr = ""
                        bufStr = sellsite&"|"&ordNo
                        bufStr = bufStr &"|"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq
                        bufStr = bufStr &"|"&ordDtlCls
                        bufStr = bufStr &"|"&ordDtlClsCd
                        bufStr = bufStr &"|"&wbCrtDt
                        bufStr = bufStr &"|"&outwConfDt
                        bufStr = bufStr &"|"&delivDtm
                        bufStr = bufStr &"|"&cnclInsDtm
                        bufStr = bufStr &"|"&oldordNo
                        
                        bufStr = bufStr &"|"&toutYn
                        bufStr = bufStr &"|"&itemCd
                        bufStr = bufStr &"|"&unitCd
                        bufStr = bufStr &"|"&itemName
                        bufStr = bufStr &"|"&unitNm
                        bufStr = bufStr &"|"&contItemCd

                       

                        ' if (whoutCritnDt<>"") then
                        '     whoutCritnDt = LEFT(whoutCritnDt,4)&"-"&MID(whoutCritnDt,5,2)&"-"&RIGHT(whoutCritnDt,2)
                        ' end if


                        sqlStr = "db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Input]"
                        paramInfo = Array(Array("@RETURN_VALUE",adInteger	,adParamReturnValue	,,0) _
                            ,Array("@SellSite" 				, adVarchar		, adParamInput		, 	32, sellsite)	_
                            ,Array("@OutMallOrderSerial"	, adVarchar		, adParamInput		,	32, Trim(ordNo)) _
                            ,Array("@OrgDetailKey"			, adVarchar		, adParamInput		,   32, Trim(ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq)) _

                            ,Array("@confirmDt"				, adVarchar     , adParamInput		,	16, Trim(confirmDt)) _
                            ,Array("@shppNo"				, adVarchar		, adParamInput		,   32, Trim("")) _
                            ,Array("@shppSeq"				, adVarchar		, adParamInput		,   10, Trim("")) _
                            ,Array("@reOrderYn"				, adVarchar		, adParamInput		,    1, Trim(toutYn)) _
                            ,Array("@delayNts"			    , adInteger		, adParamInput		,     , Trim(0)) _
                            ,Array("@splVenItemId"			, adInteger		, adParamInput		,     , Trim(split(contItemCd,"_")(0))) _
                            ,Array("@outMallGoodsNo"		, adVarchar		, adParamInput		,   16, Trim(itemCd)) _
                            ,Array("@outMallOptionNo"		, adVarchar		, adParamInput		,   32, Trim(unitCd)) _
                            ,Array("@ordQty"			    , adInteger		, adParamInput		,     , Trim(outwQty)) _
                            ,Array("@shppDivDtlNm"		    , adVarchar		, adParamInput		,   20, Trim(shppDivDtlNm)) _
                            ,Array("@uitemNm"		        , adVarchar		, adParamInput		,   128, Trim(unitNm)) _
                            ,Array("@shppRsvtDt"			, adDate		, adParamInput		,	  , Trim("")) _
                            ,Array("@whoutCritnDt"			, adDate		, adParamInput		,	  , Trim("")) _
                            ,Array("@autoShortgYn"			, adVarchar		, adParamInput		,    1, Trim("")) _
                            ,Array("@outorderstatus"		, adVarchar		, adParamInput		,   30, Trim(orderStatus)) _

                            ,Array("@shppTypeDtlNm"		, adVarchar		, adParamInput		,   16, Trim("")) _
                            ,Array("@delicoVenId"		, adVarchar		, adParamInput		,   16, Trim("")) _
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
                        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
                        
                        successCnt = successCnt+1
                    next
                    set detailSubNodeItem = nothing
                end if
                
                set detailSubNodes = nothing
            next
            set masterSubNodes = nothing
        End If
    Set Nodes = nothing

    '' 주문번호 매핑.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "상세건수:"&successCnt
    rw "======================================"

	Get_CjmallOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->