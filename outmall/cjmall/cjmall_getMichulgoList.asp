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

rw "�Ϸ�"


' for k=0 to datelen-1
'     thedate=dateadd("d",-1*k,iedyyyymmdd)
'     if k<5 then
'     call Get_WMPOrderListByStatus(thedate,thedate,"NEW","�ֹ��뺸")
'     response.flush
'     end if
'     call Get_WMPOrderListByStatus(thedate,thedate,"CONFIRM","�ֹ�Ȯ��")
'     response.flush
'     call Get_WMPOrderListByStatus(thedate,thedate,"DELIVERY","���Ϸ�") 
'     response.flush

'     '' call Get_WMPOrderListByStatus(istyyyymmdd,iedyyyymmdd,"COMPLETE","��ۿϷ�")
'     '' response.flush
' next


'response.write("<script>setTimeout(alert('�Ϸ�'),1000);self.close();</script>")

function getTESTSamplXML()
Dim ret
ret=ret&"<?xml version=""1.0"" encoding=""euc-kr"" standalone=""yes""?>"
ret=ret&"<ns1:ifResponse ns1:ifId=""IF_04_01"" xmlns:ns1=""http://www.example.org/ifpa"">"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191203160279</ns1:ordNo>"
ret=ret&"        <ns1:custNm>����*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0503)6337-1280</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>650829</ns1:zipno>"
ret=ret&"            <ns1:addr_1>�泲 �뿵�� ������ �׸��� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1567-2���� ���ǵ��� 401ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0503)6337-1280</ns1:telno>"
ret=ret&"            <ns1:cellno>0503)6337-1280</ns1:cellno>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56828733</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12091758553</ns1:unitCd>"
ret=ret&"            <ns1:itemName>[�������и��͸�] OKK ���º� 4����Ʈ����ǰ���</ns1:itemName>"
ret=ret&"            <ns1:unitNm>[�������и��͸�] OKK ���º� 4��...</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>����*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1561-2886</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>446715</ns1:zipno>"
ret=ret&"            <ns1:addr_1>��� ���ν� ���ﱸ �ߵ� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1050���� ���������ھƷ����Ʈ 4307�� 1801ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1561-2886</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1561-2886</ns1:cellno>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>54205903</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12077600723</ns1:unitCd>"
ret=ret&"            <ns1:itemName>�ó׸� ���͸� ����Ʈ�ڽ� �� ���工�� 36x18cm</ns1:itemName>"
ret=ret&"            <ns1:unitNm>�ó׸� ���͸� ����Ʈ�ڽ� �� ���工��...</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>��*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0503)6337-4710</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>INTERNET</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>314080</ns1:zipno>"
ret=ret&"            <ns1:addr_1>�泲 ���ֽ� ���е� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>389���� ����e-���Ѽ��� 108�� 904ȣ ����E-���Ѽ���)</ns1:addr_2>"
ret=ret&"            <ns1:telno>0503)6337-4710</ns1:telno>"
ret=ret&"            <ns1:cellno>0503)6337-4710</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>�� ��</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>60574396</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12111573809</ns1:unitCd>"
ret=ret&"            <ns1:itemName>���γ�ī��</ns1:itemName>"
ret=ret&"            <ns1:unitNm>XL,��Ʈ</ns1:unitNm>"
ret=ret&"            <ns1:contItemCd>2588673_Z430</ns1:contItemCd>"
ret=ret&"            <ns1:wbIdNo>20000791401656</ns1:wbIdNo>"
ret=ret&"            <ns1:outwQty>1.0</ns1:outwQty>"
ret=ret&"            <ns1:realslAmt>24000.0</ns1:realslAmt>"
ret=ret&"            <ns1:outwAmt>24000.0</ns1:outwAmt>"
ret=ret&"            <ns1:delivInfo>�� ��</ns1:delivInfo>"
ret=ret&"            <ns1:ordDtm>2019-12-04 01:24:27</ns1:ordDtm>"
ret=ret&"            <ns1:custDeliveryCost>2500.0</ns1:custDeliveryCost>"
ret=ret&"            <ns1:costGroup>001</ns1:costGroup>"
ret=ret&"        </ns1:instructionDetail>"
ret=ret&"    </ns1:instruction>"
ret=ret&"    <ns1:instruction>"
ret=ret&"        <ns1:ordNo>20191204007609</ns1:ordNo>"
ret=ret&"        <ns1:custNm>�迬*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1558-5787</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>200931</ns1:zipno>"
ret=ret&"            <ns1:addr_1>���� ��õ�� ��ȭ�� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>267-1���� ��õL-Ÿ�� 1�� 619ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1558-5787</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1558-5787</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>���� �� �����տ� �����ּ���. �����մϴ�!</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>55961289</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12087167253</ns1:unitCd>"
ret=ret&"            <ns1:itemName>����� ����Ŀ�� ������� 2�ܺи� ȭ��ǰ ������</ns1:itemName>"
ret=ret&"            <ns1:unitNm>�׷���</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>����*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1567-9290</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>151778</ns1:zipno>"
ret=ret&"            <ns1:addr_1>���� ���Ǳ� ��õ�� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>1718���� ������־���Ʈ 103�� 102ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1567-9290</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1567-9290</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>����� ���տ������ּ���</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>55480625</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12084571791</ns1:unitCd>"
ret=ret&"            <ns1:itemName>[����ī��Ʈ] ����� ���� ���� ī��Ʈ 200x230</ns1:itemName>"
ret=ret&"            <ns1:unitNm>��ũ��Ʈ</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>���*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1920-3414</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>���*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>415808</ns1:zipno>"
ret=ret&"            <ns1:addr_1>��� ������ ǳ���� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>354-1���� ǳ��Ǫ������ 108�� 2401ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1920-3414</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1920-3414</ns1:cellno>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56293607</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12088910442</ns1:unitCd>"
ret=ret&"            <ns1:itemName>��Ʈ������-ȭ��Ʈ</ns1:itemName>"
ret=ret&"            <ns1:unitNm>��Ʈ������-ȭ��Ʈ</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>����*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1919-1079</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-05+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>426785</ns1:zipno>"
ret=ret&"            <ns1:addr_1>��� �Ȼ�� ��ϱ� ���ǵ� </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>444���� �ְ�1��������Ʈ 113�� 1303ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1919-1079</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1919-1079</ns1:cellno>"
ret=ret&"            <ns1:msgSpec>���ñ��� �� �ڵ��� �����ּ���. �����մϴ�.</ns1:msgSpec>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>56879787</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12091978787</ns1:unitCd>"
ret=ret&"            <ns1:itemName>���� ���̽�ũ������� ��Ʈ_2056004</ns1:itemName>"
ret=ret&"            <ns1:unitNm>���� ���̽�ũ������� ��Ʈ_(205...</ns1:unitNm>"
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
ret=ret&"        <ns1:custNm>����*</ns1:custNm>"
ret=ret&"        <ns1:custTelNo>0507)1922-4336</ns1:custTelNo>"
ret=ret&"        <ns1:custDeliveryCost>0.0</ns1:custDeliveryCost>"
ret=ret&"        <ns1:instructionDetail>"
ret=ret&"            <ns1:ordGSeq>001</ns1:ordGSeq>"
ret=ret&"            <ns1:ordDSeq>001</ns1:ordDSeq>"
ret=ret&"            <ns1:ordWSeq>001</ns1:ordWSeq>"
ret=ret&"            <ns1:ordDtlCls>�ֹ� </ns1:ordDtlCls>"
ret=ret&"            <ns1:ordDtlClsCd>10</ns1:ordDtlClsCd>"
ret=ret&"            <ns1:wbCrtDt>2019-12-04+09:00</ns1:wbCrtDt>"
ret=ret&"            <ns1:outwConfDt>2019-12-04+09:00</ns1:outwConfDt>"
ret=ret&"            <ns1:delivDtm>2019-12-06+09:00</ns1:delivDtm>"
ret=ret&"            <ns1:toutYn>N</ns1:toutYn>"
ret=ret&"            <ns1:chnNm>MOBILE</ns1:chnNm>"
ret=ret&"            <ns1:receverNm>����*</ns1:receverNm>"
ret=ret&"            <ns1:zipno>791755</ns1:zipno>"
ret=ret&"            <ns1:addr_1>��� ���׽� �ϱ� ���ﵿ </ns1:addr_1>"
ret=ret&"            <ns1:addr_2>464���� �Ѷ�Ÿ���Ǽ� 302�� 1803ȣ</ns1:addr_2>"
ret=ret&"            <ns1:telno>0507)1922-4336</ns1:telno>"
ret=ret&"            <ns1:cellno>0507)1922-4336</ns1:cellno>"
ret=ret&"            <ns1:packYn>�Ϲ�</ns1:packYn>"
ret=ret&"            <ns1:itemCd>57250966</ns1:itemCd>"
ret=ret&"            <ns1:unitCd>12093622037</ns1:unitCd>"
ret=ret&"            <ns1:itemName>���̶����� ��Ƽ ������ ����Ŀư</ns1:itemName>"
ret=ret&"            <ns1:unitNm>���̶����� ��Ƽ ������ ����Ŀư</ns1:unitNm>"
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
	'// ��¥����
	''selldate = "2017-11-10"
	

    IF application("Svr_Info")="Dev" THEN
        'cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' �׽�Ʈ����
        cjMallAPIURL = "http://210.122.101.154:8210/IFPAServerAction.action"	'' ����� CJ QA���� URL
    Else
        'cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' �Ǽ���
        cjMallAPIURL = "https://api.cjmall.com/IFPAServerAction.action"			'' �Ǽ���
    End if

    xmlSelldate = ""
	xmlSelldate = xmlSelldate &"<?xml version=""1.0"" encoding=""UTF-8""?>"
    xmlSelldate = xmlSelldate &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
    xmlSelldate = xmlSelldate &"<tns:vendorId>411378</tns:vendorId>"
    xmlSelldate = xmlSelldate &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    xmlSelldate = xmlSelldate &"<tns:contents>"
    xmlSelldate = xmlSelldate &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"  ''1 :���, 2: ���
    xmlSelldate = xmlSelldate &"	<tns:wbCrtDt>"&thedate&"</tns:wbCrtDt>" ''��ȸ��¥ yyyy-mm-dd
    xmlSelldate = xmlSelldate &"</tns:contents>"
    xmlSelldate = xmlSelldate &"</tns:ifRequest>"

    rw "�Ⱓ�˻�:"&thedate&"~"&thedate&" ����:��ü"
	'// =======================================================================
	'// ����Ÿ ��������
    
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
            response.write "ERROR : ��ſ���" & objXML.Status
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
            'response.write "�ֹ��Ǽ�(" & obj1.length & ") " & "<br />"

            For each masterSubNodes in Nodes
            
                ordNo           = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '�ֹ���ȣ
                ''custDeliveryCost = masterSubNodes.getElementsByTagName("ns1:custDeliveryCost")(0).Text	'��ۺ�
                
                ' shppSeq			= ""			'������û󼼹�ȣ
                ' reOrderYn ="N" ''���ֹ����� 
                ' delayNts  =""  ''�����ϼ�

                ' shppTypeDtlNm   = obj1.get(i).delivery.shipMethod
                ' delicoVenId     = ""	           								'�ù��ۻ��ڵ�
                ' wblNo           = obj1.get(i).delivery.invoiceNo									'������ȣ
                ' if (shppTypeDtlNm="��Ÿ���") then 
                '     wblNo = wblNo & obj1.get(i).delivery.shipMethodMessage					'��۹�� �޼��� ��۹���� [��Ÿ���]�� ��� �Է¹޴� �޼���
                ' end if
                ' delicoVenNm     = obj1.get(i).delivery.parcelCompany
                ' orderStatus     = obj1.get(i).delivery.shipStatus              '���ּ����� | ACCEPT/INSTRUCT/DEPARTURE/DELIVERING/FINAL_DELIVERY/NONE_TRACKING

                ' whoutCritnDt    = obj1.get(i).originShipDate	 '' �߼۱���.
                ' outjFixedDt     = obj1.get(i).shipCompleteDate ''����Ȯ������  - ��ü�����ΰ�� 7���� �Ϸ�ȴ�. ������ �ȵǸ� ��ü�������� �����ؾ��Ѵ�.

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:�ֹ���ǰ����], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:�ֹ��󼼼���], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:�ֹ�ó������], 001

                        ''������ �� �Ӽ��� �����Ͽ� ���ڿ� ������
                        ''1. �ֹ������ڵ�(J007)�� �ڵ�� (�ֹ�, ���,  ��ȯ, ��ȯ���, �ֹ�(������))
                        ''2. ������(�����M�� ��۹���ڵ�:B005�� 93)
                        ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text	        ' �ֹ����� - �ֹ�����, �ֹ�

                        ''10=�ֹ�,20=���,30=��ǰ,31=��ǰ���,40=��ȯ���,41=��ȯ������,45=��ȯȸ��,46=��ȯȸ�����
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' �ֹ����� - �ֹ������ڵ�, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' �ֹ����� - ��������, 2013-05-22+09:00

                        
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '�ֹ����� - �����ϱ���(Y-������,N-��������), N
                        'chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '�ֹ����� - ä�α���, INTERNET
                        
                        'receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '�ֹ����� - �μ���, ä����
                        
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '��ǰ���� - �Ǹ��ڵ�, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '��ǰ���� - ��ǰ�ڵ�, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '��ǰ���� - �ǸŻ�ǰ��, 24K Gold ���������ܽ�ƼĿ
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '��ǰ���� - ��ǰ��, ES-01 �ߵɲ���
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '��ǰ���� - ���»��ǰ�ڵ�, 279751_0011

                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '��ǰ���� - ����, 1.0

                        ''�ʼ��� �ȳѾ����������.
                        outwConfDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0) Is Nothing)) Then
                            outwConfDt = detailSubNodeItem.getElementsByTagName("ns1:outwConfDt")(0).Text       '�ֹ����� - ���Ȯ������
                        end if
                        delivDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0) Is Nothing)) Then
                            delivDtm = detailSubNodeItem.getElementsByTagName("ns1:delivDtm")(0).Text        '�ֹ����� - ��ۿϷ���
                        end if
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '�ֹ����� - �������
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '�ֹ����� - ���ֹ���ȣ
                        end if
                        
                        
                        delicoVenNm = ""
                        wblNo = ""
                        invoiceUpDt = ""
                        outjFixedDt = ""

                        orderStatus = "�ֹ�Ȯ��" ''"�ֹ��뺸"  '' cjmall �ֹ�Ȯ��ó���� ���� ������ ����
                        if (outwConfDt<>"") then
                            orderStatus = "���Ϸ�" ''�������
                        end if
                        if (delivDtm<>"") then ''�μ��������.
                            orderStatus = "��ۿϷ�"
                        end if

                        'shppDivDtlNm = ""
                        shppDivDtlNm = ordDtlCls
                        if (shppDivDtlNm="�ֹ�") then shppDivDtlNm="�Ϲ����"
                        if (cnclInsDtm<>"") then
                            shppDivDtlNm = "�ֹ����"
                        end if
                        if (toutYn="Y") then
                            shppDivDtlNm = CHKIIF(shppDivDtlNm<>"","/","")&"������"
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
                        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
                        
                        successCnt = successCnt+1
                    next
                    set detailSubNodeItem = nothing
                end if
                
                set detailSubNodes = nothing
            next
            set masterSubNodes = nothing
        End If
    Set Nodes = nothing

    '' �ֹ���ȣ ����.
    ' strSql = "exec db_temp.[dbo].[usp_TEN_xSiteTmpMichulList_Maporder] '"&sellsite&"','"&confirmDt&"'"
    ' dbget.Execute strSql

    rw "�󼼰Ǽ�:"&successCnt
    rw "======================================"

	Get_CjmallOrderListByStatus = True

end function
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->