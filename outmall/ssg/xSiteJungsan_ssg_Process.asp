<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
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
istyyyymmdd = request("yyyymmdd")
''istyyyymmdd = "20180301"
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)

if (sellsite="ssg6006") then
	if (istyyyymmdd>="20181227") then
		call getSsgJungsanList(istyyyymmdd,"7006")
	else
		call getSsgJungsanList(istyyyymmdd,"6006")
	end if
elseif (sellsite="ssg6007") then
	if (istyyyymmdd>="20181227") then
		call getSsgJungsanList(istyyyymmdd,"7007")
	else
		call getSsgJungsanList(istyyyymmdd,"6007")
	end if
else
	if (istyyyymmdd>="20181227") then
		call getSsgJungsanList(istyyyymmdd,"7006")
		'call getSsgJungsanList(istyyyymmdd,"7007") ''����Ʈ ������ ����������.
	else
		call getSsgJungsanList(istyyyymmdd,"6006")
		call getSsgJungsanList(istyyyymmdd,"6007")
	end if
end if

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
	objXML.setTimeouts 5000,90000,90000,90000
	objXML.send()

'  	response.write "<textarea cols=100 rows=30>"&objXML.responseText&"</textarea>"
'  exit function
    Dim siteNo,settleDt,orordNo,ordNo,ordItemSeq
    Dim txnDivNm,itemId,itemNm,uitemId,salesQty
    Dim settleAmt,settleVat,salesAmt ,splvenBdnDcAmt , o_settleAmt
    Dim owncoBdnDcAmt,netAmt ,mrgrt ,dvShppcstAmt ,dvShppcstVat , extTenMeachulPrice
    Dim custBdnShppAmt,splvenBdnShppAmt ,txnDivCd
    Dim extCommPrice, extVatYN, dvShppcstTTL
    Dim isDataExists : isDataExists = FALSE
    Dim isDlvPayExists
	Dim p_ordNo, p_ordItemSeq, o_ordNo, o_ordItemSeq, divideNo, shppNo
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
					''���κ��� 12/27
					if (siteNo="7006") then siteNo="6006"
					if (siteNo="7007") then siteNo="6007"

    			    settleDt         = LagrgeNode(i).SelectSingleNode("settleDt").Text               ''*���� ���� ����
    			    orordNo          = LagrgeNode(i).SelectSingleNode("orordNo").Text                ''���ֹ���ȣ
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*�ֹ���ȣ [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*�ֹ�����? [1]
					o_ordNo			= ordNo
					o_ordItemSeq    = ordItemSeq
    			    txnDivNm     = LagrgeNode(i).SelectSingleNode("txnDivNm").Text                ''���������ڵ�� ����/�鼼
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''��ǰ��ȣ  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''��ǰ��    [���ֶ���ƼĿ]
    			    uitemId      = LagrgeNode(i).SelectSingleNode("uitemId").Text                   '' ��ǰ ID

    			    salesQty       = LagrgeNode(i).SelectSingleNode("salesQty").Text             ''    ���� ����
    			    settleAmt      = LagrgeNode(i).SelectSingleNode("settleAmt").Text            '' ���� �ݾ�(vat ����)
					settleVat	   = LagrgeNode(i).SelectSingleNode("settleVat").Text			 '' ���� �ݾ� Vat (2020/03/17)
    			    salesAmt      = LagrgeNode(i).SelectSingleNode("salesAmt").Text              '' ���� �ݾ�(vat ����)
    			    splvenBdnDcAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnDcAmt").Text    ''��ü ���� �δ� �ݾ�
    			    owncoBdnDcAmt    = LagrgeNode(i).SelectSingleNode("owncoBdnDcAmt").Text      ''�ڻ�(SSG) ���� �δ� �ݾ�
    			    netAmt    = LagrgeNode(i).SelectSingleNode("netAmt").Text            		''�� �Ǹ� �ݾ�(vat ����)
    			    mrgrt    = LagrgeNode(i).SelectSingleNode("mrgrt").Text            			''������
    			    dvShppcstAmt    = LagrgeNode(i).SelectSingleNode("dvShppcstAmt").Text            ''��� �ݾ�(vat ����)
    			    dvShppcstVat    = LagrgeNode(i).SelectSingleNode("dvShppcstVat").Text            ''��� VAT �ݾ�
    			    custBdnShppAmt    = LagrgeNode(i).SelectSingleNode("custBdnShppAmt").Text            ''�� �δ� ��� �ݾ�
    			    splvenBdnShppAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnShppAmt").Text            ''��ü �δ� ��� �ݾ�

					If yyyymmdd="20220921" Then
						shppNo	= LagrgeNode(i).SelectSingleNode("shppNo").Text            ''���ID
					End If


    			    dvShppcstTTL = CLNG(dvShppcstAmt)+CLNG(dvShppcstVat)
    			    isDlvPayExists = (dvShppcstTTL<>0)
    			    extVatYN = "Y"
                    if (txnDivNm="�鼼") then extVatYN="N"


					''o_settleAmt = settleAmt
					settleAmt = settleAmt*1 + settleVat*1

					' if (ordNo = "20180816756677") then
					'    	rw "salesQty:"&salesQty&"/"&"settleAmt:"&settleAmt&"/"&"salesAmt:"&salesAmt&"/"&"splvenBdnDcAmt:"&splvenBdnDcAmt&"/"&"owncoBdnDcAmt:"&owncoBdnDcAmt&"/"&"netAmt:"&netAmt&"/"
					' end if
					divideNo = salesQty
					if (divideNo=0) then divideNo=1

					'' ����ݾ� = (settleAmt-���»� ���κδ���� ���)/ ���� 2020/03/17 , �鼼�� �հ� �̻��ϴ�.
					settleAmt = settleAmt-splvenBdnDcAmt

                    ' if (extVatYN = "Y") then
    			    '     ''settleAmt = CLNG(settleAmt*1.1)  ''2018/03/06 �߰�
					' 	if (settleAmt<0) then
					' 		settleAmt = ROUND(ABS(settleAmt/divideNo)*1.1+0.0001,0)*divideNo
					' 	else
					' 		settleAmt = ROUND(settleAmt/divideNo*1.1+0.0001,0)*divideNo
					' 	end if
					' 	' if (settleAmt<>ROUND(o_settleAmt*1.1,0)) then
					' 	' 	rw settleAmt&"::"&ROUND(o_settleAmt*1.1,0)
					' 	' end if
					' else
					' 	'settleAmt = settleAmt-ROUND((netAmt-settleAmt)*0.1,0)
					' 	'settleAmt = ROUND(settleAmt-(netAmt-settleAmt)*0.1,0)
					' 	'' settleAmt ����ݾ�VAT����.
					' 	'' netAmt    ���Ǹűݾ� VAT����.
					' 	if (ABS(netAmt-settleAmt)<0) then
					' 		settleAmt = settleAmt-ROUND(ABS(netAmt-settleAmt)/divideNo*0.1+0.0001,0)*divideNo
					' 	else
					' 		settleAmt = settleAmt-ROUND((netAmt-settleAmt)/divideNo*0.1+0.0001,0)*divideNo
					' 	end if

					' 	if (isDlvPayExists) and (isiteno="6006" or isiteno="7006") then settleAmt=settleAmt+250   ''20180806160515 CASE
    			    ' end if



    			    if (isDlvPayExists) then
    			        salesAmt = salesAmt - dvShppcstTTL
    			        settleAmt = settleAmt - dvShppcstTTL
    			        netAmt = netAmt - dvShppcstTTL
    			    end if

    			    if (salesQty<0) then  ''��ǰ�ΰ��.
						' if (ordNo=p_ordNo) and (ordItemSeq=p_ordItemSeq) then ''�ߺ��ϰ��.
						' 	ordNo = ordNo&"-"&ordItemSeq&"-"&ordItemSeq
						' else
    			        	ordNo = ordNo&"-"&ordItemSeq
						' end if
    			    end if

    			    'if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			    '    ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[����۸޸�]","")                 ''[[����۸޸�]��۸޼���]
    			    'end if


                    ''extCommPrice = salesAmt-settleAmt
					extCommPrice = netAmt-settleAmt

					if (extVatYN="N") then

						extCommPrice = extCommPrice + ROUND((netAmt-settleAmt)*0.1+0.0001)

						settleAmt    = settleAmt    - ROUND((netAmt-settleAmt)*0.1+0.0001)
					end if

					'2022-02-03 ������..Ư�����̽�
					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202201017588FD" and yyyymmdd="20220121" and LagrgeNode(i).SelectSingleNode("ordItemSeq").Text = "2" Then
						salesQty = 1
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220125E6C095" and yyyymmdd="20220209" Then
						salesQty = -7
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220917F40E03" and yyyymmdd="20220928" and LagrgeNode(i).SelectSingleNode("ordItemSeq").Text = "2" Then
						salesQty = -2
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20221113E99C78" and yyyymmdd="20221202" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-909" Then
						salesQty = -1
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022122084FB7D" and yyyymmdd="20221225" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "2727" Then
						salesQty = 1
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022120650DBCA" and yyyymmdd="20221215" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-727" Then
						salesQty = -1
					End If

					If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20230319E40F17" and yyyymmdd="20230405" Then
						salesQty = -1
					End If

                    if (orordNo=ordNo) then orordNo=""

                    if (isDlvPayExists and (salesQty=0) ) then
                        '' ��ǰ��ۺ�
                    elseif ((NOT isDlvPayExists) and (salesQty=0) and (salesAmt=0) and (settleAmt=0) and (netAmt=0))then
                        '' �糦.
						''rw "netAmt:"&netAmt
                    else
                        if salesQty=0 then
							salesQty=1
							ordItemSeq = ordItemSeq & "-" & replace(settleDt,"-","")
						end if
						'netAmt = CLNG(netAmt/(salesQty))							'' extReducedPrice

						salesAmt = CLNG(salesAmt/(salesQty)*100)/100				'' extItemCost
						owncoBdnDcAmt = CLNG(owncoBdnDcAmt/(salesQty)*100)/100		'' extOwnCouponPrice
						splvenBdnDcAmt = CLNG(splvenBdnDcAmt/(salesQty)*100)/100	'' extTenCouponPrice
						'on Error Resume Next
						extCommPrice = CLNG(extCommPrice/(salesQty)*100)/100		'' extCommPrice
						' if Err then
						' 	rw extCommPrice
						' 	rw salesQty
						' 	response.end
						' end if
						settleAmt = CLNG(settleAmt/(salesQty)*100)/100        		 ''extTenJungsanPrice

						extTenMeachulPrice = salesAmt-owncoBdnDcAmt-splvenBdnDcAmt
						netAmt = CLNG(extTenMeachulPrice)
						extCommPrice = extTenMeachulPrice-settleAmt

						'2021-12-01 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202111025555BF" and yyyymmdd="20211108" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-02-03 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202201099FB90E" and yyyymmdd="20220120" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-03-02 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220125E6C095" and yyyymmdd="20220209" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-04-01 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220313C4B704" and yyyymmdd="20220318" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-6-02 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202204197BD0BF" and yyyymmdd="20220518" Then
							ordItemSeq = ordItemSeq & "-1"
						End If
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220510D78065" and yyyymmdd="20220523" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-09-22 ������..Ư�����̽�..shppNo�� �ۿ� ���� �� ��
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220907D0E9BF" and yyyymmdd="20220921" and shppNo = "D2460953857" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220917F40E03" and yyyymmdd="20220928" and LagrgeNode(i).SelectSingleNode("ordItemSeq").Text = "2" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2023-01-02 ������..Ư�����̽�
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20221113E99C78" and yyyymmdd="20221202" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-909" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022112424407D" and yyyymmdd="20221212" Then
							ordItemSeq = ordItemSeq & "-2"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022122084FB7D" and yyyymmdd="20221225" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "2727" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022120650DBCA" and yyyymmdd="20221215" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-727" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20221211620529" and yyyymmdd="20221215" and LagrgeNode(i).SelectSingleNode("ordItemSeq").Text = "3" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022121570C810" and yyyymmdd="20230106" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-16484" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022120448FEDA" and yyyymmdd="20230209" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20230319E40F17" and yyyymmdd="20230405" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20230613423FD7" and yyyymmdd="20230616" and LagrgeNode(i).SelectSingleNode("settleAmt").Text = "-1696" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						if (p_ordNo=o_ordNo) and (p_ordItemSeq = o_ordItemSeq) then  '' 20180719769649 CASE
							sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and siteNo='"&siteNo&"' and extOrderserial='"&ordNo&"' and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"'" &vbCRLF
							''sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and siteNo='"&siteNo&"' and extOrderserial='"&ordNo&"' and extOrderserSeq='" &ordItemSeq&"'" &vbCRLF
							rsget.CursorLocation = adUseClient
							rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
							if Not rsget.Eof then
								if (rsget("CNT")>0) then
									ordItemSeq = ordItemSeq + "-"&rsget("CNT")
								end if
							end if
							rsget.Close

							' if ordItemSeq="26-20181119-1" then
							' 	if (netAmt<>"72") then
							' 	ordItemSeq = "26-20181119-2"
							' 	end if
							' end if
						end if

        			    sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
        			    sqlStr = sqlStr + " (sellsite, siteNo, extOrderserial, extOrderserSeq, extOrgOrderserial, extItemNo, extItemCost"
        			    sqlStr = sqlStr + " , extReducedPrice, extOwnCouponPrice, extTenCouponPrice, extJungsanType, extCommPrice, extTenMeachulPrice"
        			    sqlStr = sqlStr + " , extTenJungsanPrice, extMeachulDate, extJungsanDate"
        			    sqlStr = sqlStr + " , extItemName, extItemOptionName, extVatYN, extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice,extItemID,extItemOption) "
        				sqlStr = sqlStr + " values('ssg', '"+Cstr(siteNo)+"', '" + CStr(ordNo) + "', '" + CStr(ordItemSeq)
        				sqlStr = sqlStr + "', '" + CStr(orordNo) + "', '" + CStr(salesQty) + "', '" + CStr(salesAmt)
        				sqlStr = sqlStr + "', '" + CStr(netAmt) + "', '" + CStr(owncoBdnDcAmt) + "', '" + CStr(splvenBdnDcAmt)
        				sqlStr = sqlStr + "', '" + CStr("C") + "', '" + CStr(extCommPrice) + "', '" + CStr(extTenMeachulPrice)
        				sqlStr = sqlStr + "', '" + CStr(settleAmt) + "', '" + CStr(settleDt) + "', '" + CStr("")
        				sqlStr = sqlStr + "', '" + CStr(itemNm) + "', convert(varchar(128),'" & CStr("") & "'), '" + CStr(extVatYN)
        				sqlStr = sqlStr + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "','"&itemId&"','"&uitemId&"') "
        				'  if (extVatYN <> "Y") then
						'   	rw i&"-"&sqlStr
						'  end if
        				dbget.execute sqlStr

						p_ordNo			= o_ordNo
						p_ordItemSeq	= o_ordItemSeq
    				end if

    				if (isDlvPayExists) then ''��ۺ�
    				    itemNm = "��ۺ�"
    				    ordItemSeq = ordItemSeq+"-D"

    				    if (salesQty=0) then
    				        salesQty = 1
    				        itemNm = "��ǰ��"
    				        ordItemSeq = ordItemSeq+"D"

							if (dvShppcstTTL<0) then ordItemSeq = ordItemSeq+"D"
							''20180807467429 CASE �����ᰡ ���� �� ����.
							'owncoBdnDcAmt = 0
    				    	'splvenBdnDcAmt = 0
							extCommPrice = owncoBdnDcAmt*-1
    				    elseif (salesQty<0) then
    				        salesQty = -1
    				        itemNm = "��ǰ��"

							owncoBdnDcAmt = 0
    				    	splvenBdnDcAmt = 0
							extCommPrice = 0
    				    else
    				        salesQty = 1
    				        itemNm = "��ۺ�"

							owncoBdnDcAmt = 0
    				    	splvenBdnDcAmt = 0
							extCommPrice = 0
    				    end if


						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022101776F9B0" AND LagrgeNode(i).SelectSingleNode("salesQty").Text = "0" Then
							salesAmt = "-103"
							netAmt   = "47"
							settleAmt = "-28"
						ElseIf LagrgeNode(i).SelectSingleNode("ordNo").Text = "2022101469AC22" AND LagrgeNode(i).SelectSingleNode("salesQty").Text = "0" Then
							salesAmt = "-103"
							netAmt   = "47"
							settleAmt = "-28"
						Else
							salesAmt = dvShppcstTTL
							netAmt   = dvShppcstTTL - owncoBdnDcAmt
							settleAmt = dvShppcstTTL  ''�����.
						End If


    				    uitemId = 0
    				    ''extVatYN = "Y"  ''��ǰ�������� SSG���ΰ� �񱳸� �����ϱ�����


    				    '' SSG�� ��ۺ� �Ⱥ��ϴµ� ������ ������ ������ ������.
    				    sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and siteNo='"&siteNo&"' and extOrderserial='"&ordNo&"' and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"'" &vbCRLF
    				    rsget.CursorLocation = adUseClient
                        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                        if Not rsget.Eof then
                            if (rsget("CNT")>0) then
                                ordItemSeq = ordItemSeq + "-"&rsget("CNT")
                            end if
                        end if
                        rsget.Close

    				    '' SSG�� ��ۺ� �Ⱥ��ϴµ� ������ ������ ������ ������.
    				    sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and siteNo='"&siteNo&"' and extOrderserial='"&ordNo&"' and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"'" &vbCRLF
    				    rsget.CursorLocation = adUseClient
                        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                        if Not rsget.Eof then
                            if (rsget("CNT")>0) then
                                ordItemSeq = ordItemSeq + "-"&rsget("CNT")
                            end if
                        end if
                        rsget.Close

    				    sqlStr = ""
						sqlStr = sqlStr & " select count(*) CNT "
						sqlStr = sqlStr & " from db_jungsan.dbo.tbl_xsite_jungsandata "
						sqlStr = sqlStr & " where sellsite = 'ssg' "
						sqlStr = sqlStr & " and extOrderserial = '"&ordNo&"' "
						sqlStr = sqlStr & " and extMeachulDate < '"& CStr(settleDt) &"' "
						sqlStr = sqlStr & " and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"' "
    				    rsget.CursorLocation = adUseClient
                        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                        if Not rsget.Eof then
                            if (rsget("CNT")>0) then
                                ordItemSeq = ordItemSeq + "-"&rsget("CNT")
                            end if
                        end if
                        rsget.Close


						''������ -1 �̰� ������� +�ΰ�찡 ����.(��ۺ�) ��ǰ��ۺ�?

    				    sqlStr = " insert into db_temp.dbo.tbl_xSite_JungsanTmp"
        			    sqlStr = sqlStr + " (sellsite, siteNo, extOrderserial, extOrderserSeq"
						sqlStr = sqlStr + " , extOrgOrderserial, extItemNo, extItemCost"
        			    sqlStr = sqlStr + " , extReducedPrice, extOwnCouponPrice, extTenCouponPrice"
						sqlStr = sqlStr + " , extJungsanType, extCommPrice, extTenMeachulPrice"
        			    sqlStr = sqlStr + " , extTenJungsanPrice, extMeachulDate, extJungsanDate"
        			    sqlStr = sqlStr + " , extItemName, extItemOptionName, extVatYN"
						sqlStr = sqlStr + " , extCommSupplyPrice, extCommSupplyVatPrice, extTenMeachulSupplyPrice, extTenMeachulSupplyVatPrice,extItemID,extItemOption) "
        				sqlStr = sqlStr + " values('ssg', '"+Cstr(siteNo)+"', '" + CStr(ordNo) + "', '" + CStr(ordItemSeq)
        				sqlStr = sqlStr + "', '" + CStr(orordNo) + "', '" + CStr(salesQty) + "', '" + CStr(CLNG(salesAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(netAmt/salesQty)) + "', '" + CStr(CLNG(owncoBdnDcAmt/salesQty)) + "', '" + CStr(CLNG(splvenBdnDcAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr("D") + "', '" + CStr(CLNG(extCommPrice/(salesQty))) + "', '" + CStr(CLNG(netAmt/(salesQty)))
        				sqlStr = sqlStr + "', '" + CStr(CLNG(settleAmt/(salesQty))) + "', '" + CStr(settleDt) + "', '" + CStr("")
        				sqlStr = sqlStr + "', '" + CStr(itemNm) + "', convert(varchar(128),'" & CStr(uitemId) & "'), '" + CStr(extVatYN)
        				sqlStr = sqlStr + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "', '" + CStr(0) + "','"&itemId&"','"&uitemId&"') "

						' if (extVatYN <> "Y") then
						'  	rw  i&"-"&sqlStr
						' end if
        				dbget.execute sqlStr

    				end if
			    Next

			End If
			Set LagrgeNode = nothing
		Set xmlDOM = nothing

	''response.end

	    if (isDataExists) then
	        sqlStr = " exec db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ssg]"
	        dbget.execute sqlStr
	    end if


	SET objXML=Nothing

end function

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->