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
		'call getSsgJungsanList(istyyyymmdd,"7007") ''사이트 구분이 없어진듯함.
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
					''법인변경 12/27
					if (siteNo="7006") then siteNo="6006"
					if (siteNo="7007") then siteNo="6007"

    			    settleDt         = LagrgeNode(i).SelectSingleNode("settleDt").Text               ''*정산 매출 일자
    			    orordNo          = LagrgeNode(i).SelectSingleNode("orordNo").Text                ''원주문번호
    			    ordNo            = LagrgeNode(i).SelectSingleNode("ordNo").Text                 ''*주문번호 [20171127616023]
    			    ordItemSeq       = LagrgeNode(i).SelectSingleNode("ordItemSeq").Text            ''*주문순번? [1]
					o_ordNo			= ordNo
					o_ordItemSeq    = ordItemSeq
    			    txnDivNm     = LagrgeNode(i).SelectSingleNode("txnDivNm").Text                ''과세구분코드명 과세/면세
    			    itemId          = LagrgeNode(i).SelectSingleNode("itemId").Text                  ''상품번호  [1000024811163]
    			    itemNm          = LagrgeNode(i).SelectSingleNode("itemNm").Text                  ''상품명    [문주란스티커]
    			    uitemId      = LagrgeNode(i).SelectSingleNode("uitemId").Text                   '' 단품 ID

    			    salesQty       = LagrgeNode(i).SelectSingleNode("salesQty").Text             ''    매출 수량
    			    settleAmt      = LagrgeNode(i).SelectSingleNode("settleAmt").Text            '' 정산 금액(vat 제외)
					settleVat	   = LagrgeNode(i).SelectSingleNode("settleVat").Text			 '' 정산 금액 Vat (2020/03/17)
    			    salesAmt      = LagrgeNode(i).SelectSingleNode("salesAmt").Text              '' 매출 금액(vat 포함)
    			    splvenBdnDcAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnDcAmt").Text    ''업체 할인 부담 금액
    			    owncoBdnDcAmt    = LagrgeNode(i).SelectSingleNode("owncoBdnDcAmt").Text      ''자사(SSG) 할인 부담 금액
    			    netAmt    = LagrgeNode(i).SelectSingleNode("netAmt").Text            		''순 판매 금액(vat 포함)
    			    mrgrt    = LagrgeNode(i).SelectSingleNode("mrgrt").Text            			''마진율
    			    dvShppcstAmt    = LagrgeNode(i).SelectSingleNode("dvShppcstAmt").Text            ''배송 금액(vat 제외)
    			    dvShppcstVat    = LagrgeNode(i).SelectSingleNode("dvShppcstVat").Text            ''배송 VAT 금액
    			    custBdnShppAmt    = LagrgeNode(i).SelectSingleNode("custBdnShppAmt").Text            ''고객 부담 배송 금액
    			    splvenBdnShppAmt    = LagrgeNode(i).SelectSingleNode("splvenBdnShppAmt").Text            ''업체 부담 배송 금액

					If yyyymmdd="20220921" Then
						shppNo	= LagrgeNode(i).SelectSingleNode("shppNo").Text            ''배송ID
					End If


    			    dvShppcstTTL = CLNG(dvShppcstAmt)+CLNG(dvShppcstVat)
    			    isDlvPayExists = (dvShppcstTTL<>0)
    			    extVatYN = "Y"
                    if (txnDivNm="면세") then extVatYN="N"


					''o_settleAmt = settleAmt
					settleAmt = settleAmt*1 + settleVat*1

					' if (ordNo = "20180816756677") then
					'    	rw "salesQty:"&salesQty&"/"&"settleAmt:"&settleAmt&"/"&"salesAmt:"&salesAmt&"/"&"splvenBdnDcAmt:"&splvenBdnDcAmt&"/"&"owncoBdnDcAmt:"&owncoBdnDcAmt&"/"&"netAmt:"&netAmt&"/"
					' end if
					divideNo = salesQty
					if (divideNo=0) then divideNo=1

					'' 정산금액 = (settleAmt-협력사 할인부담금을 깐다)/ 수량 2020/03/17 , 면세는 먼가 이상하다.
					settleAmt = settleAmt-splvenBdnDcAmt

                    ' if (extVatYN = "Y") then
    			    '     ''settleAmt = CLNG(settleAmt*1.1)  ''2018/03/06 추가
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
					' 	'' settleAmt 정산금액VAT제외.
					' 	'' netAmt    순판매금액 VAT포함.
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

    			    if (salesQty<0) then  ''반품인경우.
						' if (ordNo=p_ordNo) and (ordItemSeq=p_ordItemSeq) then ''중복일경우.
						' 	ordNo = ordNo&"-"&ordItemSeq&"-"&ordItemSeq
						' else
    			        	ordNo = ordNo&"-"&ordItemSeq
						' end if
    			    end if

    			    'if NOT (LagrgeNode(i).SelectSingleNode("ordMemoCntt") is Nothing) then
    			    '    ordMemoCntt     = replace(LagrgeNode(i).SelectSingleNode("ordMemoCntt").Text,"[고객배송메모]","")                 ''[[고객배송메모]배송메세지]
    			    'end if


                    ''extCommPrice = salesAmt-settleAmt
					extCommPrice = netAmt-settleAmt

					if (extVatYN="N") then

						extCommPrice = extCommPrice + ROUND((netAmt-settleAmt)*0.1+0.0001)

						settleAmt    = settleAmt    - ROUND((netAmt-settleAmt)*0.1+0.0001)
					end if

					'2022-02-03 김진영..특이케이스
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
                        '' 반품배송비
                    elseif ((NOT isDlvPayExists) and (salesQty=0) and (salesAmt=0) and (settleAmt=0) and (netAmt=0))then
                        '' 재낌.
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

						'2021-12-01 김진영..특이케이스
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202111025555BF" and yyyymmdd="20211108" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-02-03 김진영..특이케이스
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202201099FB90E" and yyyymmdd="20220120" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-03-02 김진영..특이케이스
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220125E6C095" and yyyymmdd="20220209" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-04-01 김진영..특이케이스
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220313C4B704" and yyyymmdd="20220318" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-6-02 김진영..특이케이스
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "202204197BD0BF" and yyyymmdd="20220518" Then
							ordItemSeq = ordItemSeq & "-1"
						End If
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220510D78065" and yyyymmdd="20220523" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2022-09-22 김진영..특이케이스..shppNo로 밖에 구분 안 됨
						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220907D0E9BF" and yyyymmdd="20220921" and shppNo = "D2460953857" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						If LagrgeNode(i).SelectSingleNode("ordNo").Text = "20220917F40E03" and yyyymmdd="20220928" and LagrgeNode(i).SelectSingleNode("ordItemSeq").Text = "2" Then
							ordItemSeq = ordItemSeq & "-1"
						End If

						'2023-01-02 김진영..특이케이스
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

    				if (isDlvPayExists) then ''배송비
    				    itemNm = "배송비"
    				    ordItemSeq = ordItemSeq+"-D"

    				    if (salesQty=0) then
    				        salesQty = 1
    				        itemNm = "반품비"
    				        ordItemSeq = ordItemSeq+"D"

							if (dvShppcstTTL<0) then ordItemSeq = ordItemSeq+"D"
							''20180807467429 CASE 수수료가 있을 수 있음.
							'owncoBdnDcAmt = 0
    				    	'splvenBdnDcAmt = 0
							extCommPrice = owncoBdnDcAmt*-1
    				    elseif (salesQty<0) then
    				        salesQty = -1
    				        itemNm = "반품비"

							owncoBdnDcAmt = 0
    				    	splvenBdnDcAmt = 0
							extCommPrice = 0
    				    else
    				        salesQty = 1
    				        itemNm = "배송비"

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
							settleAmt = dvShppcstTTL  ''정산액.
						End If


    				    uitemId = 0
    				    ''extVatYN = "Y"  ''상품을따라가자 SSG어드민과 비교를 쉽게하기위해


    				    '' SSG는 배송비도 안분하는듯 동일한 내역이 있으면 더하자.
    				    sqlStr = " select count(*) CNT from db_temp.dbo.tbl_xSite_JungsanTmp where sellsite='ssg' and siteNo='"&siteNo&"' and extOrderserial='"&ordNo&"' and LEFT(extOrderserSeq,LEN('"&ordItemSeq&"'))='" &ordItemSeq&"'" &vbCRLF
    				    rsget.CursorLocation = adUseClient
                        rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
                        if Not rsget.Eof then
                            if (rsget("CNT")>0) then
                                ordItemSeq = ordItemSeq + "-"&rsget("CNT")
                            end if
                        end if
                        rsget.Close

    				    '' SSG는 배송비도 안분하는듯 동일한 내역이 있으면 더하자.
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


						''수량이 -1 이고 매출액은 +인경우가 있음.(배송비) 반품배송비?

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