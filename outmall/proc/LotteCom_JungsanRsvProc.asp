<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/Chilkatlib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/outmall/lotteCom/inc_dailyAuthCheck.asp"-->
<%
dim yyyymmdd : yyyymmdd=request("yyyymmdd")


Dim paramAdd 
paramAdd = "&start_date="+Replace(yyyymmdd,"-","")
paramAdd = paramAdd + "&end_date="+Replace(yyyymmdd,"-","")
paramAdd = paramAdd + "&pur_shp_cd=3"  ''3 위탁판매. (2판매분매입)

response.write lotteAPIURL & "/openapi/settleCompleteListOpenApi.lotte?subscriptionId=" & lotteAuthNo & paramAdd

Dim iResult, iMessage
Dim retXML, iSettleCount, iSettleInfo, iItmNo, SubNodes
Dim objXML, xmlDOM
    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    objXML.Open "GET", lotteAPIURL & "/openapi/settleCompleteListOpenApi.lotte?subscriptionId=" & lotteAuthNo & paramAdd, false
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send()
    If (objXML.Status) = "200" Then
        Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	
    	IF (Trim(objXML.ResponseBody)="") THEN
            rw "결과 빈값"
            iResult="-1"
        ELSE
            retXML = BinaryToText(objXML.ResponseBody, "euc-kr")
        	xmlDOM.LoadXML retXML
        	
        	iSettleCount		= Trim(xmlDOM.getElementsByTagName("SettleCount").item(0).text)			'SettleCount
        	
        	rw "<textarea cols=80 rows=30>"&retXML&"</textarea>"
        	rw "iSettleCount:"&iSettleCount
        	
        	if (iSettleCount>0) then 
        	Set iSettleInfo = xmlDOM.getElementsByTagName("SettleInfo")
        	    For each SubNodes in iSettleInfo
				    iItmNo	    = Trim(SubNodes.getElementsByTagName("ItmNo").item(0).text)
				    rw iItmNo
				Next
			Set iSettleInfo = Nothing
		    End if
        END IF	
        Set xmlDOM= Nothing
        
        
    else
        if (IsAutoScript) then
            rw "롯데닷컴과 통신중에 오류가 발생했습니다. Status : "&objXML.Status
        else    
    	    rw "롯데닷컴과 통신중에 오류가 발생했습니다. Status : "&objXML.Status
    	ENd IF
    	dbget.Close: Response.End
    end if
    Set objXML = Nothing
    
if (iResult="-1") then 
    dbget.Close: Response.End
end if

if (iSettleCount<1) then
    dbget.Close: Response.End
end if
   
'' 상세내역조회
Dim iGoods_no, iProcessDate, iOrderNo, iSettleCnt, iTotalAmt, iSupplyAmt, iVatAmt
iGoods_no = iItmNo
iSettleCount = 0
retXML = ""

    paramAdd = "&start_date="+Replace(yyyymmdd,"-","")
    paramAdd = paramAdd + "&end_date="+Replace(yyyymmdd,"-","")
    paramAdd = paramAdd + "&goods_no="&iGoods_no     ''파라메터 소문자임;;

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    objXML.Open "GET", lotteAPIURL & "/openapi/settleCompleteDetailListOpenApi.lotte?subscriptionId=" & lotteAuthNo & paramAdd, false
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send()
    If (objXML.Status) = "200" Then
        Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	
    	IF (Trim(objXML.ResponseBody)="") THEN
            rw "결과 빈값"
            iResult="-1"
        ELSE
            retXML = BinaryToText(objXML.ResponseBody, "euc-kr")
        	xmlDOM.LoadXML retXML
        	
        	iSettleCount		= Trim(xmlDOM.getElementsByTagName("SettleCount").item(0).text)			'SettleCount
        	
        	rw "<textarea cols=80 rows=30>"&retXML&"</textarea>"
        	rw "iSettleCount:"&iSettleCount
        	
        	if (iSettleCount>0) then 
        	Set iSettleInfo = xmlDOM.getElementsByTagName("SettleDetailInfo")
        	    For each SubNodes in iSettleInfo
				    iProcessDate	= Trim(SubNodes.getElementsByTagName("ProcessDate").item(0).text)
				    iOrderNo	    = Trim(SubNodes.getElementsByTagName("OrderNo").item(0).text)
				    iSettleCnt      = Trim(SubNodes.getElementsByTagName("SettleCnt").item(0).text)
				    iTotalAmt        = Trim(SubNodes.getElementsByTagName("TotalAmt").item(0).text)
				    iSupplyAmt       = Trim(SubNodes.getElementsByTagName("SupplyAmt").item(0).text)
				    iVatAmt          = Trim(SubNodes.getElementsByTagName("VatAmt").item(0).text)
				    
				    rw iGoods_no&"|"&iProcessDate&"|"&iOrderNo&"|"&iSettleCnt&"|"&iTotalAmt&"|"&iSupplyAmt&"|"&iVatAmt
				Next
			Set iSettleInfo = Nothing
		    End if
        END IF	
        Set xmlDOM= Nothing
        
        
    else
        if (IsAutoScript) then
            rw "롯데닷컴과 통신중에 오류가 발생했습니다. Status : "&objXML.Status
        else    
    	    rw "롯데닷컴과 통신중에 오류가 발생했습니다. Status : "&objXML.Status
    	ENd IF
    	dbget.Close: Response.End
    end if
    Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->