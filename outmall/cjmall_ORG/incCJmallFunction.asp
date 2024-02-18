<%
Dim isCJ_DebugMode : isCJ_DebugMode = True
Dim cjMallAPIURL, ret1, retErrStr

IF application("Svr_Info")="Dev" THEN
	'cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' 테스트서버
	cjMallAPIURL = "http://210.122.101.154:8210/IFPAServerAction.action"	'' 개편될 CJ QA서버 URL
Else
	cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' 실서버
End if

Function getXMLString(byval iitemid, mode, paramA)
	Dim oCJMallItem
	Dim strRst, bufRET, buf1, notitemId, notmakerid

	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectMode = mode
		oCJMallItem.FRectItemID = iitemid
	IF (mode="ORDLIST") or (mode="ORDCANCELLIST") or (mode="ORDLISTUP") then
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst &"<tns:contents>"
        IF (mode="ORDLIST") or (mode="ORDLISTUP") THEN
            strRst = strRst &"	<tns:instructionCls>"&"1"&"</tns:instructionCls>"
        ELSEIF (mode="ORDCANCELLIST") then
            strRst = strRst &"	<tns:instructionCls>"&"2"&"</tns:instructionCls>"
        END IF
        strRst = strRst &"	<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>" ''조회날짜
        strRst = strRst &"</tns:contents>"
        strRst = strRst &"</tns:ifRequest>"
        getXMLString = strRst
    ELSEIF (mode="CSLIST") then
		'// CS내역일 경우 iitemid 는 날짜이다.
        strRst="<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst=strRst&"<tns:ifRequest tns:ifId=""IF_04_02"" xmlns:tns=""http://www.example.org/ifpa"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_02.xsd "">"
        strRst=strRst&"<tns:vendorId>411378</tns:vendorId>"
        strRst=strRst&"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst=strRst&"<tns:contents>"
        strRst=strRst&"<tns:wbCrtDt>"&iitemid&"</tns:wbCrtDt>"
        strRst=strRst&"<tns:autoFlag></tns:autoFlag>" ''조회조건 - 자동회수확정여부 Enum(""=전체, 0=N, 1=Y)
        strRst=strRst&"</tns:contents>"
        strRst=strRst&"</tns:ifRequest>"

        getXMLString = strRst
	ELSEIF (mode="commonCD") then
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_02_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_02_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
        strRst = strRst &"<tns:lgrpCd>"&iitemid&"</tns:lgrpCd>"
        strRst = strRst &"</tns:ifRequest>"
        getXMLString = strRst
	End If
	SET oCJMallItem = Nothing
End Function

Function getCjMallPrdNoByItemID(byval iitemid)
	Dim ret
	Dim sqlStr
	If iitemid = "" Then Exit Function
	sqlStr = " select isNULL(cjmallprdno,'') as cjmallprdno from db_outmall.dbo.tbl_cjmall_regitem where itemid="&iitemid
	rsCTget.Open sqlStr, dbCTget, 1
	If Not(rsCTget.EOF or rsCTget.BOF) Then
		ret = rsCTget("cjmallprdno")
	End If
	rsCTget.close
	getCjMallPrdNoByItemID = ret
End Function

function getCjMallfirstItemoption(byval iitemid)
    dim ret
    dim sqlStr, iOption

    if iitemid="" then Exit function

    sqlStr = " select top 1 itemoption from db_outmall.dbo.tbl_OutMall_regedoption"
    sqlStr = sqlStr & " where mallid='"&CMALLNAME&"'"
    sqlStr = sqlStr & " and itemid="&iitemid

    rsCTget.Open sqlStr, dbCTget
	If Not(rsCTget.EOF or rsCTget.BOF) Then
		ret = rsCTget("itemoption")
	End If
	rsCTget.close

	if (ret="") then
		sqlStr = "select top 1 itemoption from db_AppWish.dbo.tbl_item_option where itemid = '"&iitemid&"' and isusing = 'Y' and optsellyn = 'Y' order by itemoption asc"
		rsCTget.Open sqlStr, dbCTget
		If Not(rsCTget.EOF or rsCTget.BOF) Then
			ret = rsCTget("itemoption")
		Else
			ret = "0000"
		End If
		rsCTget.close
	end if
	getCjMallfirstItemoption = ret
end function

Function getOriginName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 areacode, areaName" & VBCRLF
	sqlStr = sqlStr & " FROM db_outmall.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
	sqlStr = sqlStr & " WHERE areaName='"&iname&"'"
	rsCTget.Open sqlStr,dbCTget,1
	if (Not rsCTget.Eof) then
		retVal = rsCTget("areacode")
		ioriginName = rsCTget("areaName")
	end if
	rsCTget.Close

	If (retVal = "") Then
		sqlStr = ""
		sqlStr = sqlStr & " SELECT top 1 areacode, areaName FROM db_outmall.dbo.[tbl_cjmall_SourceAreaCode]" & VBCRLF
		sqlStr = sqlStr & " WHERE CharIndex('"&iname&"',areaName) > 0" & VBCRLF
		sqlStr = sqlStr & " or CharIndex(areaName,'"&iname&"') > 0" & VBCRLF
		rsCTget.Open sqlStr,dbCTget,1
		If (Not rsCTget.Eof) then
			retVal = rsCTget("areacode")
			ioriginName = rsCTget("areaName")
		End If
		rsCTget.Close
	End If

	If (retVal = "") Then
		retVal="000"
		ioriginName = "없음"
	End If

	getOriginName2Code=retVal
	Exit Function
End Function

Function getmakerName2Code(iname, byref ioriginName)
	Dim sqlStr , retVal
	sqlStr = ""
	sqlStr = sqlStr & " SELECT top 1 code, makerName" & VBCRLF
	sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_cjmall_makerName" & VBCRLF
	sqlStr = sqlStr & " WHERE makerName='"&iname&"'"
	rsCTget.Open sqlStr,dbCTget,1
	if (Not rsCTget.Eof) then
		retVal = rsCTget("code")
		ioriginName = rsCTget("makerName")
	end if
	rsCTget.Close

	If (retVal = "") Then
		retVal="15210"
		ioriginName = "텐바이텐"
	End If

	getmakerName2Code = retVal
	Exit Function
End Function
'######################################################## 함수 모음 ###############################################################
'#######################################				1.신규 상품 등록 시작				#######################################
Function regCjMallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow
	Dim oCJMallItem, sellmoney
	Dim strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCJMallNotRegItemList
	    If (oCJMallItem.FResultCount > 0) Then
			If (oCJMallItem.FItemList(0).FCddKey = "") Then
				Response.Write "<script language=javascript>alert('상품분류 매칭을 하지 않은 상품번호: [" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If

			If (oCJMallItem.FItemList(0).Flimityn = "Y") and (oCJMallItem.FItemList(0).Flimitno - oCJMallItem.FItemList(0).Flimitsold < CMAXLIMITSELL) Then
				ierrStr = ierrStr & " - 한정수량 부족 ("&oCJMallItem.FItemList(0).Flimitno - oCJMallItem.FItemList(0).Flimitsold&") 개 남음"
				cause = "limitErr"
			End If

			sqlStr = ""
			sqlStr = sqlStr & " IF NOT Exists(SELECT * FROM db_outmall.dbo.tbl_cjmall_regItem where itemid="&iitemid&")"
			sqlStr = sqlStr & " BEGIN"& VbCRLF
			sqlStr = sqlStr & " INSERT INTO db_outmall.dbo.tbl_cjmall_regItem "
	        sqlStr = sqlStr & " (itemid, regdate, reguserid, cjmallstatCD, regitemname)"
	        sqlStr = sqlStr & " VALUES ("&iitemid&", getdate(), '"&session("SSBctID")&"', '1', '"&html2db(oCJMallItem.FItemList(0).FItemName)&"')"
			sqlStr = sqlStr & " END "
			dbCTget.Execute sqlStr
			'##상품옵션 검사(옵션수가 맞지 않거나 모두 전체 제외옵션일 경우 Pass)
			If oCJMallItem.FItemList(0).checkTenItemOptionValid Then
			    On Error Resume Next
				strParam = ""
				strParam = oCJMallItem.FItemList(0).getCjmallItemRegXML()
				If Err <> 0 Then
				    rw Err.Description
					Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
					dbCTget.Close: Response.End
				End If

				If CLng(10000 - oCJMallItem.FItemList(0).Fbuycash / oCJMallItem.FItemList(0).Fsellcash * 100 * 100) / 100 < 15 Then
					sellmoney = oCJMallItem.FItemList(0).Forgprice
				Else
					sellmoney = oCJMallItem.FItemList(0).Fsellcash
				End If

	            On Error Goto 0
	            iErrStr = ""
				ret1 = cjmallOneItemReg(iitemid, strParam, iErrStr, sellmoney, oCJMallItem.FItemList(0).getCjmallSellYn, oCJMallItem.FItemList(0).FLimityn, oCJMallItem.FItemList(0).FLimitNo, oCJMallItem.FItemList(0).FLimitSold, html2db(oCJMallItem.FItemList(0).FItemName), "createNewProduct")
	            If (ret1) Then
	            	regCjMallOneItem = true
	            Else
	                CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
	                retErrStr = retErrStr & iErrStr
	            End If
			Else
				CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
				iErrStr = "["&iitemid&"] : 옵션검사 실패 | 한정옵션의 개수가 5개 이하일 수 있음"
				retErrStr = retErrStr & iErrStr
			End If

		    If (retErrStr <> "") Then
		        Response.Write "<script language=javascript>alert('"&Replace(retErrStr,"'","")&"');</script>"
		    End If
		Else
	    	regCjMallOneItem = false
            CALL Fn_AcctFailTouch("cjmall",iitemid,"등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...")

	        If (IsAutoScript) Then
	            rw "S_ERR|등록가능상품 없음 :등록조건 확인: 판매Y, 할인..."
	            dbCTget.Close: Response.End
	        Else
	            Response.Write "<script language=javascript>alert('["&iitemid&"]:등록가능상품 없음 :등록조건 확인: 판매Y, 옵션추가액...');</script>"
				dbCTget.Close: Response.End
			End If
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallOneItemReg(iitemid, strParam, byRef iErrStr, iSellCash, ihomeplusSellYn, ilimityn, ilimitno, ilimitsold, iitemname, mode)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql
	Dim retCode, goodsCd, iMessage, AssignedRow
	If (xmlStr = "") Then
		cjmallOneItemReg = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
		goodsCd		= xmlDOM.getElementsByTagName("ns1:itemCd").item(0).text
		retCode		= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text

		If retCode = "true" Then		'성공(200)
			strSql = ""
			strSql = strSql & " UPDATE R"
			strSql = strSql & "	Set cjmallLastUpdate = getdate() "
			strSql = strSql & "	, cjmallPrdNo = '" & goodsCd & "'"
			strSql = strSql & "	, cjmallPrice = " &iSellCash
			strSql = strSql & "	, accFailCnt = 0"
			strSql = strSql & "	, cjmallRegdate = isNULL(cjmallRegdate, getdate())"
		    strSql = strSql & "	, cjmallStatCd=(CASE WHEN isNULL(cjmallStatCd, -1) < 3 then 3 ELSE cjmallStatCd END)"	'등록완료(임시)
			strSql = strSql & "	From db_outmall.dbo.tbl_cjmall_regItem R"
			strSql = strSql & " Where R.itemid = '" & iitemid & "'"
			dbCTget.Execute(strSql)
			cjmallOneItemReg = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:"&iMessage
		Else						'실패(E)
			cjmallOneItemReg = false
		    iErrStr =  "상품 등록중 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#######################################				1.신규 상품 등록 끝					#######################################
'#######################################				2.상품 조회 시작					#######################################
Function oneCjMallItemConfirm(iitemid, ierrStr)
	Dim AssignedRow
	Dim cjMallPrdNo : cjMallPrdNo = getCjMallPrdNoByItemID(iitemid)
	Dim firstItemoption
	Dim strParam, strRst
    On Error Resume Next
		If (cjMallPrdNo <> "") Then
			strRst = ""
			strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
			strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
			strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"
			strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"
			strRst = strRst &"<tns:contents>"
			strRst = strRst &"	<tns:sinstDtFrom>"&"2013-04-01"&"</tns:sinstDtFrom>"
			strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
			strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
			strRst = strRst &"	<tns:itemCd>"&cjMallPrdNo&"</tns:itemCd>"
			strRst = strRst &"</tns:contents>"
			strRst = strRst &"</tns:ifRequest>"
			strParam = strRst
		Else
			firstItemoption = getCjMallfirstItemoption(iitemid)
			strRst = ""
			strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
			strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_03_07"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_03_07.xsd"">"
			strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"
			strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"
			strRst = strRst &"<tns:contents>"
			strRst = strRst &"	<tns:sinstDtFrom>"&"2013-04-01"&"</tns:sinstDtFrom>"
			strRst = strRst &"	<tns:sinstDtTo>"&Left(now, 10)&"</tns:sinstDtTo>"
			strRst = strRst &"	<tns:schnCd>30</tns:schnCd>"
			strRst = strRst &"	<tns:vpn>"&iitemid&"_"&firstItemoption&"</tns:vpn>"
			strRst = strRst &"</tns:contents>"
			strRst = strRst &"</tns:ifRequest>"
			strParam = strRst
	    End If

		If Err <> 0 Then
		    rw Err.Description
			Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
			dbCTget.Close: Response.End
		End If
	On Error Goto 0

    iErrStr = ""
	ret1 = cjmallOneItemConfirm(iitemid, strParam, iErrStr)
    If (ret1) Then
    	oneCjMallItemConfirm = true
    Else
        CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
        retErrStr = retErrStr & iErrStr
    End If
End Function

Function cjmallOneItemConfirm(iitemid, strParam, byRef iErrStr)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, strSql
	Dim AssignedRow, Nodes, SubNodes
	Dim OverLapNo, SelOK, AssignedItemCnt
	Dim XitemCd, Xstatus, XslCls, XHapvpn, Xvpn, XunitCd, Xitemcode
	Dim uprItemNm, itemNm, slprc,exLeadtm, zClassId, packInd, purchvat, taxYn

	SelOK = 0
	AssignedItemCnt = 0

	If (xmlStr = "") Then
		cjmallOneItemConfirm = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		Set Nodes = xmlDOM.getElementsByTagName("ns1:unit")
		If (Not (xmlDOM is Nothing)) Then
			For each SubNodes in Nodes
				XitemCd = SubNodes.getElementsByTagName("ns1:itemCd")(0).Text		'판매코드
				Xstatus = SubNodes.getElementsByTagName("ns1:status")(0).Text		'결재상태
				XslCls 	= SubNodes.getElementsByTagName("ns1:slCls")(0).Text		'판매구분(상태)
				XHapvpn	= SubNodes.getElementsByTagName("ns1:vpn")(0).Text			'업체상품코드
				XunitCd = SubNodes.getElementsByTagName("ns1:unitCd")(0).Text		'단품코드
	
				uprItemNm= SubNodes.getElementsByTagName("ns1:uprItemNm")(0).Text	'판매상품명
				itemNm  = SubNodes.getElementsByTagName("ns1:itemNm")(0).Text		'단품상세
				slprc   = SubNodes.getElementsByTagName("ns1:slprc")(0).Text		'판매가
				exLeadtm= SubNodes.getElementsByTagName("ns1:exLeadtm")(0).Text		'리드타임(L/T)
				packInd = SubNodes.getElementsByTagName("ns1:packInd")(0).Text
				purchvat = SubNodes.getElementsByTagName("ns1:purchvat")(0).Text 	'매입가 vat포함?
				taxYn    = SubNodes.getElementsByTagName("ns1:taxYn")(0).Text

				Xvpn 		= Split(XHapvpn, "_")(0)
				Xitemcode 	= replace(Split(XHapvpn, "_")(1), "Q", "")

				'1.tbl_OutMall_regedoption 테이블에 있으면 업데이트 없으면 인서트 시키기
				strSql = ""
				strSql = strSql & " IF Exists(SELECT * FROM db_outmall.dbo.tbl_OutMall_regedoption WHERE itemid="&iitemid&" and mallid = '"&CMALLNAME&"' and itemoption = '"&Xitemcode&"' )"
				strSql = strSql & " BEGIN "
				strSql = strSql & " UPDATE db_outmall.dbo.tbl_OutMall_regedoption SET "
				strSql = strSql & " outmallsellyn='"&CHKIIF(XslCls="I","N","Y")&"'"
				If (Xitemcode <> "0000") Then
				    strSql = strSql & " , outmallOptName='"&html2DB(itemNm)&"'"
				End If
				strSql = strSql & " , outmallAddPrice="&slprc
				strSql = strSql & " , outmallleadTime='"&exLeadtm&"'"
				strSql = strSql & " , checkdate = getdate() "
				strSql = strSql & " , outmallsuppPrc="&purchvat*1.1
				strSql = strSql & " , outmallOptCode='"&XunitCd&"'"
				strSql = strSql & " WHERE itemid = '"&Xvpn&"' and itemoption = '"&Xitemcode&"' "
				strSql = strSql & " and mallid='"&CMALLNAME&"'"
				strSql = strSql & " END ELSE "
				strSql = strSql & " BEGIN "
				strSql = strSql & " INSERT INTO db_outmall.dbo.tbl_OutMall_regedoption "
				strSql = strSql & " (itemid, itemoption, mallid, outmallOptCode, outmallOptName, outmallSellyn, outmalllimityn, outmalllimitno, outmallAddPrice,outmallleadTime, outmallsuppPrc,lastUpdate, checkdate) "
				strSql = strSql & " VALUES "
				strSql = strSql & " ('"&Xvpn&"'"
				strSql = strSql & ",  '"&Xitemcode&"'"
				strSql = strSql & ", '"&CMALLNAME&"'"
				strSql = strSql & ", '"&XunitCd&"'"
				strSql = strSql & ", '"&html2db(CHKIIF(Xitemcode<>"0000", itemNm, "단일상품"))&"'"
				strSql = strSql & ", '"&CHKIIF(XslCls="I", "N", "Y")&"'"
				strSql = strSql & ", '"&"N"&"'"
				strSql = strSql & ", '0'"
				strSql = strSql & ", '"&slprc&"'"
				strSql = strSql & ", '"&exLeadtm&"'"
				strSql = strSql & ", "&purchvat*1.1&""
				strSql = strSql & ", getdate() "
				strSql = strSql & ", getdate()) "
				strSql = strSql & " END "
				dbCTget.Execute strSql, AssignedRow
				SelOK = SelOK + 1
				rw XHapvpn&"|"&XitemCd&"|"&XunitCd&"|"&Xstatus&"|"&XslCls&"|"&uprItemNm&"|"&itemNm&"|"&slprc&"|"&purchvat*1.1&"|"&exLeadtm&"|"&packInd
			Next

			'2.tbl_cjmall_regitem 테이블의 cjmallStatCd, lastStatCheckDate, cjmallsellyn, cjMallPrice, regedOptCnt 등 수정하기
			'2015-01-06 김진영 cjmallprdno도 수정 => cjmallprdno가 null인것 발견!
			strSql = ""
			strSql = strSql & " UPDATE R " & VBCRLF
			strSql = strSql & " SET cjmallregdate = isNULL(cjmallregdate,getdate())" & VBCRLF
			strSql = strSql & " ,cjmallStatCd = 7" & VBCRLF
			strSql = strSql & " ,lastStatCheckDate = getdate()" & VBCRLF                               ''수정
			strSql = strSql & " ,cjmallsellyn=(CASE WHEN T.SellCNT>0 THEN 'Y' ELSE 'N' END)"
            strSql = strSql & " ,cjMallPrice=(CASE WHEN T.mayItemPrice>0 then T.mayItemPrice ELSE R.cjMallPrice END)"
            strSql = strSql & " ,regedOptCnt=isNULL(T.regedOptCnt,0)"
            strSql = strSql & " ,cjmallprdno="&XitemCd
            strSql = strSql & " from db_outmall.dbo.tbl_cjmall_regItem R"
            strSql = strSql & " 	Join ("
            strSql = strSql & " 		select itemid, count(*) as optCNT"
            strSql = strSql & " 		, sum(CASE WHEN outmallsellyn='Y' THEN 1 ELSE 0 END) as SellCNT"
            strSql = strSql & " 		, min(outmallAddPrice) as mayItemPrice"
            strSql = strSql & " 		, sum(CASE WHEN itemoption<>'0000' THEN 1 ELSE 0 END) as regedOptCnt"
            strSql = strSql & " 		from db_outmall.dbo.tbl_OutMall_regedoption"
            strSql = strSql & " 		where itemid="&iitemid&""
            strSql = strSql & " 		and mallid='cjmall'"
            strSql = strSql & " 		group by itemid"
            strSql = strSql & " 	) T on R.itemid=T.itemid"
            strSql = strSql & " where R.itemid="&iitemid&""
			dbCTget.Execute strSql
			AssignedItemCnt = AssignedItemCnt + 1

			If SelOK = 0 Then
				If (iitemid <> "") Then
					''체크실패시 반복되지 않도록
					strSql = ""
					strSql = strSql & " update R"
					strSql = strSql & " set lastStatCheckDate = getdate()" & VBCRLF
					strSql = strSql & " from db_outmall.dbo.tbl_cjmall_regItem R" & VBCRLF
					strSql = strSql & " where itemid="&iitemid
					dbCTget.Execute strSql
				End If
				rw iitemid & "::"&"검색 결과 없음"
				cjmallOneItemConfirm = false
			Else
				rw "[" & iitemid & "]:상품 조회 성공 "&AssignedItemCnt&" 건 sync"
				cjmallOneItemConfirm = true
			End If
		End If
		on Error Goto 0
		Set Nodes = Nothing
	End If
End Function
'#######################################				2.상품 조회 끝						#######################################
'#######################################				3.판매 상태 수정 시작				#######################################
Function editSellStatusCjmallOneItem(byval iitemid, byRef ierrStr, cmd)
	If (cmd <> "Y") AND (cmd <> "N") Then
		rw "상품 상태가 Y or N이 아닙니다"
		Exit Function
	End If
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FREsultCount > 0) Then
 			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallItemSellStatusDTXML(cmd)

			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If

	        On Error Goto 0
	        iErrStr = ""
			ret1 = cjmallOneItemSellStatEdit(iitemid, oCJMallItem.FItemList(0).FCjmallPrdno, cmd, iErrStr, strParam, "setProductStatus")
			
	        If (ret1) Then
	        	editSellStatusCjmallOneItem = true
	        Else
	            CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
	            retErrStr = retErrStr & iErrStr
	        End If
		Else
			editSellStatusCjmallOneItem = false
		End If
	SET oCJMallItem = Nothing
End Function

Function cjmallOneItemSellStatEdit(iitemid, icjmallPrdNo, ichgSellYn, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr

	If (xmlStr = "") Then
		cjmallOneItemSellStatEdit = false
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
		retCode		= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text

		If retCode = "true" Then		'성공(200)
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE R"
			sqlStr = sqlStr & " SET cjmallSellYn = '"&ichgSellYn&"'"
			sqlStr = sqlStr & " ,cjmallLastUpdate = getdate()"
			sqlStr = sqlStr & " ,accFailCNT=0"
			sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_cjmall_regItem as R"
			sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item as i on R.itemid = i.itemid"
			sqlStr = sqlStr & " WHERE R.itemid = "&iitemid
			dbCTget.Execute sqlStr,AssignedRow
			cjmallOneItemSellStatEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]: 판매상태("&ichgSellYn&")으로 변경"
		Else						'실패(E)
			cjmallOneItemSellStatEdit = false
		    iErrStr =  "상품 상태 수정 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#######################################				3.판매 상태 수정 끝					#######################################
'#######################################				4.상품 정보 수정 시작				#######################################
Function editCjmallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FResultCount > 0) Then
			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallItemModXML()
			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
	        On Error Goto 0
            iErrStr = ""
			ret1 = cjmallOneItemEdit(iitemid, iErrStr, strParam, "updateProduct")
            If (ret1) Then
            	editCjmallOneItem = true
            Else
                CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
                retErrStr = retErrStr & iErrStr
                editCjmallOneItem = false
            End If
		Else
			editCjmallOneItem = false
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallOneItemEdit(iitemid, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr
	If (xmlStr = "") Then
		cjmallOneItemEdit = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		retCode		= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
		iMessage	= xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text

		If retCode = "true" Then		'성공(200)
			sqlStr = ""
			sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_cjmall_regItem "
			sqlStr = sqlStr & " SET regitemname = B.itemname "
			sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_cjmall_regItem A "
			sqlStr = sqlStr & " JOIN db_AppWish.dbo.tbl_item B on A.itemid = B.itemid "
			sqlStr = sqlStr & " WHERE A.itemid='" & iitemid & "'"
			dbCTget.Execute(sqlStr)
			cjmallOneItemEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
			rw "[" & iitemid & "]:정보 수정 성공"
		Else							'실패(E)
			'lastStatCheckDate수정하는 이유 : 옵션이 새로 추가가 되면 regedoption에 저장 하는 것이 아님, 상품 조회에서만 CJ측 단품코드를 얻을 수 있음.
			'그래서 스케줄링을 lastStatCheckDate ASC로 하기 때문에 아래 작업이 필요함.
			If (Trim(iMessage)="1번째 단품:유효하지 않은 값입니다.[단품정보-협력사상품코드(Vpn)]가 이미 존재합니다. 새로운 값을 사용하세요.") then
				sqlStr = ""
				sqlStr = sqlStr & " UPDATE R"
				sqlStr = sqlStr & " SET lastStatCheckDate=NULL"                   '''등록실패
				sqlStr = sqlStr & " FROM db_outmall.dbo.tbl_cjmall_regitem as R"
				sqlStr = sqlStr & " WHERE R.itemid = "&iitemid
				dbCTget.Execute(sqlStr)
			End If

			cjmallOneItemEdit = false
		    iErrStr =  "상품 정보 수정 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#######################################				4.상품 정보 수정 끝					#######################################
'#######################################				5.단품 상태 수정 시작				#######################################
Function editDTCjmallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FResultCount > 0) Then
			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallOptSellModXML()
			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
	        On Error Goto 0
            iErrStr = ""
			ret1 = cjmallOptionSellStatEdit(iitemid, iErrStr, strParam, oCJMallItem.FItemList(0).FMaySoldout, "updateOptSellyn")
            If (ret1) Then
            	editDTCjmallOneItem = true
            Else
                CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
                retErrStr = retErrStr & iErrStr
                editDTCjmallOneItem = false
            End If
		Else
			editDTCjmallOneItem = false
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallOptionSellStatEdit(iitemid, byRef iErrStr, strParam, imaySoldout, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr
    Dim Nodes, SubNodes
    Dim itemCd_zip, packInd, typeCd
    Dim sellynCnt, maySellYn
    Dim AssignedItemCnt : AssignedItemCnt = 0
	If (xmlStr = "") Then
		cjmallOptionSellStatEdit = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If

		Set Nodes = xmlDOM.getElementsByTagName("ns1:itemStates")
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					retCode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					If retCode = "true" Then		'성공(200)
					    iMessage    = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
					    typeCd      = SubNodes.getElementsByTagName("ns1:typeCd").item(0).text
						itemCd_zip 	= SubNodes.getElementsByTagName("ns1:itemCd_zip").item(0).text
						packInd		= SubNodes.getElementsByTagName("ns1:packInd").item(0).text
						
						rw "["&iitemid&"]:" &CHKIIF(typeCd="02","옵션","상품")&","&itemCd_zip&","&CHKIIF(packInd="A","판매","품절")&","&CHKIIF(retCode<>"true",iMessage,"판매상태 수정완료")
						If typeCd = "02" Then
							sqlStr = ""
							sqlStr = sqlStr & " UPDATE [db_outmall].[dbo].tbl_OutMall_regedoption  " & VBCRLF
							sqlStr = sqlStr & " SET outmallSellyn = '"&chkiif(packInd="A","Y","N")&"'" & VBCRLF
							sqlStr = sqlStr & " , lastupdate = getdate() " & VBCRLF
							sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'  " & VBCRLF
							sqlStr = sqlStr & " and outmallOptCode = '"&itemCd_zip&"' " & VBCRLF
							sqlStr = sqlStr & " and mallid='"&CMALLNAME&"'"&VbCRLF
							dbCTget.Execute sqlStr
							AssignedItemCnt = AssignedItemCnt + 1
						ElseIf typeCd = "01" Then
							AssignedItemCnt = AssignedItemCnt + 1
						End If
					Else
						iMessage    = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
						CALL Fn_AcctFailTouch("cjmall", iitemid, iMessage)
						rw "["&iitemid&"]:" &CHKIIF(typeCd="02","옵션","상품")&","&itemCd_zip&","&CHKIIF(packInd="A","판매","품절")&","&CHKIIF(retCode<>"true",iMessage,"판매상태 수정완료")
					End If
				End If
			Next

			If AssignedItemCnt > 0 Then
				sqlStr = ""
				sqlStr = sqlStr & " SELECT count(*) as cnt FROM db_outmall.dbo.tbl_Outmall_regedoption WHERE itemid="&iitemid&" and mallid = 'cjmall' and outmallSellyn = 'Y' "
				rsCTget.Open sqlStr, dbCTget
					sellynCnt = rsCTget("cnt")
				rsCTget.Close
		
				If (imaySoldout = "Y") OR (sellynCnt = 0) Then
					maySellYn = "N"
				Else
					maySellYn = "Y"
				End If

		        sqlStr = ""
		        sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_cjmall_regItem SET cjmallLastUpdate = getdate() "
				sqlStr = sqlStr & " ,cjmallSellyn = '"&maySellYn&"'"
		        sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'"
		        dbCTget.Execute sqlStr
		        cjmallOptionSellStatEdit = true
				Set objXML = Nothing
				Set xmlDOM = Nothing
		    Else
		    	cjmallOptionSellStatEdit = false
			    iErrStr =  "단품 상태 수정 오류 [" & iitemid & "]"
				Set objXML = Nothing
				Set xmlDOM = Nothing
			    Exit Function
		    End If
		Else
			cjmallOptionSellStatEdit = false
		End If
		On Error Goto 0
	End If
End Function
'#######################################				5.단품 상태 수정 끝					#######################################
'#######################################				5.단품 수량 수정 시작				#######################################
Function editqtyCjmallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FResultCount > 0) Then
			If oCJMallItem.FItemList(0).FMaySoldout = "Y" Then
				rw "["&iitemid&"]:품절에 해당하는 상품으로 수량 수정을 하지 않습니다."
				SET oCJMallItem = nothing
				Exit Function
			End If
			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallItemQTYXML()
			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
	        On Error Goto 0
			iErrStr = ""
			ret1 = cjmallOptionSuEdit(iitemid, iErrStr, strParam, "updateOptSu")
			If (ret1) Then
				editqtyCjmallOneItem = true
			Else
				CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				editqtyCjmallOneItem = false
			End If
		Else
			editqtyCjmallOneItem = false
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallOptionSuEdit(iitemid, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr
    Dim Nodes, SubNodes
    Dim unitCd, strDt, endDt, availSupQty
    Dim AssignedItemCnt : AssignedItemCnt = 0
	If (xmlStr = "") Then
		cjmallOptionSuEdit = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
'		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
On Error Resume Next
		Set Nodes = xmlDOM.getElementsByTagName("ns1:ltSupplyPlans")
On Error Goto 0
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					retCode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					If retCode = "true" Then		'성공(200)
                        iMessage        = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
                        unitCd          = SubNodes.getElementsByTagName("ns1:unitCd").item(0).text
                        strDt           = SubNodes.getElementsByTagName("ns1:strDt").item(0).text
                        endDt           = SubNodes.getElementsByTagName("ns1:endDt").item(0).text
                        availSupQty     = SubNodes.getElementsByTagName("ns1:availSupQty").item(0).text

                        If (strDt = endDt) Then
                            availSupQty=0
                        End If
                        rw "["&iitemid&"]:" &unitCd&", 시작:"&strDt&", 종료:"&endDt& ", 수량:"&availSupQty&", "&CHKIIF(retCode<>"true",iMessage,"판매수량 수정완료")

                        sqlStr = "UpDate R"&VbCRLF
    				    sqlStr = sqlStr & " SET outmalllimitno="&availSupQty&VbCRLF
    				    If availSupQty < 0 Then
    				    	sqlStr = sqlStr & " ,outmalllimityn='N'"
    				    Else
    						sqlStr = sqlStr & " ,outmalllimityn='Y'"
    					End If
    				    sqlStr = sqlStr & " from db_outmall.dbo.tbl_OutMall_regedoption R"&VbCRLF
    				    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"&VbCRLF
    				    sqlStr = sqlStr & "  and itemid="&iitemid&VbCRLF
    				    sqlStr = sqlStr & "  and outmallOptCode='"&unitCd&"'"&VbCRLF
    				    dbCTget.Execute sqlStr
    				    AssignedItemCnt = AssignedItemCnt + 1
					Else
						iMessage    = SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
						rw "["&iitemid&"]:" &unitCd&", 시작:"&strDt&", 종료:"&endDt& ", 수량:"&availSupQty&", "&CHKIIF(retCode<>"true",iMessage,"판매수량 수정완료")
					End If
				End If
			Next
			If AssignedItemCnt > 0 Then
		        cjmallOptionSuEdit = true
				Set objXML = Nothing
				Set xmlDOM = Nothing
			Else
		    	cjmallOptionSuEdit = false
			    iErrStr =  "단품 수량 수정 오류 [" & iitemid & "]"
				Set objXML = Nothing
				Set xmlDOM = Nothing
			    Exit Function
			End If
		Else
			cjmallOptionSuEdit = false
			Set objXML = Nothing
			Set xmlDOM = Nothing
		    Exit Function
		End If
'		On Error Goto 0
	End If
End Function
'#######################################				5.단품 수량 수정 끝					#######################################
'#######################################				5.상품 가격 수정 시작				#######################################
Function editSellPriceCjmallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FResultCount > 0) Then
			If oCJMallItem.FItemList(0).FMaySoldout = "Y" Then
				rw "["&iitemid&"]:품절에 해당하는 상품으로 가격 수정을 하지 않습니다."
				SET oCJMallItem = nothing
				Exit Function
			End If
			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallItemSellPriceModXML()
			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
	        On Error Goto 0
			iErrStr = ""
			ret1 = cjmallSellPriceEdit(iitemid, iErrStr, strParam, "updateSellPrice")
			If (ret1) Then
				editSellPriceCjmallOneItem = true
			Else
				CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				editSellPriceCjmallOneItem = false
			End If
		Else
			editSellPriceCjmallOneItem = false
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallSellPriceEdit(iitemid, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr
    Dim Nodes, SubNodes
    Dim typeCD, itemCD_ZIP, newUnitRetail, newUnitCost
	If (xmlStr = "") Then
		cjmallSellPriceEdit = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
		retCode			= xmlDOM.getElementsByTagName("ns1:successYn").item(0).text
		iMessage        = xmlDOM.getElementsByTagName("ns1:errorMsg").item(0).text
		typeCD 	        = xmlDOM.getElementsByTagName("ns1:typeCD").item(0).text
		itemCD_ZIP		= xmlDOM.getElementsByTagName("ns1:itemCD_ZIP").item(0).text
		newUnitRetail	= xmlDOM.getElementsByTagName("ns1:newUnitRetail").item(0).text
		newUnitCost	    = xmlDOM.getElementsByTagName("ns1:newUnitCost").item(0).text
		If retCode = "true" Then		'성공(200)
	        sqlStr = ""
	        sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_cjmall_regItem SET cjmallLastUpdate = getdate() "
			sqlStr = sqlStr & " ,cjmallprice = '"&newUnitRetail&"'"
			sqlStr = sqlStr & " ,accFailCnt = 0"
	        sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'"
	        dbCTget.Execute sqlStr
		    rw "["&iitemid&"]:CJ상품코드:"&itemCD_ZIP&", 판매가:"&newUnitRetail&",공급가:"&newUnitCost&","&CHKIIF(retCode<>"true",iMessage,"상품 가격 수정완료")
	        cjmallSellPriceEdit = true
			Set objXML = Nothing
			Set xmlDOM = Nothing
		Else
			cjmallSellPriceEdit = false
			iErrStr =  "상품 가격 수정 오류 [" & iitemid & "]:"&iMessage
			Set objXML = Nothing
			Set xmlDOM = Nothing
			Exit Function
		End If
		On Error Goto 0
	End If
End Function
'#######################################				5.상품 가격 수정 끝					#######################################
'#######################################				6.단품 가격 수정 시작				#######################################
Function editPriceCjmallOneItem(byval iitemid, byRef ierrStr)
	Dim sqlStr, AssignedRow, ret1
	Dim oCJMallItem, strParam
	SET oCJMallItem = new CCjmall
		oCJMallItem.FRectItemID = iitemid
		oCJMallItem.getCjmallEditedItemList
		If (oCJMallItem.FResultCount > 0) Then
			If oCJMallItem.FItemList(0).FMaySoldout = "Y" Then
				rw "["&iitemid&"]:품절에 해당하는 상품으로 가격 수정을 하지 않습니다."
				SET oCJMallItem = nothing
				Exit Function
			End If

			If (oCJMallItem.FItemList(0).FaccFailCnt > 0) and oCJMallItem.FItemList(0).FcjmallSellYn = "N" Then
				If Instr(oCJMallItem.FItemList(0).FlastErrStr, "SCM일시중단된 단품") > 0 Then
					rw "["&iitemid&"]:CJ몰 SCM일시중단된 단품메세지 출력으로 가격 수정하지 않습니다."
					SET oCJMallItem = nothing
					Exit Function
				End If
			End If

			On Error Resume Next
			strParam = ""
			strParam = oCJMallItem.FItemList(0).getcjmallOptionPriceModXML()
			If Err <> 0 Then
			    rw Err.Description
				Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
				dbCTget.Close: Response.End
			End If
	        On Error Goto 0
			iErrStr = ""
			ret1 = cjmallOptionSellPriceEdit(iitemid, iErrStr, strParam, "updateOptionSellPrice")
			If (ret1) Then
				editPriceCjmallOneItem = true
			Else
				CALL Fn_AcctFailTouch("cjmall", iitemid, iErrStr)
				retErrStr = retErrStr & iErrStr
				editPriceCjmallOneItem = false
			End If
		Else
			iErrStr =  "단품 가격 수정 오류 [" & iitemid & "]:수정 리스트에 포함되지 않음"
			editPriceCjmallOneItem = false
		End If
	SET oCJMallItem = nothing
End Function

Function cjmallOptionSellPriceEdit(iitemid, byRef iErrStr, strParam, mode)
    Dim xmlStr : xmlStr = strParam
    Dim objXML, xmlDOM, retCode, iMessage
    Dim strRst, sqlStr
    Dim Nodes, SubNodes
    Dim typeCD, itemCD_ZIP, newUnitRetail, newUnitCost
    Dim AssignedItemCnt : AssignedItemCnt = 0
	If (xmlStr = "") Then
		cjmallOptionSellPriceEdit = false
		Exit Function
    End If
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)
	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
'			response.write objXML.ResponseText
'			response.end
		If (session("ssBctID")="icommang") or (session("ssBctID")="kjy8517") Then
'			response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		End If
		Set Nodes = xmlDOM.getElementsByTagName("ns1:itemPrices")
		If (Not (Nodes is Nothing)) Then
			For each SubNodes in Nodes
				If (Not (SubNodes is Nothing)) Then
					retCode		= SubNodes.getElementsByTagName("ns1:successYn").item(0).text
					If retCode = "true" Then		'성공(200)
						iMessage		= SubNodes.getElementsByTagName("ns1:errorMsg").item(0).text
						typeCD			= SubNodes.getElementsByTagName("ns1:typeCD").item(0).text
						itemCD_ZIP		= SubNodes.getElementsByTagName("ns1:itemCD_ZIP").item(0).text
						newUnitRetail	= SubNodes.getElementsByTagName("ns1:newUnitRetail").item(0).text
						newUnitCost		= SubNodes.getElementsByTagName("ns1:newUnitCost").item(0).text
						If (typeCD = "01") Then
						    sqlStr = ""
						    sqlStr = sqlStr & " UPDATE db_outmall.dbo.tbl_cjmall_regItem SET cjmallLastUpdate = getdate() "
							sqlStr = sqlStr & " ,cjmallprice = '"&newUnitRetail&"'"
							sqlStr = sqlStr & " ,accFailCnt = 0"
							sqlStr = sqlStr & " ,lastpriceCheckDate = getdate()"
						    sqlStr = sqlStr & " WHERE itemid = '"&iitemid&"'"
						    dbCTget.Execute sqlStr
						    rw "["&iitemid&"]:CJ상품코드:"&itemCD_ZIP&", 판매가:"&newUnitRetail&",공급가:"&newUnitCost&","&CHKIIF(retCode<>"true",iMessage,"상품 가격 수정완료")
						ElseIf (typeCD = "02") Then
						    sqlStr = "UpDate R " 
						    sqlStr = sqlStr & " SET outmallAddPrice="&newUnitRetail
						    sqlStr = sqlStr & " ,lastupdate=getdate()"
						    sqlStr = sqlStr & " ,checkdate=getdate()"
						    sqlStr = sqlStr & " from db_outmall.dbo.tbl_OutMall_regedoption R"
						    sqlStr = sqlStr & "  where mallid='"&CMALLNAME&"'"
						    sqlStr = sqlStr & "  and itemid="&iitemid
						    sqlStr = sqlStr & "  and outmallOptCode='"&itemCD_ZIP&"'"
						    dbCTget.Execute sqlStr
						    rw "["&iitemid&"]:CJ단품코드:"&itemCD_ZIP&", 판매가:"&newUnitRetail&",공급가:"&newUnitCost&","&CHKIIF(retCode<>"true",iMessage,"단품가격 수정완료")
						End If
						AssignedItemCnt = AssignedItemCnt + 1
					End If
				End If
			Next
			If AssignedItemCnt > 0 Then
		        cjmallOptionSellPriceEdit = true
				Set objXML = Nothing
				Set xmlDOM = Nothing
			Else
		    	cjmallOptionSellPriceEdit = false
			    iErrStr =  "단품 가격 수정 오류 [" & iitemid & "]"
				Set objXML = Nothing
				Set xmlDOM = Nothing
			    Exit Function
			End If
		Else
			cjmallOptionSellPriceEdit = false
		End If
		On Error Goto 0
	End If
End Function
'#######################################				6.단품 가격 수정 끝					#######################################
'#######################################				6.공통코드 조회 시작				#######################################
Function getcjCommonCodeList(ccd)
	Dim AssignedRow
	Dim strParam, strRst
    On Error Resume Next
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_02_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_02_01.xsd"">"
        strRst = strRst &"<tns:vendorId>"&CVENDORID&"</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>"&CVENDORCERTKEY&"</tns:vendorCertKey>"
        strRst = strRst &"<tns:lgrpCd>"&ccd&"</tns:lgrpCd>"
        strRst = strRst &"</tns:ifRequest>"
        strParam = strRst
		If Err <> 0 Then
		    rw Err.Description
			Response.Write "<script language=javascript>alert('텐바이텐 상품정보 생성중 오류가 발생했습니다.\n관리자에게 전달 부탁드립니다.[상품번호:" & iitemid & "]');</script>"
			dbCTget.Close: Response.End
		End If
	On Error Goto 0

    iErrStr = ""
	Call cjmallCommCd(strParam, iErrStr)
End Function

Function cjmallCommCd(strParam, byRef iErrStr)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM
	If (xmlStr = "") Then
		Exit Function
    End If

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "POST", cjMallAPIURL, false
		objXML.setRequestHeader "CONTENT_TYPE", "text/xml"
		objXML.send(xmlStr)

	If objXML.Status = "200" Then
		Set xmlDOM = server.createobject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		On Error Resume Next
			response.write objXML.ResponseText
			Set objXML = Nothing
			Set xmlDOM = Nothing
		On Error Goto 0
	End If
End Function
'#######################################				6.공통코드 조회 끝					#######################################
'######################################################## 함수 모음 끝 ############################################################
'CJ주문내역 조회
Function getCjOrderList(imode,sday) ''"ORDLIST" , "ORDCANCELLIST"
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode, "")

	If (xmlStr = "") Then
		getCjOrderList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
    'rw retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjOrderList = saveORDListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'CJ CS내역 조회
Function getCjCsList(imode,sday)
    Dim mode : mode = imode
	Dim xmlStr : xmlStr = getXMLString(sday, mode, "")

	If (xmlStr = "") Then
		getCjCsList = False
		Exit Function
    End If

    Dim retDoc, sURL
    sURL = cjMallAPIURL

    SET retDoc = xmlSend(sURL, xmlStr)
	''response.write retDoc.XML
	    If (isCJ_DebugMode) Then
	        CALL XMLFileSave(retDoc.XML, "RET_"&mode, sday)
	    End If
		getCjCsList = saveCSListResult(retDoc, mode, sday)
    SET retDoc = Nothing
End Function

'주문내역 저장용
Function saveORDListResult(retDoc, mode, sday)
    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false
    Dim ordNo,custNm,custTelNo,custDeliveryCost

    Dim ordGSeq, ordDSeq, ordWSeq, ordDtlCls, ordDtlClsCd, wbCrtDt, outwConfDt, delivDtm, cnclInsDtm
    Dim oldordNo, toutYn, chnNm, receverNm, recvName, zipno, addr_1, addr_2, addr, telno, cellno
    Dim msgSpec, delvplnDt, packYn, itemNm, itemCd, unitCd, itemName, unitNm, contItemCd, wbIdNo
    Dim outwQty, realslAmt, outwAmt, delivInfo, promGiftSpec, cnclRsn, cnclRsnSpec, ordDtm, juminNum, dccouponCjhs, dccouponVendor
    Dim dtlCnt

	dim IsDetailItemAllCancel, IsCancelOrgOrder
	dim strSql

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

    Set Nodes = retDoc.getElementsByTagName("ns1:errorMsg")
    If (Not (Nodes is Nothing)) Then
        For each ErrNode in Nodes
            errorMsg = Nodes.item(0).text
            isErrExists = true
            rw "["&sday&"]"&errorMsg
        next
    end if

    if (Not isErrExists) then
        Set Nodes = retDoc.getElementsByTagName("ns1:instruction")

        If (Not (Nodes is Nothing)) Then
            For each masterSubNodes in Nodes
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '주문번호
                custNm      = masterSubNodes.getElementsByTagName("ns1:custNm")(0).Text	        '주문자
                custTelNo   = masterSubNodes.getElementsByTagName("ns1:custTelNo")(0).Text	    '주문자 전화
                custDeliveryCost = masterSubNodes.getElementsByTagName("ns1:custDeliveryCost")(0).Text	'배송비

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                ''rw ordNo&"|"&custNm&"|"&custTelNo&"|"&custDeliveryCost

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        requireDetail = ""
                        ierrStr = ""

                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:주문상품순번], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:주문상세순번], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:주문처리순번], 001
                        ordDtlCls = detailSubNodeItem.getElementsByTagName("ns1:ordDtlCls")(0).Text	        ' 주문정보 - 주문구분, 주문
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' 주문정보 - 주문구분코드, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' 주문정보 - 지시일자, 2013-05-22+09:00
                        ''outwConfDt	'주문정보 - 출고확정일자
                        ''delivDtm	    '주문정보 - 배송완료일
                        ''cnclInsDtm	'주문정보 - 취소일자
                        ''oldordNo	    '주문정보 - 원주문번호
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '주문정보 - 기출하구분(Y-기출하,N-정상출하), N
                        chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '주문정보 - 채널구분, INTERNET

                        if (mode<>"ORDLISTUP") then
                        receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '주문정보 - 인수자, 채현아
                        end if

                        'recvName	    '주문정보 - 수취인 영문명
                        zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text	                '주문정보 - 우편번호, 110809
                        addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text	            '주문정보 - 주소, 서울 종로구 동숭동
                        addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text	            '주문정보 - 상세주소, 1-45번지 자유빌딩 6층
                        'addr	        '주문정보 - 주소
                        telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text	                '주문정보 - 인수자tel, 02)973-8514
                        cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text	            '주문정보 - 인수자HP, 010)2715-8514
                        'msgSpec	    '주문정보 - 배송참고
                        'delvplnDt	    '주문정보 - 배송예정일
                        packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text	            '상품정보 - 세트여부, 일반
                        'itemNm	        '상품정보 - 세트상품명
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '상품정보 - 판매코드, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '상품정보 - 단품코드, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '상품정보 - 판매상품명, 24K Gold 전자파차단스티커
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '상품정보 - 단품상세, ES-01 잘될꺼야
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '상품정보 - 협력사상품코드, 279751_0011
                        wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text	            '상품정보 - 운송장식별번호, 20000420537940
                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '상품정보 - 수량, 1.0
                        realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text	        '상품정보 - 판매가, 1800.0
                        outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text	            '상품정보 - 고객결제가, 1800.0  :: 수량*판매가 인지, 수량*실판매가인지 확인 = 수량*실판매가
                        'delivInfo	    '기타정보 - 비고
                        'promGiftSpec	'기타정보 - 사은품내역
                        'juminNum       '주문정보-주민번호(아님), 발송 금지!
                        'cnclRsn	    '기타정보 - 교환/취소사유
                        'cnclRsnSpec	'기타정보 - 교환/취소사유상세
                        ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text	            '주문정보-주문일시, 2013-05-22 15:05:02


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
                        recvName =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text        '주문정보 - 수취인 영문명
                        end if
                        addr =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text        '주문정보 - 주소all?
                        end if
                        msgSpec =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text        '주문정보 -배송참고
                        end if
                        delvplnDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text        '주문정보 -배송예정일
                        end if
                        itemNm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text        '상품정보 -세트상품명
                        end if
                        juminNum =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text       '주문정보-주민번호(아님), 발송 금지!
                        end if
                        dccouponCjhs =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text       '주문정보 - 할인(당사부담)금액
                        end if
                        dccouponVendor =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text      '주문정보 - 할인(협력사부담)금액
                        end if

                        orderDlvPay = custDeliveryCost
                        if (dtlCnt>1) then
                            orderDlvPay = 0 ''첫번째 값만 넣음.
                        end if

                        orderCsGbn = ""
						if (toutYn <> "N") then
							'// 기출하 주문 스킵
							ordDtlClsCd = "99"
						end if
                        if (ordDtlClsCd="10") then
                            orderCsGbn="0"
                        elseif (ordDtlClsCd="20") then
                            orderCsGbn="2"  ''취소
                        end if

                        ''requireDetail = juminNum '' 주문제작문구
                        if (juminNum<>"") then                          ''2013/06/05 수정: 주문제작문구 아님?.
                            if (msgSpec<>"") then
                                msgSpec=msgSpec&VbCRLF&juminNum
                            else
                                msgSpec=juminNum
                            end if
                        end if

                        ierrCode = 0
                        ierrStr  = ""

                        if (mode="ORDLIST") then
                            if (orderCsGbn<>"") then

    							IsDetailItemAllCancel = False
    							IsCancelOrgOrder = False

    							if (orderCsGbn = "2") then
    								'// 취소
    								strSql = " select matchState, orderDlvPay, orgOrderCNT from db_temp.dbo.tbl_xSite_TMPOrder "
    								strSql = strSql + " where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' "
    								''rw strSql
    								rsget.Open strSql,dbget,1
    								if (Not rsget.Eof) then
    									if (CLng(outwQty) = rsget("orgOrderCNT")) then
    										'// 특정상품 전체취소
    										IsDetailItemAllCancel = True
    										if (rsget("matchState") = "I") then
    											'// 주문입력이전
    											IsCancelOrgOrder = True
    										end if
    									end if
    								end if
    								rsget.Close

    								if (IsDetailItemAllCancel and IsCancelOrgOrder) then
    									strSql = " update db_temp.dbo.tbl_xSite_TMPOrder set matchState = 'D' "
    									strSql = strSql + " where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' and matchState = 'I' "
    									''rw strSql
    									rsget.Open strSql, dbget, 1
    								end if
    							end if

                                '''899506_Q0055 ::
                                if (LEFT(splitvalue(contItemCd,"_",1),1)="Q") then
                                    contItemCd = replace(contItemCd,"Q","")
                                end if
                                if (outwQty<>"0") and (outwQty<>"1") and (outwQty<>"-1") and (outwQty<>"") then
                                    outwAmt = CLNG(outwAmt/outwQty)
                                end if
    							if (IsDetailItemAllCancel) then
    								'// 우선 수량 전체취소만 처리(수량 일부취소는 내역 입력되면 처리)
    								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
    										, custNm , custTelNo, custTelNo _
    										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
    										, realslAmt, outwAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA" _
    										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
    										, ierrCode, ierrStr)

    								'// 원주문 삭제되었으면 CS도 삭제
    								strSql = " if exists (select OutMallOrderSeq from db_temp.dbo.tbl_xSite_TMPOrder where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "' and matchState = 'D') "
    								strSql = strSql + " begin "
    								strSql = strSql + " 	update db_temp.dbo.tbl_xSite_TMPOrder set matchState = 'D' where SellSite = 'cjmall' and OutMallOrderSerial = '" + CStr(ordNo) + "' and OrgDetailKey = '" & ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq & "-CA' and matchState = 'I' "
    								strSql = strSql + " end "
    								rsget.Open strSql, dbget, 1
    							else
    								retVal = saveORDOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
    										, custNm , custTelNo, custTelNo _
    										, receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
    										, realslAmt, outwAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq _
    										, msgSpec, requireDetail, orderDlvPay, orderCsGbn _
    										, ierrCode, ierrStr)
    							end if
                            else
                                retVal = false
                                ierrStr = "주문구분 [ordDtlClsCd="&ordDtlClsCd&"] 정의되지 않음"
                            end if
                        else
                            rw ordNo&"|"&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"|"&realslAmt&"|"&outwAmt&"|"&outwAmt/outwQty

                            if (orderCsGbn<>"") then
                                sqlStr = " Update T"
                                sqlStr = sqlStr & " SET realSellPrice=(CASE WHEN SellPrice<>realSellPrice THEN realSellPrice ELSE "&outwAmt/outwQty&" END )"
                                sqlStr = sqlStr & " ,PRE_USE_UNITCOST=0"
                                sqlStr = sqlStr & " ,tenCpnUint=0"
                                sqlStr = sqlStr & " ,mallCpnUnit="&realslAmt-outwAmt/outwQty&""
                                sqlStr = sqlStr & " From db_temp.dbo.tbl_xSite_tmporder T"
                				sqlStr = sqlStr & " where sellsite='cjmall'"
                                sqlStr = sqlStr & " and outmallorderserial='"&ordNo&"'"
                                sqlStr = sqlStr & " and OrgDetailKey='"&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"'"
                                sqlStr = sqlStr & " and mallCpnUnit is NULL" ''2014/02/01
''rw sqlStr
                				dbget.Execute sqlStr
            				end if
                        end if

                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
                            failCnt = failCnt+1
                            if (ierrCode=-1) then skipCnt = skipCnt+1

                            if (mode="ORDCANCELLIST") then
                                rw "<font color='red'>["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"-CA]</font> "&ierrStr &" / "&custNm
                            else
                                rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr &" / "&custNm
                            end if
                        end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "건 입력"
    rw failCnt & "건 실패" & "("&skipCnt&" 건 skip)"

End function

'주문내역 저장용
Function saveCSListResult(retDoc, mode, sday)

	'// TODO : !!!!
    Exit function

    Dim Nodes, masterSubNodes, detailSubNodes, detailSubNodeItem, ErrNode, errorMsg
    Dim isErrExists : isErrExists = false
    Dim ordNo

    Dim ordGSeq, ordDSeq, ordWSeq, wbClsCd, wbCls
    Dim wbCrtDt, outwConfDt, confirmChk, cnclInsDtm
    Dim oldordNo, toutYn, chnNm, receverNm, recvName, zipno, addr_1, addr_2, addr, telno, cellno
    Dim msgSpec, delvplnDt, packYn, itemNm, itemCd, unitCd, itemName, unitNm, contItemCd, wbIdNo
    Dim outwQty, realslAmt, outwAmt, delivInfo, promGiftSpec, cnclRsn, cnclRsnSpec, ordDtm, juminNum, dccouponCjhs, dccouponVendor
    Dim dtlCnt

    Dim requireDetail, orderDlvPay, orderCsGbn, ierrStr, ierrCode
    dim succCnt : succCnt=0
    dim failCnt : failCnt=0
    dim skipCnt : skipCnt=0
    dim retVal

    Set Nodes = retDoc.getElementsByTagName("ns1:errorMsg")
    If (Not (Nodes is Nothing)) Then
        For each ErrNode in Nodes
            errorMsg = Nodes.item(0).text
            isErrExists = true
            rw "["&sday&"]"&errorMsg
        next
    end if

    if (Not isErrExists) then
        Set Nodes = retDoc.getElementsByTagName("ns1:instruction")

        If (Not (Nodes is Nothing)) Then
            For each masterSubNodes in Nodes
                ordNo       = masterSubNodes.getElementsByTagName("ns1:ordNo")(0).Text	        '주문번호

                Set detailSubNodes = masterSubNodes.getElementsByTagName("ns1:instructionDetail")

                ''rw ordNo&"|"&custNm&"|"&custTelNo&"|"&custDeliveryCost

                dtlCnt = 1
                If (Not (detailSubNodes is Nothing)) Then
                    For each detailSubNodeItem in detailSubNodes
                        requireDetail = ""
                        ierrStr = ""

                        ordGSeq = detailSubNodeItem.getElementsByTagName("ns1:ordGSeq")(0).Text	    '[ID:주문상품순번], 001
                        ordDSeq = detailSubNodeItem.getElementsByTagName("ns1:ordDSeq")(0).Text	    '[ID:주문상세순번], 001
                        ordWSeq = detailSubNodeItem.getElementsByTagName("ns1:ordWSeq")(0).Text	    '[ID:주문처리순번], 001

                        wbClsCd = detailSubNodeItem.getElementsByTagName("ns1:wbClsCd")(0).Text	        ' 주문정보 - 진행구분코드
                        ''------------------------------------------------------------------------------------------------------------
                        ordDtlClsCd = detailSubNodeItem.getElementsByTagName("ns1:ordDtlClsCd")(0).Text	    ' 주문정보 - 주문구분코드, 10
                        wbCrtDt = detailSubNodeItem.getElementsByTagName("ns1:wbCrtDt")(0).Text	            ' 주문정보 - 지시일자, 2013-05-22+09:00
                        ''outwConfDt	'주문정보 - 출고확정일자
                        ''delivDtm	    '주문정보 - 배송완료일
                        ''cnclInsDtm	'주문정보 - 취소일자
                        ''oldordNo	    '주문정보 - 원주문번호
                        toutYn = detailSubNodeItem.getElementsByTagName("ns1:toutYn")(0).Text	            '주문정보 - 기출하구분(Y-기출하,N-정상출하), N
                        chnNm = detailSubNodeItem.getElementsByTagName("ns1:chnNm")(0).Text	                '주문정보 - 채널구분, INTERNET
                        receverNm = detailSubNodeItem.getElementsByTagName("ns1:receverNm")(0).Text	        '주문정보 - 인수자, 채현아
                        'recvName	    '주문정보 - 수취인 영문명
                        zipno = detailSubNodeItem.getElementsByTagName("ns1:zipno")(0).Text	                '주문정보 - 우편번호, 110809
                        addr_1 = detailSubNodeItem.getElementsByTagName("ns1:addr_1")(0).Text	            '주문정보 - 주소, 서울 종로구 동숭동
                        addr_2 = detailSubNodeItem.getElementsByTagName("ns1:addr_2")(0).Text	            '주문정보 - 상세주소, 1-45번지 자유빌딩 6층
                        'addr	        '주문정보 - 주소
                        telno = detailSubNodeItem.getElementsByTagName("ns1:telno")(0).Text	                '주문정보 - 인수자tel, 02)973-8514
                        cellno = detailSubNodeItem.getElementsByTagName("ns1:cellno")(0).Text	            '주문정보 - 인수자HP, 010)2715-8514
                        'msgSpec	    '주문정보 - 배송참고
                        'delvplnDt	    '주문정보 - 배송예정일
                        packYn = detailSubNodeItem.getElementsByTagName("ns1:packYn")(0).Text	            '상품정보 - 세트여부, 일반
                        'itemNm	        '상품정보 - 세트상품명
                        itemCd = detailSubNodeItem.getElementsByTagName("ns1:itemCd")(0).Text	            '상품정보 - 판매코드, 21899852
                        unitCd = detailSubNodeItem.getElementsByTagName("ns1:unitCd")(0).Text	            '상품정보 - 단품코드, 10047125217
                        itemName = detailSubNodeItem.getElementsByTagName("ns1:itemName")(0).Text	        '상품정보 - 판매상품명, 24K Gold 전자파차단스티커
                        unitNm = detailSubNodeItem.getElementsByTagName("ns1:unitNm")(0).Text	            '상품정보 - 단품상세, ES-01 잘될꺼야
                        contItemCd = detailSubNodeItem.getElementsByTagName("ns1:contItemCd")(0).Text	    '상품정보 - 협력사상품코드, 279751_0011
                        wbIdNo = detailSubNodeItem.getElementsByTagName("ns1:wbIdNo")(0).Text	            '상품정보 - 운송장식별번호, 20000420537940
                        outwQty = detailSubNodeItem.getElementsByTagName("ns1:outwQty")(0).Text	            '상품정보 - 수량, 1.0
                        realslAmt = detailSubNodeItem.getElementsByTagName("ns1:realslAmt")(0).Text	        '상품정보 - 판매가, 1800.0
                        outwAmt = detailSubNodeItem.getElementsByTagName("ns1:outwAmt")(0).Text	            '상품정보 - 고객결제가, 1800.0  :: 수량*판매가 인지, 수량*실판매가인지 확인
                        'delivInfo	    '기타정보 - 비고
                        'promGiftSpec	'기타정보 - 사은품내역
                        'juminNum       '주문정보-주민번호(아님), 발송 금지!
                        'cnclRsn	    '기타정보 - 교환/취소사유
                        'cnclRsnSpec	'기타정보 - 교환/취소사유상세
                        ordDtm = detailSubNodeItem.getElementsByTagName("ns1:ordDtm")(0).Text	            '주문정보-주문일시, 2013-05-22 15:05:02


                        ''필수로 안넘어오는정보들.
                        wbCls =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0) Is Nothing)) Then
                            wbCls = detailSubNodeItem.getElementsByTagName("ns1:wbCls")(0).Text       '주문정보 - 진행구분
                        end if

                        confirmChk =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0) Is Nothing)) Then
                            confirmChk = detailSubNodeItem.getElementsByTagName("ns1:confirmChk")(0).Text        '주문정보 - 협력사실제회수확인 0,1
                        end if
                        ''-------------------------------------------------------------------------------------------------------------------------
                        cnclInsDtm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0) Is Nothing)) Then
                            cnclInsDtm = detailSubNodeItem.getElementsByTagName("ns1:cnclInsDtm")(0).Text        '주문정보 - 취소일자
                        end if
                        oldordNo =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0) Is Nothing)) Then
                            oldordNo = detailSubNodeItem.getElementsByTagName("ns1:oldordNo")(0).Text        '주문정보 - 원주문번호
                        end if
                        recvName =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:recvName")(0) Is Nothing)) Then
                            recvName = detailSubNodeItem.getElementsByTagName("ns1:recvName")(0).Text        '주문정보 - 수취인 영문명
                        end if
                        addr =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:addr")(0) Is Nothing)) Then
                            addr = detailSubNodeItem.getElementsByTagName("ns1:addr")(0).Text        '주문정보 - 주소all?
                        end if
                        msgSpec =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0) Is Nothing)) Then
                            msgSpec = detailSubNodeItem.getElementsByTagName("ns1:msgSpec")(0).Text        '주문정보 -배송참고
                        end if
                        delvplnDt =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0) Is Nothing)) Then
                            delvplnDt = detailSubNodeItem.getElementsByTagName("ns1:delvplnDt")(0).Text        '주문정보 -배송예정일
                        end if
                        itemNm =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0) Is Nothing)) Then
                            itemNm = detailSubNodeItem.getElementsByTagName("ns1:itemNm")(0).Text        '상품정보 -세트상품명
                        end if
                        juminNum =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0) Is Nothing)) Then
                            juminNum = detailSubNodeItem.getElementsByTagName("ns1:juminNum")(0).Text       '주문정보-주민번호(아님), 발송 금지!
                        end if
                        dccouponCjhs =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0) Is Nothing)) Then
                            dccouponCjhs = detailSubNodeItem.getElementsByTagName("ns1:dccouponCjhs")(0).Text       '주문정보 - 할인(당사부담)금액
                        end if
                        dccouponVendor =""
                        if (Not (detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0) Is Nothing)) Then
                            dccouponVendor = detailSubNodeItem.getElementsByTagName("ns1:dccouponVendor")(0).Text      '주문정보 - 할인(협력사부담)금액
                        end if

                        orderDlvPay = custDeliveryCost
                        if (dtlCnt>1) then
                            orderDlvPay = 0 ''첫번째 값만 넣음.
                        end if

                        orderCsGbn = ""
                        if (ordDtlClsCd="10") then
                            orderCsGbn="0"
                        elseif (ordDtlClsCd="20") then
                            orderCsGbn="2"  ''취소
                        end if

                        requireDetail = juminNum '' 주문제작문구
                        ierrCode = 0
                        ierrStr  = ""

                        if (orderCsGbn<>"") then
                            retVal = saveCsOneTmp(ordNo,ordDtm,splitvalue(contItemCd,"_",0),splitvalue(contItemCd,"_",1),itemName, unitNm _
                                    , custNm , custTelNo, custTelNo _
                                    , receverNm, telno, cellno, LEFT(zipno,3)&"-"&Right(zipno,3), addr_1, addr_2 _
                                    , realslAmt, realslAmt, CLNG(outwQty), ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq _
                                    , msgSpec, requireDetail, orderDlvPay, orderCsGbn _
                                    , ierrCode, ierrStr)
                        else
                            retVal = false
                            ierrStr = "주문구분 [ordDtlClsCd="&ordDtlClsCd&"] 정의되지 않음"
                        end if

                        dtlCnt = dtlCnt+1

                        if (retVal) then
                            succCnt = succCnt+1
                        else
                            failCnt = failCnt+1
                            if (ierrCode=-1) then skipCnt = skipCnt+1

                            if (mode="ORDCANCELLIST") then
                                rw "<font color='red'>["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"]</font> "&ierrStr &" / "&custNm
                            else
                                rw "["&ordNo&"-"&ordGSeq&"-"&ordDSeq&"-"&ordWSeq&"] "&ierrStr &" / "&custNm
                            end if
                        end if

                    Next
                end if

                Set detailSubNodes = Nothing
            Next
        end if
    end if

    Set Nodes = Nothing
    rw succCnt & "건 입력"
    rw failCnt & "건 실패" & "("&skipCnt&" 건 skip)"

End function

function saveORDOneTmp(OutMallOrderSerial,SellDate,matchItemID,matchItemOption,partnerItemName,partnerOptionName _
        , OrderName, OrderTelNo, OrderHpNo _
        , ReceiveName, ReceiveTelNo, ReceiveHpNo, ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 _
        , SellPrice, RealSellPrice, ItemOrderCount, OrgDetailKey _
        , deliverymemo, requireDetail, orderDlvPay, orderCsGbn _
        , byref ierrCode, byref ierrStr )
    dim paramInfo, retParamInfo
    dim SellSite : SellSite = "cjmall"
    dim PayType  : PayType  = "50"
    dim sqlStr
	dim countryCode

	if countryCode="" then countryCode="KR"

    saveORDOneTmp =false

    OrderTelNo = replace(OrderTelNo,")","-")
    OrderHpNo = replace(OrderHpNo,")","-")
    ReceiveTelNo = replace(ReceiveTelNo,")","-")
    ReceiveHpNo = replace(ReceiveHpNo,")","-")

    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        ,Array("@SellSite" , adVarchar	, adParamInput, 32, SellSite)	_
		,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,32, OutMallOrderSerial)	_
		,Array("@SellDate"	,adDate, adParamInput,, SellDate) _
		,Array("@PayType"	,adVarchar, adParamInput,32, PayType) _
		,Array("@Paydate"	,adDate, adParamInput,, SellDate) _
		,Array("@matchItemID"	,adInteger, adParamInput,, matchItemID) _
		,Array("@matchItemOption"	,adVarchar, adParamInput,4, matchItemOption) _
		,Array("@partnerItemID"	,adVarchar, adParamInput,32, matchItemID) _
		,Array("@partnerItemName"	,adVarchar, adParamInput,128, partnerItemName) _
		,Array("@partnerOption"	,adVarchar, adParamInput,128, matchItemOption) _
		,Array("@partnerOptionName"	,adVarchar, adParamInput,128, partnerOptionName) _
		,Array("@OrderUserID"	,adVarchar, adParamInput,32, "") _
		,Array("@OrderName"	,adVarchar, adParamInput,32, OrderName) _
		,Array("@OrderEmail"	,adVarchar, adParamInput,100, "") _
		,Array("@OrderTelNo"	,adVarchar, adParamInput,16, OrderTelNo) _
		,Array("@OrderHpNo"	,adVarchar, adParamInput,16, OrderHpNo) _
		,Array("@ReceiveName"	,adVarchar, adParamInput,32, ReceiveName) _
		,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, ReceiveTelNo) _
		,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, ReceiveHpNo) _
		,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, ReceiveZipCode) _
		,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, ReceiveAddr1) _
		,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, ReceiveAddr2) _
		,Array("@SellPrice"	,adCurrency, adParamInput,, SellPrice) _
		,Array("@RealSellPrice"	,adCurrency, adParamInput,, RealSellPrice) _
		,Array("@ItemOrderCount"	,adInteger, adParamInput,, ItemOrderCount) _
		,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, OrgDetailKey) _
		,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
		,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
		,Array("@deliverymemo"	,adVarchar, adParamInput,400, deliverymemo) _
		,Array("@requireDetail"	,adVarchar, adParamInput,400, requireDetail) _
		,Array("@orderDlvPay"	,adCurrency, adParamInput,, orderDlvPay) _
		,Array("@orderCsGbn"	,adInteger, adParamInput,, orderCsGbn) _
    	,Array("@countryCode"	,adVarchar, adParamInput,2, countryCode) _
    	,Array("@outMallGoodsNo"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerMallName" ,adVarchar, adParamInput,64, "") _
    	,Array("@shoplinkerPrdCode"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerOrderID"	,adVarchar, adParamInput,16, "") _
    	,Array("@shoplinkerMallID"	,adVarchar, adParamInput,32, "") _
		,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    if (matchItemOption<>"") and (matchItemID<>"-1") and (matchItemID<>"") then
        sqlStr = "db_temp.dbo.sp_TEN_xSite_TmpOrder_Insert"
        retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

        ierrCode = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드
        ierrStr  = GetValue(retParamInfo, "@retErrStr")   ' 에러메세지
    else
        ierrCode = -999
        ierrStr = "상품코드 또는 옵션코드  매칭 실패" & OrgDetailKey & " 상품코드 =" & matchItemID&" 옵션명 = "&partnerOptionName
        rw "["&ierrCode&"]"&retErrStr
        dbget.close() : response.end
    end if

    saveORDOneTmp = (ierrCode=0)
end function

Function XMLSend(url, xmlStr)
	Dim poster, SendDoc, retDoc
	Set SendDoc = server.createobject("MSXML2.DomDocument.3.0")
		SendDoc.async = False
		SendDoc.LoadXML(xmlStr)

	Set poster = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		poster.open "POST", url, false
		poster.setRequestHeader "CONTENT_TYPE", "text/xml"
		poster.setTimeouts 5000,90000,90000,90000  ''2013/07/22 추가
		poster.send SendDoc

	Set retDoc = server.createobject("MSXML2.DomDocument.3.0")
		retDoc.async = False
		retDoc.LoadXML(poster.responseTEXT)

	'response.write poster.responseTEXT
	'response.end

	Set XMLSend = retDoc
	Set SendDoc = Nothing
	Set poster = Nothing
End Function

Function getCurrDateTimeFormat()
    dim nowtimer : nowtimer= timer()
    getCurrDateTimeFormat = left(now(),10)&"_"&nowtimer
End Function

Sub CheckFolderCreate(sFolderPath)
	Dim objfile
	Set objfile=Server.CreateObject("Scripting.FileSystemObject")
		IF NOT  objfile.FolderExists(sFolderPath) THEN
			objfile.CreateFolder sFolderPath
		END IF
	Set objfile=Nothing
End Sub

Function XMLFileSave(xmlStr, mode, iitemid)
   Exit function  ''로그 안남김

	Dim fso,tFile
	Dim opath
	Select Case mode
		Case "REG", "RET_REG"
			opath = "/admin/etc/cjmall/xmlFiles/INSERT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "LIST", "DayLIST", "RET_LIST", "RET_DayLIST", "commonCD", "RET_commonCD", "RET_SONGJANG", "cjItemCheck", "RET_cjItemCheck"
			opath = "/admin/etc/cjmall/xmlFiles/SELECT/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "EDT", "RET_EDT", "MDT", "RET_MDT", "PRI", "RET_PRI", "PRI2", "RET_PRI2", "QTY", "RET_QTY", "DateRes", "RET_DateRes"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
		Case "SLD", "RET_SLD"
			opath = "/admin/etc/cjmall/xmlFiles/UPDATE_SellStatus/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	    Case "RET_ORDLIST", "RET_ORDCANCELLIST", "RET_CSLIST"
	        opath = "/admin/etc/cjmall/xmlFiles/ORDER/"&year(now())&"-"&Format00(2,month(now()))&"-"&Format00(2,day(now()))&"/"
	End Select

	Dim defaultPath : defaultPath = server.mappath(opath) + "\"
	Dim fileName
	If mode = "LIST" or mode = "DayLIST" Then
		fileName = mode &"_"& getCurrDateTimeFormat& ".xml"
	Else
		fileName = mode &"_"& getCurrDateTimeFormat&"_"&iitemid&".xml"
	End If

	CALL CheckFolderCreate(defaultPath)
	''debug
	Set fso = CreateObject("Scripting.FileSystemObject")
		Set tFile = fso.CreateTextFile(defaultPath & FileName )
			tFile.Write(xmlStr)
			tFile.Close
		Set tFile = nothing
	Set fso = nothing
End Function

function getLastOrderInputDT()
    dim sqlStr
    sqlStr = "select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr&" from db_temp.dbo.tbl_XSite_TMpOrder"
    sqlStr = sqlStr&" where sellsite='cjmall'"
    sqlStr = sqlStr&" order by selldate desc"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastOrderInputDT = rsget("lastOrdInputDt")
	end if
	rsget.Close

end function

function getLastOrderInputDTUp()
    dim sqlStr
    sqlStr = " select top 1 convert(varchar(10),selldate,21) as lastOrdInputDt"
    sqlStr = sqlStr & " from db_temp.dbo.tbl_xSite_tmporder"
    sqlStr = sqlStr & " where sellsite='cjmall'"
    sqlStr = sqlStr & " and  convert(varchar(10),selldate,21)>("
    sqlStr = sqlStr & " 	select top 1 convert(varchar(10),selldate,21) as slDt from db_temp.dbo.tbl_xSite_tmporder"
    sqlStr = sqlStr & " 	where sellsite='cjmall'"
    sqlStr = sqlStr & " 	and tenCpnUint is Not NULL"
    sqlStr = sqlStr & " 	order by selldate desc"
    sqlStr = sqlStr & " ) order by selldate"

    rsget.Open sqlStr,dbget,1
	if (Not rsget.Eof) then
		getLastOrderInputDTUp = rsget("lastOrdInputDt")
	end if
	rsget.Close

    'getLastOrderInputDTUp="2013-06-14"
end function
%>
