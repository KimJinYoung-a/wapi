<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),20)
Dim arrItemid : arrItemid = request("cksel")
Dim chkXML : chkXML = request("chkXML")
Dim i, strParam, iErrStr, ret1
Dim sqlStr, strSql, AssignedRow, SubNodes
Dim chgSellYn, actCnt, retErrStr
Dim buf, buf2, CNT10, CNT20, CNT30, iitemid
Dim ArrRows
Dim retFlag
Dim iMessage
dim iItemName, pregitemname

retFlag   = request("retFlag")
chgSellYn = request("chgSellYn")
arrItemid = Trim(arrItemid)

If cmdparam = "cjmallCommonCode" Then
	Dim ccd
	ccd = request("CommCD")
	call getcjCommonCodeList(ccd)
	response.end
End If

If cmdparam = "CategoryView" Then
	call getcjCategoryView()
	response.end
ElseIf cmdparam = "DivView" Then
	call getcjDivView()
	response.end
ElseIf cmdparam = "DivCodeView" Then
	call getcjDivCodeView()
	response.end
End If

Function getcjCategoryView()
	Dim strParam, strRst
    strRst = ""
    strRst = strRst &"<?xml version=""1.0""?>"
    strRst = strRst &"<tns:ifRequest xsi:schemaLocation=""http://www.example.org/ifpa ../IF_01_02.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_01_02"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    strRst = strRst &"</tns:ifRequest>"
	strParam = strRst

    iErrStr = ""
	Call cjmallCommCd(strParam, iErrStr, "CATE")
End Function

Function getcjDivView()
	Dim strParam, strRst
    strRst = ""
    strRst = strRst &"<?xml version=""1.0""?>"
    strRst = strRst &"<tns:ifRequest xsi:schemaLocation=""http://www.example.org/ifpa ../IF_01_01.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_01_01"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    strRst = strRst &"</tns:ifRequest>"
	strParam = strRst

    iErrStr = ""
	Call cjmallCommCd(strParam, iErrStr, "DIVVIEW")
End Function

Function getcjDivCodeView()
	Dim strParam, strRst
    strRst = ""
    strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
    strRst = strRst &"<tns:ifRequest xsi:schemaLocation=""http://www.example.org/ifpa ../IF_01_06.xsd"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_01_06"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    strRst = strRst &"<tns:allYn>Y</tns:allYn>"
    strRst = strRst &"</tns:ifRequest>"
	strParam = strRst

    iErrStr = ""
	Call cjmallCommCd(strParam, iErrStr, "DIVCODEVIEW")
End Function



Function getcjCommonCodeList(ccd)
	Dim AssignedRow
	Dim strParam, strRst
    On Error Resume Next
	    strRst = ""
        strRst = strRst &"<?xml version=""1.0"" encoding=""UTF-8""?>"
        strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_02_01"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_02_01.xsd"">"
        strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
        strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
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

Function cjmallCommCd(strParam, byRef iErrStr, v)
	Dim xmlStr : xmlStr = strParam
	Dim objXML, xmlDOM, Nodes, SubNodes
	Dim cjMallAPIURL, strSql
	Dim itemtypeCd, lrgNm, midNm, smNm, dtlNm
	IF application("Svr_Info")="Dev" THEN
		cjMallAPIURL = "http://210.122.101.154:8110/IFPAServerAction.action"	'' 테스트서버
	Else
		cjMallAPIURL = "http://api.cjmall.com/IFPAServerAction.action"			'' 실서버
	End if

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
		If v = "DIVVIEW" Then
			Set Nodes = xmlDOM.getElementsByTagName("ns1:itemCategory")
			If (Not (xmlDOM is Nothing)) Then
				strSql = ""
				strSql = " DELETE FROM db_temp.[dbo].[tbl_cjmallMng_category] "
				dbget.Execute(strSql)
				For each SubNodes in Nodes
					' rw SubNodes.getElementsByTagName("ns1:itemTypeCd")(0).Text
					' rw SubNodes.getElementsByTagName("ns1:itemLgrpNm")(0).Text
					' rw SubNodes.getElementsByTagName("ns1:itemMgrpNm")(0).Text
					' rw SubNodes.getElementsByTagName("ns1:itemSgrpNm")(0).Text
					' rw SubNodes.getElementsByTagName("ns1:itemTgrpNm")(0).Text
					' rw "-----------"
					itemtypeCd = SubNodes.getElementsByTagName("ns1:itemTypeCd")(0).Text
					lrgNm = SubNodes.getElementsByTagName("ns1:itemLgrpNm")(0).Text
					midNm = SubNodes.getElementsByTagName("ns1:itemMgrpNm")(0).Text
					smNm = SubNodes.getElementsByTagName("ns1:itemSgrpNm")(0).Text
					dtlNm = SubNodes.getElementsByTagName("ns1:itemTgrpNm")(0).Text

					strSql = ""
					strSql = strSql & "	INSERT INTO db_temp.[dbo].[tbl_cjmallMng_category] "
					strSql = strSql & "	(itemtypeCd, lrgNm, midNm, smNm, dtlNm) VALUES "
					strSql = strSql & "	('"&itemtypeCd&"', '"&lrgNm&"', '"&midNm&"', '"&smNm&"', '"&dtlNm&"') "
					dbget.Execute(strSql)
				Next
				rw "완료"
			End If
			Set Nodes = Nothing
			Set xmlDOM = Nothing
			Set objXML = Nothing
		Else
			response.write objXML.ResponseText
			Set xmlDOM = Nothing
			Set objXML = Nothing
		End If
		On Error Goto 0
	End If
End Function
%>
<script type="text/javascript">
	var items = "<%=arrItemid%>";
	var itemArr = items.split(", ");
	var rotation;
	var rno = 0;

	function loadRotation() {
		if(itemArr[rno] == undefined){
			//alert('완료하였습니다')
			window.parent.postMessage({
				action: "systemAlert"
				, message: "완료하였습니다"
			}, "*");
			return;
		}

		rotation = arrSubmit(itemArr[rno]);
		rno++;
		if(rno > itemArr.length-1){
			clearTimeout(rotation);
			//setTimeout("alert('완료하였습니다')", 500);
		}else{
			//setTimeout('loadRotation()', 2000);
		}
	}

	function arrSubmit(ino){
		document.frmSvArr.target = "xLink2";
        document.frmSvArr.act.value = "<%=cmdparam%>";
        document.frmSvArr.itemid.value = ino;
		document.frmSvArr.chgSellYn.value = "<%=chgSellYn%>";
		document.frmSvArr.chkXML.value = "<%=chkXML%>";
        document.frmSvArr.action = '/outmall/cjmall/cjmallActProc.asp';
        document.frmSvArr.submit();
	}
	window.onload = new Function('setTimeout("loadRotation()", 200)');
</script>
<form name="frmSvArr">
	<input type="hidden" name="act">
	<input type="hidden" name="itemid">
	<input type="hidden" name="chgSellYn">
	<input type="hidden" name="chkXML">
</form>

<div id="actStr"></div>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
