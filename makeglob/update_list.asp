<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim sqlStr, objXML, i, j
	Dim rsCd1, rsCd2

'	파일 저장시
'	Dim savePath, FileName
'	savePath = server.mappath("/makeglob/") + "\"
'	FileName = "update_list.xml"

	sqlStr = " Select product_key, product_code, product_language, convert(varchar(19), lastupdate, 120) as lastupdate From db_item.dbo.tbl_makeglob_product Where makeglobYN='N' order by product_key "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) THEN
		rsCd1 = rsget.getRows()
	Else
		Response.Write "상품 리스트 없음."
		Response.End
	END IF
	rsget.close




	'// XML 데이터 생성
	 Set objXML = server.CreateObject("Microsoft.XMLDOM")
	 objXML.async = False

	'----- XML 해더 생성
	objXML.appendChild(objXML.createProcessingInstruction("xml","version=""1.0"" encoding=""utf-8"""))
	objXML.appendChild(objXML.createElement("root"))
	objXML.documentElement.appendChild(objXML.createElement("version"))
	objXML.documentElement.childNodes(0).text = "1.0"
	objXML.documentElement.appendChild(objXML.createElement("shop_id"))
	objXML.documentElement.childNodes(1).text = "tenbyten1010"
	objXML.documentElement.appendChild(objXML.createElement("modified"))
	objXML.documentElement.childNodes(2).text = rsCd1(3,0)
	objXML.documentElement.appendChild(objXML.createElement("cate_url"))
	objXML.documentElement.childNodes(3).text = "http://wapi.10x10.co.kr/makeglob/category.asp"
	objXML.documentElement.appendChild(objXML.createElement("item_url"))
	objXML.documentElement.childNodes(4).text = "http://wapi.10x10.co.kr/makeglob/product_list/product_info.asp?code={{product_key}}"

	'카테고리 정보 생성
	objXML.documentElement.appendChild(objXML.createElement("update_index"))

	for i=0 to Ubound(rsCd1,2)
		objXML.documentElement.childNodes(5).appendChild(objXML.createElement("product_key"))
		'objXML.documentElement.childNodes(4).childNodes(i).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
		objXML.documentElement.childNodes(5).childNodes(i).text = rsCd1(0,i)
	next

	'// XML 결과 출력
	Response.Clear
	response.Charset="UTF-8"
	Response.ContentType = "text/xml"
	Response.Write objXML.xml

	'// 파일 저장시
	''objXML.save(savePath & FileName)
	''rstMsg = "데이터 파일 [" & FileName & "] 생성 완료"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->