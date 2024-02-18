<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.charset = "utf-8" %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/search/searchcls.asp" -->
<%
DIM oPpkDoc, arrList, arrTg, iRows
	SET oPpkDoc = NEW SearchItemCls
		oPpkDoc.FPageSize = 30
		arrList = oPpkDoc.getPopularKeyWords()					'인기검색어 일반형태
		'oPpkDoc.getPopularKeyWords2 arrList,arrTg				'인기검색어 순위정보 포함
	SET oPpkDoc = NOTHING

	IF isArray(arrList)  THEN
		if Ubound(arrList)>0 then
			FOR iRows =0 To UBOUND(arrList)
				Response.Write arrList(iRows) &"<br>"& vbCrLf
			Next
		END IF
	END IF
%>