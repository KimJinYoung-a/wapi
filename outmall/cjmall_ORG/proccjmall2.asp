<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/incOutMallCommonFunction.asp"-->
<%
'// 저장 모드 접수
Dim mode, sqlStr, iErrMsg
mode = Request("mode")

'// 상품번호/옵션번호를 받는다 //
Dim dispNo, cdl, cdm, cds, infodiv, CdmKey
dispNo	= requestCheckvar(Request("dspNo"),32)
cdl		= requestCheckvar(Request("cdl"),10)
cdm		= requestCheckvar(Request("cdm"),10)
cds		= requestCheckvar(Request("cds"),10)
infodiv	= requestCheckvar(Request("infodiv"),10)
CdmKey	= requestCheckvar(Request("CdmKey"),10)

If (mode = "saveCate") OR (mode = "delGbn") OR (mode = "delPrddiv") Then
	If (dispNo = "" ) OR cdl = "" OR cdm = "" OR cds = "" Then
		Call Alert_move("전송된 값이 없습니다.\n처리가 종료되었습니다.","about:blank")
		dbget.Close: response.End
	End If
End If

'// 모드별 분기
Select Case mode
	Case "saveCate"
        '중복 확인
        sqlStr = "Select CddKey From db_etcmall.dbo.tbl_cjMall_prdDiv_mapping "  & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'"  & VbCrlf
		sqlStr = sqlStr& " 	and CddKey='" & dispNo & "'"
		rsget.Open sqlStr,dbget,1

		If rsget.EOF Then
			'신규등록
			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_etcmall.dbo.tbl_cjMall_prdDiv_mapping  " & VbCrlf
			sqlStr = sqlStr & " (CddKey, infodiv, CdmKey, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '"&infodiv&"', '"&CdmKey&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbget.execute sqlStr

			sqlStr = ""
			sqlStr = sqlStr & " Insert into db_outmall.dbo.tbl_cjMall_prdDiv_mapping  " & VbCrlf
			sqlStr = sqlStr & " (CddKey, infodiv, CdmKey, tenCateLarge, tenCateMid, tenCateSmall, lastUpdate)" & VbCrlf
			sqlStr = sqlStr & " VALUES('" & dispNo & "'"  & VbCrlf
			sqlStr = sqlStr & ", '"&infodiv&"', '"&CdmKey&"', '" & cdl & "','" & cdm & "','" & cds & "', getdate()) "
			dbCTget.execute sqlStr
		Else
		    iErrMsg = "이미 매핑된 상품분류는 ["&dispNo&"] 추가할 수 없습니다."
		End If
		rsget.Close

	Case "delPrddiv"
		'매칭된 텐바이텐 카테고리 삭제
		sqlStr = "Delete From db_etcmall.dbo.tbl_cjMall_prdDiv_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and CddKey='" & dispNo & "'"
		sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"
		dbget.execute(sqlStr)

		sqlStr = "Delete From db_outmall.dbo.tbl_cjMall_prdDiv_mapping " & VbCrlf
		sqlStr = sqlStr& " Where tenCateLarge='" & cdl & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateMid='" & cdm & "'" & VbCrlf
		sqlStr = sqlStr& " 	and tenCateSmall='" & cds & "'" & VbCrlf
		sqlStr = sqlStr& " 	and CddKey='" & dispNo & "'"
		sqlStr = sqlStr& " 	and infodiv='" & infodiv & "'"
		dbCTget.execute(sqlStr)
End Select

%>
<script language="javascript">
<% If (iErrMsg<>"") Then %>
alert("<%=iErrMsg %>");
<% Else %>
alert("정상적으로 처리되었습니다.");
parent.opener.history.go(0);
parent.self.close();
<% End If %>
</script>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->