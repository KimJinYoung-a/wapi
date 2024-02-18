<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 카카오기프트 배송비 수정
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbDatamartopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim nowDate, sqlStr
nowDate = LEFT(Date(), 7)

Dim cnt, i
Dim orderserialArr, idxArr, orgitemcostArr
sqlStr = ""
sqlStr = sqlStr & " SELECT d.orderserial, d.idx, d.itemid, d.itemoption, d.makerid, d.itemname, d.itemoptionname, d.itemno, d.orgitemcost, d.itemcostCouponNotApplied "
sqlStr = sqlStr & " ,d.itemcost, d.reducedPrice, d.omwdiv, d.beasongdate,d.dlvfinishdt, d.jungsanfixdate "
sqlStr = sqlStr & " ,i.orgprice,i.sellcash, i.mwdiv "
sqlStr = sqlStr & " FROM db_replica.dbo.tbl_order_master m WITH(NOLOCK) "
sqlStr = sqlStr & " JOIN db_replica.dbo.tbl_order_detail d WITH(NOLOCK) on m.orderserial=d.orderserial "
sqlStr = sqlStr & " JOIN db_item.dbo.tbl_item i  WITH(NOLOCK) on d.itemid=i.itemid "
sqlStr = sqlStr & " WHERE m.sitename='kakaogift' "
sqlStr = sqlStr & " and m.cancelyn='N' "
sqlStr = sqlStr & " and d.cancelyn<>'Y' "
sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) >= '"& nowDate &"-01' "
sqlStr = sqlStr & " and isNULL(d.jungsanFixdate,d.beasongdate) < convert(varchar(10),dateadd(m,1, '"& nowDate &"-01'),121) "
sqlStr = sqlStr & " and i.orgprice<>d.itemcost "
sqlStr = sqlStr & " and i.sellcash<>d.itemcost "
sqlStr = sqlStr & " and d.itemid not in (0,100) "
sqlStr = sqlStr & " and d.itemcost - d.orgitemcost in ('2500', '3000') "
sqlStr = sqlStr & " ORDER BY CASE WHEN (d.itemcost - d.orgitemcost = 3000 or d.itemcost - d.orgitemcost = 2500) THEN 2 ELSE 1 END DESC,	2,1 desc "
dbDatamart_rsget.CursorLocation = adUseClient
dbDatamart_rsget.Open sqlStr, dbDatamart_dbget, adOpenForwardOnly, adLockReadOnly
cnt = dbDatamart_rsget.RecordCount
ReDim orderserialArr(cnt)
ReDim idxArr(cnt)
ReDim orgitemcostArr(cnt)
i = 0
If Not dbDatamart_rsget.Eof Then
	Do Until dbDatamart_rsget.eof
		orderserialArr(i)   = dbDatamart_rsget("orderserial")
		idxArr(i)           = dbDatamart_rsget("idx")
		orgitemcostArr(i)   = dbDatamart_rsget("orgitemcost")
		i=i+1
		dbDatamart_rsget.MoveNext
	Loop
End If
dbDatamart_rsget.close

If (cnt < 1) Then
	response.Write "S_NONE.."
	dbDatamart_dbget.Close() : response.end
Else
	rw "CNT="&CNT
	For i = LBound(orderserialArr) To UBound(orderserialArr)
		If (orderserialArr(i) <> "") Then
			sqlStr = ""
			sqlStr = sqlStr & " EXEC [db_jungsan].[dbo].[usp_Ten_OUTAMLL_KakaoOrderEdit_WithDlvPrice] '"& orderserialArr(i) &"', '"& idxArr(i) &"', "& orgitemcostArr(i) &", '' "
			dbget.Execute(sqlStr)
			rw sqlStr
		End If
	Next
End If

%>
<!-- #include virtual="/lib/db/dbDatamartClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->