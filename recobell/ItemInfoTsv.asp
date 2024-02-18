<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% response.Charset="UTF-8" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp" -->
<%
	Dim vQuery, objFso, objFile

Dim appPath : appPath = server.mappath("/recobell/") + "\"
Dim FileName: FileName = "10x10ItemTsv.tsv"
Dim fso, tFile, oFile, oZip, Schk, success, vTxtValue

vTxtvalue=""

Set oFile = CreateObject("ADODB.Stream")
With oFile
	.Charset = "UTF-8"
	.Open
End With

	vQuery = "    Select char(13)+char(10)+stuff(( "
	vQuery = vQuery & " Select "
	vQuery = vQuery & " + '""'+cast(t1.itemid as varchar(max))+'""'+char(9) "
	vQuery = vQuery & " + '""'+replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(t1.itemname,char(92),'\\'),char(47),'\/'),char(34),'\""'),char(13),'\r'),char(12),'\f'),char(11),'\v'),char(10),'\n'),char(9),'\t'),char(8),'\b'), '\v',''),'\t','') +'""'+char(9) "
	vQuery = vQuery & " + '""'+substring(cast(t2.catecode as nvarchar(max)), 1, 3)+'""'+char(9) "
	vQuery = vQuery & " + '""'+substring(cast(t2.catecode as nvarchar(max)), 4, 3)+'""'+char(9) "
	vQuery = vQuery & " + '""""'+char(9) "
	vQuery = vQuery & " + '""'+cast(convert(bigint, t1.sellcash) as nvarchar(max))+'""'+char(9) "
	vQuery = vQuery & " + '""'+cast(convert(bigint, t1.orgprice) as nvarchar(max))+'""'+char(9) "
	vQuery = vQuery & " 	+ '""http://webimage.10x10.co.kr/image/icon1/'+case when len(cast(t1.itemid/10000 as varchar(max)))=1 then '0'+cast(t1.itemid/10000 as varchar(max)) else cast(t1.itemid/10000 as varchar(max)) end+'/'+convert(nvarchar(50), t1.icon1image, 120)+'""'+char(9) "
	vQuery = vQuery & " + '""""'+char(9) "
	vQuery = vQuery & " + '""'+convert(nvarchar(10), t1.regdate, 120)+'""'+char(9) "
	vQuery = vQuery & " + '""'+convert(nvarchar(10), t1.lastupdate, 120)+'""'+char(9) "
	vQuery = vQuery & " + '""""'+char(9) "
	vQuery = vQuery & " + '""""'+char(9) "
	vQuery = vQuery & " +char(13)+char(10)"
	vQuery = vQuery & " from db_AppWish.[dbo].[tbl_item] t1  "
	vQuery = vQuery & " inner join db_AppWish.[dbo].[tbl_display_cate_item] t2 on t1.itemid = t2.itemid And t2.isdefault = 'y'  "
	vQuery = vQuery & " Where t1.isusing='Y' And t1.itemid <> 0 And t1.sellyn <> 'N' "
	vQuery = vQuery & " for xml path(''), type).value('.','nvarchar(max)'), 1, 0, '') "

	dbCTget.CommandTimeOut = 480
	rsCTget.Open vQuery,dbCTget, adOpenForwardOnly, adLockReadOnly
	IF Not rsCTget.Eof Then
		oFile.WriteText "pid	pnm	catlv11	catlv12	catlv13	sale_price	original_price	pimg	purl	reg_date	update_date	pnm_en	pnm_cn"&rsCTget(0)
		oFile.SaveToFile appPath & FileName, 2
		response.write "RecoBell TsvDataComplete<p>"
		Schk = "1"
	End IF
	rsCTget.close
    oFile.Close
	Set oFile = Nothing



	If Schk = "1" Then 
		Set oZip = Server.CreateObject("Chilkat.Zip2")
		success =	oZip.UnlockComponent("10X10CZIP_4HmoweDQnXfy")

		If success <> 1 Then
			response.write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"
			response.End
		End If

		success = oZip.NewZip(appPath&"10x10ItemTsv.zip")
		If (success <> 1) Then
			Response.Write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"

		End If

		success = oZip.AppendOneFileOrDir(""&appPath&FileName&"", 0)
		If success <> 1 Then
			response.write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"
			response.End
		End If



		response.write "RecoBell TsvDataZipCreateComplete"
	End If
	success = oZip.WriteZipAndClose()
	If success <> 1 Then
		response.write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"
		response.End
	End If

'	oZip.close
	Set oZip = Nothing

%>
<!-- #include virtual="/lib/db/dbCTclose.asp" -->