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
Dim FileName: FileName = "10x10ItemJson.txt"
Dim fso, tFile, oFile, oZip, Schk, success

Set oFile = CreateObject("ADODB.Stream")
With oFile
	.Charset = "UTF-8"
	.Open
End With

	vQuery = "    select '{""pinfo"":{' + STUFF(( "
	vQuery = vQuery & "           select  "
	vQuery = vQuery & "            ',""'+cast(rank() OVER (ORDER BY t1.itemid) as nvarchar(max))+'"":{""pid"":""' + cast(t1.itemid as varchar(max)) + '""' "
	vQuery = vQuery & "             + ',""pnm"":""' + replace(replace(replace(replace(replace(replace(replace(replace(replace(t1.itemname,char(92),'\\'),char(47),'\/'),char(34),'\""'),char(13),'\r'),char(12),'\f'),char(11),'\v'),char(10),'\n'),char(9),'\t'),char(8),'\b')  + '""' "
	vQuery = vQuery & "             + ',""catlvl1"":""' + substring(cast(t2.catecode as nvarchar(max)), 1, 3) + '""' "
	vQuery = vQuery & "             + ',""catlvl2"":""' + substring(cast(t2.catecode as nvarchar(max)), 4, 3) + '""' "
	vQuery = vQuery & "             + ',""catlvl3"":""""' + '' "
	vQuery = vQuery & "             + ',""sale_price"":'+cast(convert(bigint, t1.sellcash) as nvarchar(max))+'' "
	vQuery = vQuery & "             + ',""original_price"":'+cast(convert(bigint, t1.orgprice) as nvarchar(max))+'' "
	vQuery = vQuery & "             + ',""pimg"":""""' + '' "
	vQuery = vQuery & "             + ',""purl"":""""' + '' "
	vQuery = vQuery & "             + ',""reg_date"":""'+convert(nvarchar(10), t1.regdate, 120)+'""' "
	vQuery = vQuery & "             + ',""update_date"":""'+convert(nvarchar(10), t1.lastupdate, 120)+'""' "
	vQuery = vQuery & "             + ',""pnm_en"":""""' + '' "
	vQuery = vQuery & "             + ',""pnm_cn"":""""' + '' "
	vQuery = vQuery & "             +'}'+char(13)+char(10) "
	vQuery = vQuery & "              from db_AppWish.[dbo].[tbl_item] t1 "
	vQuery = vQuery & "			  inner join db_AppWish.[dbo].[tbl_display_cate_item] t2 on t1.itemid = t2.itemid And t2.isdefault = 'y' "
	vQuery = vQuery & "			  Where t1.isusing='Y' And t1.itemid <> 0 And t1.sellyn <> 'N' "
	vQuery = vQuery & "              for xml path(''), type "
	vQuery = vQuery & "             ).value('.', 'nvarchar(max)'), 1, 1, '') + '}}' "


	dbCTget.CommandTimeOut = 480
	rsCTget.Open vQuery,dbCTget,1
	IF Not rsCTget.Eof Then
		oFile.WriteText rsCTget(0)
		oFile.SaveToFile appPath & FileName, 2
		Schk = "1"
		response.write "RecoBell JsonDataComplete<p>"
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

		success = oZip.NewZip(appPath&"10x10ItemJson.zip")
		If (success <> 1) Then
			Response.Write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"

		End If

		success = oZip.AppendOneFileOrDir(""&appPath&FileName&"", 0)
		If success <> 1 Then
			response.write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"
			response.End
		End If

		success = oZip.QuickAppend(appPath&"10x10ItemJson.zip")
		If (success <> 1) Then
			Response.Write "<pre>" & Server.HTMLEncode( oZip.LastErrorText) & "</pre>"
			response.End
		End If

		response.write "RecoBell JsonDataZipCreateComplete"
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