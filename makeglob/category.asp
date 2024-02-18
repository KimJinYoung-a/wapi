<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim sqlStr, objXML, i, j, k
	Dim rsCd1, rsCd2, rsCd3

'	파일 저장시
'	Dim savePath, FileName
'	savePath = server.mappath("/makeglob/") + "\"
'	FileName = "category.xml"

	'''Rows :        0          1       2       3
	sqlStr = "select MG_cateCd, cateNm, sortNo, (select convert(varchar(19),max(lastupdate),21) from db_item.dbo.tbl_makeglob_Category) as lastupdate " &_
			" from db_item.dbo.tbl_makeglob_Category " &_
			" where depth=1 and isUsing='Y' " &_
			" group by MG_cateCd, cateNm, sortNo " &_
			" order by sortNo, MG_cateCd"
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) THEN
		rsCd1 = rsget.getRows()
	Else
		Response.Write "생성할 카테고리 없음"
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
	objXML.documentElement.appendChild(objXML.createElement("category_language"))
	objXML.documentElement.childNodes(3).text = "KO"

	'카테고리 정보 생성
	objXML.documentElement.appendChild(objXML.createElement("category"))

	for i=0 to Ubound(rsCd1,2)
		objXML.documentElement.childNodes(4).appendChild(objXML.createElement("cate1"))

			objXML.documentElement.childNodes(4).childNodes(i).appendChild(objXML.createElement("name"))
			objXML.documentElement.childNodes(4).childNodes(i).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
			objXML.documentElement.childNodes(4).childNodes(i).childNodes(0).childNodes(0).text = rsCd1(1,i)

			objXML.documentElement.childNodes(4).childNodes(i).appendChild(objXML.createElement("cate_key"))
			objXML.documentElement.childNodes(4).childNodes(i).childNodes(1).text = rsCd1(0,i)
			objXML.documentElement.childNodes(4).childNodes(i).appendChild(objXML.createElement("attr"))
			objXML.documentElement.childNodes(4).childNodes(i).childNodes(2).text = ""

			'// 2depth 추가
			'''Rows :        0          1
			sqlStr = "select MG_cateCd, cateNm " &_
					" from db_item.dbo.tbl_makeglob_Category " &_
					" where parentCd=" & rsCd1(0,i) & " and depth=2 and isUsing='Y' " &_
					" order by sortNo, MG_cateCd"
			rsget.CursorLocation = adUseClient
			rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
			IF Not (rsget.EOF OR rsget.BOF) THEN
				rsCd2 = rsget.getRows()
			END IF
			rsget.close

			IF isArray(rsCd2) THEN
				objXML.documentElement.childNodes(4).childNodes(i).appendChild(objXML.createElement("cate1_sub"))

				for j=0 to Ubound(rsCd2,2)
					objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).appendChild(objXML.createElement("cate2"))

						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).appendChild(objXML.createElement("name"))
						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(0).childNodes(0).text = rsCd2(1,j)

						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).appendChild(objXML.createElement("cate_key"))
						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(1).text = rsCd2(0,j)
						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).appendChild(objXML.createElement("attr"))
						objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(2).text = ""

						'// 3depth 추가
						'''Rows :        0          1
						sqlStr = "select MG_cateCd, cateNm " &_
								" from db_item.dbo.tbl_makeglob_Category " &_
								" where parentCd=" & rsCd2(0,j) & " and depth=3 and isUsing='Y' " &_
								" order by sortNo, MG_cateCd"
						rsget.CursorLocation = adUseClient
						rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
						IF Not (rsget.EOF OR rsget.BOF) THEN
							rsCd3 = rsget.getRows()
						END IF
						rsget.close

						IF isArray(rsCd3) THEN
							objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).appendChild(objXML.createElement("cate2_sub"))

							for k=0 to Ubound(rsCd3,2)
								objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).appendChild(objXML.createElement("cate3"))

									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).appendChild(objXML.createElement("name"))
									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).childNodes(0).childNodes(0).text = rsCd3(1,k)

									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).appendChild(objXML.createElement("cate_key"))
									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).childNodes(1).text = rsCd3(0,k)
									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).appendChild(objXML.createElement("attr"))
									objXML.documentElement.childNodes(4).childNodes(i).childNodes(3).childNodes(j).childNodes(3).childNodes(k).childNodes(2).text = ""
							next

						end if
						rsCd3 = ""

				next

			end if
			rsCd2 = ""
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