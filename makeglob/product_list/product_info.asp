<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #INCLUDE Virtual="/lib/util/commlib.asp" -->
<!-- #include virtual="/lib/classes/item/iteminfoCls.asp" -->
<!-- #include virtual="/lib/classes/item/CategoryPrdCls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim sqlStr, objXML, i, j
	Dim rsCd1, rsCd2, vProductKey, webImgUrl, tmpcateindex
	Dim arroptionidx, optionname, arroptionvalue, arroptionprice, arroptionstock, arroptionsoldout, arroptionhidden
	Dim spiltLen, tmparridx, tmparrvalue, tmparrprice, tmparrstock, tmparrsoldout, tmparrhidden
	Dim tmpBrand, tb, brandLinkVal


	'// 상품영역은 기본적으로 상품코드를 받아온다. 없으면 아무것도 표시 안함

	vProductKey = request("code")

	If vProductKey="" Or isnull(vProductKey) Then
		Response.write "잘못된 접근"
    	dbget.close()	:	response.end
	End If

    if Not(isNumeric(vProductKey)) then
    	response.write "잘못된 접근(2)"
    	dbget.close()	:	response.end
    end If

	IF application("Svr_Info")="Dev" Then
	 	webImgUrl		= "http://testwebimage.10x10.co.kr"			'웹이미지
	Else
 		webImgUrl		= "http://webimage.10x10.co.kr"				'웹이미지
	End If


'	파일 저장시(상품은 각 상품별로 저장해야 하므로 하단 코드가 루프문 안에 들어감)
'	Dim savePath, FileName
'	savePath = server.mappath("/makeglob/") + "\"
'	FileName = "상품코드.xml"

	Dim product_key '// 메이크글로비용 상품코드
	Dim product_code '// 텐바이텐 상품코드
	Dim product_language '// 언어코드(일단 무조건 한국어만 넘어감KO)
	Dim vcurrency '// 통화코드(한국어만 넘어가므로 KRW)
	Dim product_name '// 상품명
	Dim product_price '// 상품가격
	Dim original_price '// 상품원래가격(보통은 판매가격 원래가격이 틀리지만 메이크글로비용은 둘다 같은 가격으로 넘김 가격제어는 자체 admin에서 처리)
	Dim supply_price '// 매입가격(요건 글로비쪽에 오픈할 이유가 없으므로 0원임)
	Dim list_img_url '// 리스트용 이미지(list120)
	Dim detail_img_url '// 요건 상품상세 이미지(icon1image)
	Dim zoom_img_url '// 요건 좀더 확대된 basic 이미지
	Dim basic600_img_url '// 600이미지
	Dim basic1000_img_url '// 1000이미지
	Dim mileage '// 마일리지는 무조건 0원으로(컨트롤은 자체 admin에서 처리)
	Dim weight '// 상품무게
	Dim maker_name '// 메이커명
	Dim madein '// 원산지
	Dim brand_name '// 브랜드명
	Dim manufacture_date '// 제작일
	Dim launching_date '// 판매시작일
	Dim keyword '// 검색키워드
	Dim product_desc '// 상품상세(여기선 현재 있는 상품 상세에서 이미지와 텍스트만 추출하여 따로 저장함)
	Dim itemsource '// 상품재료
	Dim itemsize '// 상품크기
	Dim hidden '// 숨김여부
	Dim soldout '// 품절여부
	Dim product_url '// 상품 다이렉트 링크(이건 좀 봐야되는걸로..)
	Dim cateindex '// 카테고리 번호 split로 잘라서 row로 뿌려줌
	Dim makeglobYN '// 메이크 글로비에 넘겼는지 여부(나중에 해당 상품내용 업데이트 하고 싶으면 다시 N처리하면 알아서 긁어감)
	Dim regdate '// 상품등록일
	Dim lastupdate '// 최종수정일
	Dim makeupdate '// 메이크글로비에 올라간 일자
	Dim mwdiv

'	sqlStr = " Select "
'	sqlStr = sqlStr & "	product_key, product_code, product_language, currency, product_name, product_price, original_price, supply_price, "
'	sqlStr = sqlStr & "	list_img_url, detail_img_url, zoom_img_url, basic600_img_url, basic1000_img_url, mileage, weight, maker_name, madein, brand_name, "
'	sqlStr = sqlStr & "	manufacture_date, launching_date, keyword, [desc] as product_desc, itemsource, itemsize, hidden, soldout, product_url, "
'	sqlStr = sqlStr & "	cateindex, makeglobYN, convert(varchar(19), regdate, 120) as regdate, convert(varchar(19), lastupdate, 120) as lastupdate, "
'	sqlStr = sqlStr & "	convert(varchar(19), makeupdate, 120) as makeupdate "
'	sqlStr = sqlStr & "	From db_item.[dbo].[tbl_makeglob_product] "
'	sqlStr = sqlStr & " Where product_key='"&vProductKey&"' "

	sqlStr = " Select "
	sqlStr = sqlStr & "	p.product_key, p.product_code, p.product_language, p.currency, p.product_name, p.product_price, p.original_price, p.supply_price, "
	sqlStr = sqlStr & "	p.list_img_url, p.detail_img_url, p.zoom_img_url, p.basic600_img_url, p.basic1000_img_url, p.mileage, p.weight, p.maker_name, p.madein, p.brand_name, "
	sqlStr = sqlStr & "	p.manufacture_date, p.launching_date, p.keyword, [desc] as product_desc, p.itemsource, p.itemsize, p.hidden, p.soldout, p.product_url, "
	sqlStr = sqlStr & "	p.cateindex, p.makeglobYN, convert(varchar(19), p.regdate, 120) as regdate, convert(varchar(19), p.lastupdate, 120) as lastupdate, "
	sqlStr = sqlStr & "	convert(varchar(19), p.makeupdate, 120) as makeupdate, i.mwdiv "
	sqlStr = sqlStr & "	From db_item.[dbo].[tbl_makeglob_product] as p "
	sqlStr = sqlStr & "	JOIN db_item.dbo.tbl_item as i on p.product_code = i.itemid "
	sqlStr = sqlStr & " Where p.product_key='"&vProductKey&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) Then
		product_key = rsget("product_key")
		product_code = rsget("product_code")
		product_language = rsget("product_language")
		vcurrency = rsget("currency")
		product_name = rsget("product_name")
		product_price = rsget("product_price")
		original_price = rsget("original_price")
		supply_price = rsget("supply_price")
		list_img_url = rsget("list_img_url")
		detail_img_url = rsget("detail_img_url")
		basic600_img_url = rsget("basic600_img_url")
		basic1000_img_url = rsget("basic1000_img_url")
		zoom_img_url = rsget("zoom_img_url")
		mileage = rsget("mileage")
		weight = rsget("weight")
		maker_name = rsget("maker_name")
		madein = rsget("madein")
		brand_name = rsget("brand_name")
		manufacture_date = rsget("manufacture_date")
		launching_date = rsget("launching_date")
		keyword = rsget("keyword")
		product_desc = rsget("product_desc")
		itemsource = rsget("itemsource")
		itemsize = rsget("itemsize")
		hidden = rsget("hidden")
		soldout = rsget("soldout")
		product_url = rsget("product_url")
		cateindex = rsget("cateindex")
		makeglobYN = rsget("makeglobYN")
		regdate = rsget("regdate")
		lastupdate = rsget("lastupdate")
		makeupdate = rsget("makeupdate")
		mwdiv = rsget("mwdiv")
	Else
		Response.Write "상품 내용 없음."
		Response.End
	END IF
	rsget.close


	'/// 상품 상세 내용 접수
	dim oItem, ItemContent

	set oItem = new CatePrdCls
	oItem.GetItemData product_code

	'// 추가 이미지
	dim oADD
	dim itemContImg, tmpExtImgArr
	'// 결과 JSON 생성
	Dim arrRst, strItemDesc, tmpDesc

	set oADD = new CatePrdCls
	oADD.getAddImage product_code

	'## 설명이미지 표시 (1순위: HTML내 이미지, 2순위: 추가이미지, 3순위: 업로드된 이미지)
	'HTML내 이미지
	tmpExtImgArr = RegExpArray("[^=']*\.(gif|jpg|bmp|png|GIF|JPG|BMP|PNG)", oItem.Prd.FItemContent)
	if isArray(tmpExtImgArr) then
		for i=0 to ubound(tmpExtImgArr)
			itemContImg = itemContImg & "<p><img src='"&chkIIF(itemContImg<>"",vbCrLf,"") & Replace(tmpExtImgArr(i),"""","")&"'></p>"
		next
	end If


	'추가이미지
	IF oAdd.FResultCount > 0 THEN
		FOR i= 0 to oAdd.FResultCount-1
			IF oAdd.FADD(i).FAddImageType=1 THEN
				itemContImg = itemContImg & "<p><img src='"&chkIIF(itemContImg<>"",vbCrLf,"") & oAdd.FADD(i).FAddimage&"'></p>"
			End IF
		NEXT
	end If

	'상품 설명 업로드 이미지
	if ImageExists(oItem.Prd.FImageMain) then
		itemContImg = itemContImg & "<p><img src='"&chkIIF(itemContImg<>"",vbCrLf,"") & oItem.Prd.FImageMain&"'></p>"
	end if
	if ImageExists(oItem.Prd.FImageMain2) then
		itemContImg = itemContImg & "<p><img src='"&chkIIF(itemContImg<>"",vbCrLf,"") & oItem.Prd.FImageMain2&"'></p>"
	end if
	if ImageExists(oItem.Prd.FImageMain3) then
		itemContImg = itemContImg & "<p><img src='"&chkIIF(itemContImg<>"",vbCrLf,"") & oItem.Prd.FImageMain3&"'></p>"
	end If



	'상품설명 Text 추출
	strItemDesc = trim(oItem.Prd.FItemContent)
	IF oItem.Prd.FUsingHTML="Y" THEN strItemDesc = replace(strItemDesc,vbCrLf,"")		'HTML사용일때 엔터 제거
	strItemDesc = replace(strItemDesc,vbTab,"")
	strItemDesc = replace(strItemDesc,"<br>",vbCrLf)
	strItemDesc = replace(strItemDesc,"<br/>",vbCrLf)
	strItemDesc = replace(strItemDesc,"<br />",vbCrLf)
	strItemDesc = replace(strItemDesc,"<BR>",vbCrLf)
	strItemDesc = replace(strItemDesc,"<BR/>",vbCrLf)
	strItemDesc = replace(strItemDesc,"<BR />",vbCrLf)
	strItemDesc = stripHTML(strItemDesc)

	'내용 빈칸 삭제
	tmpDesc = split(strItemDesc,vbCrLf)
	strItemDesc = ""

	for i=0 to ubound(tmpDesc)
		if trim(tmpDesc(i))<>"" then
			strItemDesc = chkIIF(strItemDesc<>"",vbCrLf,"") & tmpDesc(i)
		end if
	Next


	'//사용자 함수
	function ImageExists(byval iimg)
		if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
			ImageExists = false
		else
			ImageExists = true
		end if
	end function

	'//객체 정리
	set oItem = Nothing
	set oADD = Nothing

	'// 상품설명
	product_desc = ""
	If mwdiv <> "M" AND mwdiv <> "W" then
		product_desc = product_desc & "<p><img src='http://webimage.10x10.co.kr/common/uploadimg/2017/china/1716265/1705301355410340028.jpg'></p>"
	End If
	product_desc = product_desc & itemContImg
	product_desc = product_desc & "<p>"&strItemDesc&"</p>"
	product_desc = product_desc & "<div class='pdtInforBox tMar05'> "
	product_desc = product_desc & "<div class='pdtInforList'><i>원산지</i>&nbsp;: "&madein&"&nbsp;&nbsp;<i>제조사</i>&nbsp;: "&maker_name&"&nbsp;<i>재질</i>: "&itemsource&" &nbsp;<i>사이즈</i>&nbsp;: "&itemsize&"&nbsp;</div> "
	product_desc = product_desc & "</div> "



	'// 브랜드 링크생성
	If cateindex <> "" Then
		tmpBrand = Split(cateindex, ",")
		For tb=0 To ubound(tmpBrand)
			brandLinkVal = "/Product/Category/list/cid/"&tmpBrand(tb)
		Next
	End If


	'// 긁어갈때 마다 메이크 글로비 업로드 여부와 업데이트 일자 update 시켜준다.
	'// 해당 아이피로 접근할 시에만 업데이트 함.
	If Left(Request.ServerVariables("REMOTE_ADDR"), 9) = "14.129.44" Or Left(Request.ServerVariables("REMOTE_ADDR"), 9) = "14.129.31" Or Left(Request.ServerVariables("REMOTE_ADDR"), 9) = "14.129.48" Or Left(Request.ServerVariables("REMOTE_ADDR"), 9) = "121.78.48" Then
		sqlstr = " update db_item.dbo.tbl_makeglob_product set makeglobYN='Y', makeupdate = getdate() Where product_key = '"&product_key&"' And product_code='"&product_code&"' "
		dbget.execute sqlstr
	End If

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
	objXML.documentElement.childNodes(2).text = lastupdate

	objXML.documentElement.appendChild(objXML.createElement("product_language"))
	objXML.documentElement.childNodes(3).text = product_language

	objXML.documentElement.appendChild(objXML.createElement("product_key"))
	objXML.documentElement.childNodes(4).text = product_key

	objXML.documentElement.appendChild(objXML.createElement("currency"))
	objXML.documentElement.childNodes(5).text = vcurrency

	objXML.documentElement.appendChild(objXML.createElement("product_name"))
	objXML.documentElement.childNodes(6).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(6).childNodes(0).text = product_name

	objXML.documentElement.appendChild(objXML.createElement("product_price"))
	objXML.documentElement.childNodes(7).text = product_price

	objXML.documentElement.appendChild(objXML.createElement("original_price"))
	objXML.documentElement.childNodes(8).text = original_price

	objXML.documentElement.appendChild(objXML.createElement("supply_price"))
	objXML.documentElement.childNodes(9).text = supply_price

	objXML.documentElement.appendChild(objXML.createElement("list_img_url"))
	objXML.documentElement.childNodes(10).appendChild(objXML.createCDATASection("name_Cdata"))
	If isnull(basic600_img_url) Or Trim(basic600_img_url)="" Then
		objXML.documentElement.childNodes(10).childNodes(0).text = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(product_code)&"/"&zoom_img_url
	Else
		objXML.documentElement.childNodes(10).childNodes(0).text = webImgUrl&"/image/basic600/"&GetImageSubFolderByItemid(product_code)&"/"&basic600_img_url
	End If

	objXML.documentElement.appendChild(objXML.createElement("detail_img_url"))
	objXML.documentElement.childNodes(11).appendChild(objXML.createCDATASection("name_Cdata"))
	If isnull(basic600_img_url) Or Trim(basic600_img_url)="" Then
		objXML.documentElement.childNodes(11).childNodes(0).text = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(product_code)&"/"&zoom_img_url
	Else
		objXML.documentElement.childNodes(11).childNodes(0).text = webImgUrl&"/image/basic600/"&GetImageSubFolderByItemid(product_code)&"/"&basic600_img_url
	End If

	objXML.documentElement.appendChild(objXML.createElement("zoom_img_url"))
	objXML.documentElement.childNodes(12).appendChild(objXML.createCDATASection("name_Cdata"))
	If isnull(basic600_img_url) Or Trim(basic600_img_url)="" Then
		objXML.documentElement.childNodes(12).childNodes(0).text = webImgUrl&"/image/basic/"&GetImageSubFolderByItemid(product_code)&"/"&zoom_img_url
	Else
		objXML.documentElement.childNodes(12).childNodes(0).text = webImgUrl&"/image/basic600/"&GetImageSubFolderByItemid(product_code)&"/"&basic600_img_url
	End If

	objXML.documentElement.appendChild(objXML.createElement("mileage"))
	objXML.documentElement.childNodes(13).text = mileage

	objXML.documentElement.appendChild(objXML.createElement("weight"))
	objXML.documentElement.childNodes(14).text = weight

	objXML.documentElement.appendChild(objXML.createElement("maker_name"))
	objXML.documentElement.childNodes(15).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(15).childNodes(0).text = maker_name

	objXML.documentElement.appendChild(objXML.createElement("madein"))
	objXML.documentElement.childNodes(16).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(16).childNodes(0).text = madein

	objXML.documentElement.appendChild(objXML.createElement("brand_name"))
	objXML.documentElement.childNodes(17).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(17).childNodes(0).text = brand_name

	objXML.documentElement.appendChild(objXML.createElement("manufacture_date"))
	If manufacture_date = "" Or isnull(manufacture_date) Then

	Else
		objXML.documentElement.childNodes(18).text = Left(manufacture_date, 10)
	End if

	objXML.documentElement.appendChild(objXML.createElement("launching_date"))
	If launching_date = "" Or isnull(launching_date) Then

	Else
		objXML.documentElement.childNodes(19).text = Left(launching_date, 10)
	End if



	objXML.documentElement.appendChild(objXML.createElement("keyword"))
	objXML.documentElement.childNodes(20).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(20).childNodes(0).text = keyword

	objXML.documentElement.appendChild(objXML.createElement("product_code"))
	objXML.documentElement.childNodes(21).text = product_code

	objXML.documentElement.appendChild(objXML.createElement("seller_product_code"))
	objXML.documentElement.childNodes(22).text = product_code

	objXML.documentElement.appendChild(objXML.createElement("desc"))
	objXML.documentElement.childNodes(23).appendChild(objXML.createCDATASection("name_Cdata"))
	objXML.documentElement.childNodes(23).childNodes(0).text = product_desc

	objXML.documentElement.appendChild(objXML.createElement("hidden"))
	objXML.documentElement.childNodes(24).text = hidden

	objXML.documentElement.appendChild(objXML.createElement("soldout"))
	objXML.documentElement.childNodes(25).text = soldout

	objXML.documentElement.appendChild(objXML.createElement("product_url"))
	objXML.documentElement.childNodes(26).text = product_url


	tmpcateindex = Split(cateindex, ",")
	objXML.documentElement.appendChild(objXML.createElement("category_index"))
	For i=0 To ubound(tmpcateindex)
		objXML.documentElement.childNodes(27).appendChild(objXML.createElement("category"))
		objXML.documentElement.childNodes(27).childNodes(i).Text = tmpcateindex(i)
	Next


	'// 옵션 데이터 가져옴.
	arroptionidx = "" '// 옵션 고유번호
	optionname = "" '// 옵션명
	arroptionvalue = "" '// 각각의 옵션명
	arroptionprice = "" '// 각각의 추가금액
	arroptionstock = "" '// 각각의 재고갯수
	arroptionsoldout = "" '// 각각의 품절여부
	arroptionhidden = "" '// 각각의 숨김여부

	sqlStr = " Select idx, product_key, product_code, option_index_name, option_index_value, option_index_price, stock, soldout, hidden "
	sqlStr = sqlStr & "	From db_item.[dbo].[tbl_makeglob_product_option] Where product_key='"&product_key&"' And product_code='"&product_code&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	IF Not (rsget.EOF OR rsget.BOF) Then

		Do Until rsget.EOF
			arroptionidx = arroptionidx &"^"&rsget("idx")
			optionname = replace(rsget("option_index_name"), Chr(32), "")
			arroptionvalue = arroptionvalue &"^"&Replace(replace(rsget("option_index_value"), Chr(32), ""), Chr(44),"")
			arroptionprice = arroptionprice &"^"&rsget("option_index_price")
			arroptionstock = arroptionstock &"^"&rsget("stock")
			arroptionsoldout = arroptionsoldout &"^"&rsget("soldout")
			arroptionhidden = arroptionhidden &"^"&rsget("hidden")
		rsget.movenext
		Loop


		spiltLen = Len(Trim(arroptionidx))
		arroptionidx = Right(arroptionidx, spiltLen-1)

		spiltLen = Len(Trim(arroptionvalue))
		arroptionvalue = Right(arroptionvalue, spiltLen-1)

		spiltLen = Len(Trim(arroptionprice))
		arroptionprice = Right(arroptionprice, spiltLen-1)

		spiltLen = Len(Trim(arroptionstock))
		arroptionstock = Right(arroptionstock, spiltLen-1)

		spiltLen = Len(Trim(arroptionsoldout))
		arroptionsoldout = Right(arroptionsoldout, spiltLen-1)

		spiltLen = Len(Trim(arroptionhidden))
		arroptionhidden = Right(arroptionhidden, spiltLen-1)

		objXML.documentElement.appendChild(objXML.createElement("option_index"))
		objXML.documentElement.childNodes(28).appendChild(objXML.createElement("option"))

		objXML.documentElement.childNodes(28).childNodes(0).appendChild(objXML.createElement("name"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(0).childNodes(0).text = optionname

		objXML.documentElement.childNodes(28).childNodes(0).appendChild(objXML.createElement("value"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(1).appendChild(objXML.createCDATASection("name_Cdata"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(1).childNodes(0).text = arroptionvalue

		objXML.documentElement.childNodes(28).childNodes(0).appendChild(objXML.createElement("price"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(2).appendChild(objXML.createCDATASection("name_Cdata"))
		objXML.documentElement.childNodes(28).childNodes(0).childNodes(2).childNodes(0).text = arroptionprice


		objXML.documentElement.appendChild(objXML.createElement("option_list_index"))

		tmparridx = Split(arroptionidx,"^")
		tmparrvalue = Split(arroptionvalue,"^")
		tmparrprice = Split(arroptionprice,"^")
		tmparrstock = Split(arroptionstock,"^")
		tmparrsoldout = Split(arroptionsoldout,"^")
		tmparrhidden = Split(arroptionhidden,"^")

		For i=0 To ubound(tmparridx)
			objXML.documentElement.childNodes(29).appendChild(objXML.createElement("option_list"))

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("name"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(0).appendChild(objXML.createCDATASection("name_Cdata"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(0).childNodes(0).text = tmparrvalue(i)

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("key"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(1).text = tmparridx(i)

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("price"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(2).text = tmparrprice(i)

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("stock"))
			If tmparrsoldout(i)="Y" Then
				objXML.documentElement.childNodes(29).childNodes(i).childNodes(3).text = "0"
			Else
				If tmparrstock(i) > 0 Then
					objXML.documentElement.childNodes(29).childNodes(i).childNodes(3).text = tmparrstock(i)
				Else
					objXML.documentElement.childNodes(29).childNodes(i).childNodes(3).text = ""
				End If
			End If

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("soldout"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(4).text = tmparrsoldout(i)

			objXML.documentElement.childNodes(29).childNodes(i).appendChild(objXML.createElement("hidden"))
			objXML.documentElement.childNodes(29).childNodes(i).childNodes(5).text = tmparrhidden(i)
		Next

	Else

		objXML.documentElement.appendChild(objXML.createElement("pdt_stock"))
		If Trim(soldout)="Y" Then
			objXML.documentElement.childNodes(28).text = "0"
		Else
			objXML.documentElement.childNodes(28).text = ""
		End If

	End If

	rsget.close


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