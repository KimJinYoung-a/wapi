<%@ CODEPAGE="65001" LANGUAGE="VBSCRIPT" %>
<% option explicit %>
<% session.codepage = "65001" %>
<% response.charset = "utf-8" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/JSON_UTIL_0.1.1.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp"-->
<!-- #include virtual="/lib/classes/item/iteminfoCls.asp"-->
<!-- #include virtual="/lib/classes/item/CategoryPrdCls.asp"-->
<%

'// ===========================================================================
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("110.93.128.107","61.252.133.2","61.252.133.70","61.252.133.10","61.252.133.80","110.93.128.114","110.93.128.113","61.252.133.67","192.168.1.67", "52.79.95.197")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function

dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
    dbDatamart_dbget.Close()
    response.end
end if


dim VISION_API_KEY : VISION_API_KEY = "AIzaSyBvu1RqNG_cM1SJOem4MdEYpnSGGRL5fUU"
dim VISION_API_URL : VISION_API_URL = "https://vision.googleapis.com/v1/images:annotate?key=" & VISION_API_KEY

dim itemid, IMAGE_URL, IMAGE_DATA, IMAGE_DATA_BASE64
dim arrImageURL, arrItemId, arrText, itemidList, imageList, itemData, imageData, jsonData
dim sqlStr, affectedRows
dim i, j, k
dim mode, data, HTTP_MODE

function GetImageDataFromURL(url)
	dim objHTTP
	Set objHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objHTTP.open "GET", url, false
	objHTTP.send
	GetImageDataFromURL = objHTTP.ResponseBody
	Set objHTTP = Nothing
end function

Function Base64Encode(sBin)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue = sBin
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function GetImageText(data)
	dim jsonString, jsonObj, outputObj
	dim objXMLhttp
	dim result

	jsonString = ""
	jsonString = jsonString + "{"
	jsonString = jsonString + "  ""requests"":["
	jsonString = jsonString + "    {"
	jsonString = jsonString + "      ""image"":{"
	jsonString = jsonString + "        ""content"":""" & data & """"
	jsonString = jsonString + "      },"
	jsonString = jsonString + "      ""features"":["
	jsonString = jsonString + "        {"
	jsonString = jsonString + "          ""type"":""TEXT_DETECTION"""
	jsonString = jsonString + "        }"
	jsonString = jsonString + "      ]"
	jsonString = jsonString + "    }"
	jsonString = jsonString + "  ]"
	jsonString = jsonString + "}"

	Set objXMLhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
	objXMLhttp.open "POST",VISION_API_URL, false
	objXMLhttp.setRequestHeader "Content-type","application/json"
	objXMLhttp.setRequestHeader "Accept","application/json"
	objXMLhttp.send jsonString
	jsonString = objXMLhttp.responseText
	Set objXMLhttp = Nothing

	Set jsonObj = New aspJSON
	jsonObj.loadJSon(jsonString)

	result = ""
	On Error Resume Next
		result = jsonObj.data("responses").item(0).item("fullTextAnnotation").item("text")
		if (Err.Number <> 0) then
			result = "ERROR"
		end if
	On Error Goto 0

	GetImageText = result

	Set jsonObj = Nothing
end Function

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

Function getImgTagURL(HTMLstring)
	dim RegEx, URL, Matches, Match
    Set RegEx = New RegExp
    With RegEx
        .Pattern = "src=[\""\']([^\""\']+)"
        .IgnoreCase = True
        .Global = True
        .Multiline = True
    End With

    Set Matches = RegEx.Execute(HTMLstring)
    URL = ""
    For Each Match in Matches
	    if URL = "" then
	    	URL = Match.Value
	    else
	        URL = URL + vbCrLf + Match.Value
	    end if
    Next

    Set Match = Nothing
    Set RegEx = Nothing

    getImgTagURL = Replace(URL, "src=""", "")
End Function

function AddString(str1, str2)
	if (str1 = "") then
		AddString = str2
	else
		AddString = str1 & vbCrLf & str2
	end if
end function

function GetImageUrlList(itemid)
	dim oItem, oADD, i
	set oItem = new CatePrdCls
	oItem.GetItemData itemid

	set oADD = new CatePrdCls
	oADD.getAddImage itemid

	dim arrImageURL : arrImageURL = ""
	dim imageURL

	if Trim(oItem.Prd.FItemContent) <> "" then
		imageURL = getImgTagURL(oItem.Prd.FItemContent)
		arrImageURL = AddString(arrImageURL, imageURL)
	end if

	IF oAdd.FResultCount > 0 THEN
		FOR i= 0 to oAdd.FResultCount-1
			IF oAdd.FADD(i).FAddImageType=1 AND oAdd.FADD(i).FIsExistAddimg THEN
				arrImageURL = AddString(arrImageURL, oAdd.FADD(i).FAddimage)
			end if
		next
	end if

	if ImageExists(oItem.Prd.FImageMain) then
		arrImageURL = AddString(arrImageURL, oItem.Prd.FImageMain)
	end if

	if ImageExists(oItem.Prd.FImageMain2) then
		arrImageURL = AddString(arrImageURL, oItem.Prd.FImageMain2)
	end if

	if ImageExists(oItem.Prd.FImageMain3) then
		arrImageURL = AddString(arrImageURL, oItem.Prd.FImageMain3)
	end if

	GetImageUrlList = arrImageURL
end function

function GetItemIdToExtract(cnt)
	dim sqlStr, affectedRows, i
	dim rows, arrItemID

	sqlStr = " select top " & cnt & " itemid "
	sqlStr = sqlStr + " from [db_contents].[dbo].[tbl_itemImageText] "
	sqlStr = sqlStr + " where 1=1 "
	sqlStr = sqlStr + " and req_yyyymmdd >= convert(varchar(10), DateAdd(d, -20, getdate()), 121) "
	sqlStr = sqlStr + " and fin_yyyymmdd is NULL "
	sqlStr = sqlStr + " order by itemid desc "
	''response.write sqlStr
	''response.end
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr,dbget,adOpenForwardOnly, adLockReadOnly

	if Not rsget.Eof then
		rows = rsget.GetRows
	end if
	rsget.Close

	arrItemID = ""
	if IsArray(rows) then
		for i = 0 to UBound(rows, 2)
			if arrItemID = "" then
				arrItemID = rows(0, i)
			else
				arrItemID = arrItemID & vbCrLf & rows(0, i)
			end if
		next
	end if

	GetItemIdToExtract = arrItemID
end function

function UpdateImageText(itemid, txt)
	dim sqlStr
	dim objCommand

	sqlStr = " update [db_contents].[dbo].[tbl_itemImageText] "
	sqlStr = sqlStr + " set imagetext = ?, updatecnt = updatecnt + 1, fin_yyyymmdd = convert(varchar(10), getdate(), 121) "
	sqlStr = sqlStr + " where itemid = ?"
	''dbget.Execute sqlStr

	Set objCommand = CreateObject("ADODB.Command")
	With objCommand
		.CommandText = sqlStr
		.CommandType = adCmdText
		.ActiveConnection = dbget
		.Parameters.Append .CreateParameter("@p1", adLongVarWChar, adParamInput, 50000, txt)
		.Parameters.Append .CreateParameter("@p2", adInteger, adParamInput, , itemid)
		.Execute
	End With

end function


mode = RequestCheckVar(request("mode"),32)
HTTP_MODE = "GET"
if (request.Form("mode") <> "") then
	HTTP_MODE = "POST"
	mode = Left(request.Form("mode"), 32)
	data = request.Form("data")
end if

select case mode
	case "selData"
		arrItemId = GetItemIdToExtract(10)

		if Trim(arrItemId) <> "" then
			arrItemId = Split(arrItemId, vbCrLf)
			redim itemidList(UBound(arrItemId))

			for i = 0 to UBound(arrItemId)
				itemid = arrItemId(i)
				arrImageURL = GetImageUrlList(itemid)
				arrImageURL = Split(arrImageURL, vbCrLf)

				set itemData = jsObject()
				itemData("itemid") = itemid
				itemData("imageURL") = arrImageURL

				Set itemidList(i) = itemData
			next
		else
			itemidList = ""
		end if

		set jsonData = jsObject()

		if IsArray(itemidList) then
			jsonData("data") = itemidList
			jsonData("count") = UBound(itemidList) + 1
		else
			jsonData("count") = 0
		end if

		Response.Write toJSON(jsonData)
	case "saveData"
		Set jsonData = new aspJSON
		jsonData.loadJSON(data)

		if (jsonData.data("count") > 0) then
			for each itemData in jsonData.data("data")
				set itemData = jsonData.data("data").item(itemData)

				Call UpdateImageText(itemData.item("itemid"), itemData.item("text"))
			next

			Response.Write "{""status"": ""OK"",""HTTP_MODE"": """ & HTTP_MODE & """}"
		else
			Response.Write "{""status"": ""FAIL"",""HTTP_MODE"": """ & HTTP_MODE & """}"
		end if
	case else
		'// error
end select


















'if (itemid <> "") then
'	arrItemId = itemid
'else
'	arrItemId = GetItemIdToExtract(10)
'end if
'
'dim otime,orgTim,diffTime
'''otime = Timer()
'''orgTim = otime
'
'if Trim(arrItemId) <> "" then
'	arrItemId = Split(arrItemId, vbCrLf)
'	for i = 0 to UBound(arrItemId)
'		otime = Timer()
'		itemid = arrItemId(i)
'		arrText = ""
'		response.write "상품코드 : " & itemid & "<br />"
'		arrImageURL = GetImageUrlList(itemid)
'		response.write Replace(arrImageURL, vbCrLf, "<br />") & "<br />"
'		arrImageURL = Split(arrImageURL, vbCrLf)
'		for j = 0 to 0 ''UBound(arrImageURL)
'			IMAGE_URL = arrImageURL(j)
'			IMAGE_DATA = GetImageDataFromURL(IMAGE_URL)
'			response.write "경과시간 : " & FormatNumber(Timer()-otime,4) & "<br /><br />"
'			IMAGE_DATA_BASE64 = Base64encode(IMAGE_DATA)
'			response.write "경과시간 : " & FormatNumber(Timer()-otime,4) & "<br /><br />"
'			''arrText = AddString(arrText, GetImageText(IMAGE_DATA_BASE64))
'			response.write "경과시간 : " & FormatNumber(Timer()-otime,4) & "<br /><br />"
'		next
'		if (Trim(arrText) = "") then
'			arrText = "NO TEXT"
'		end if
'		''Call UpdateImageText(itemid, arrText)
'		response.write "경과시간 : " & FormatNumber(Timer()-otime,4) & "<br /><br />"
'	next
'end if





''IMAGE_DATA = Base64encode(GetImageDataFromURL(IMAGE_URL))
''response.write GetImageText(IMAGE_DATA)

''itemid = 1902829

''arrImageURL = GetImageUrlList(itemid)



%>
<!-- #include virtual="/lib/db/dbclose.asp" -->