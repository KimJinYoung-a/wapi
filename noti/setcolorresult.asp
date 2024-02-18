<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Response.CharSet="UTF-8"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%
function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("192.168.1.70","192.168.1.72","110.93.128.107","61.252.133.2","61.252.133.69","61.252.133.70","61.252.133.80","61.252.143.71","61.252.133.12","110.93.128.114","110.93.128.113","61.252.133.72")
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
    'response.write ref
    'response.end
end if


'//바이너리 데이터 TEXT형태로 변환
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'원본 데이터 타입
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' 변환할 데이터 캐릭터셋
	 BinaryStream.CharSet = CharSet

	'변환한 데이터 반환
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function


dim queidx, oneScore, ioneScore, i, oneColor, ioneColor
Dim colorName, colorScore, Colors_Score, Colors_PixelFraction, Colors_Color
dim rcvData , lngBytesCount
dim oJSON, sqlStr
i=0
If (Request.TotalBytes > 0) Then
    lngBytesCount = Request.TotalBytes
    rcvData = BinaryToText(Request.BinaryRead(lngBytesCount),"utf-8")
    
    Set oJSON = New aspJSON
    oJSON.loadJSON(rcvData)
    queidx = oJSON.data("queidx")
    'Scores = oJSON.data("Scores")
    
    for each oneScore in oJSON.data("Scores")
        set ioneScore = oJSON.data("Scores").item(oneScore)
	    colorName = ioneScore.item("Name")
        colorScore = ioneScore.item("Score")

        set ioneScore = Nothing

        sqlStr = "db_etcmall.[dbo].[usp_Ten_ColorImage_Que_Result_SET] "&queidx&","&i+1&",'"&colorName&"',"&ColorScore&""
        dbget.Execute sqlStr
        i=i+1
    next

    i=0
    for each oneColor in oJSON.data("Colors")
        set ioneColor = oJSON.data("Colors").item(oneColor)
	    
        Colors_Score = ioneColor.item("Score")
        Colors_PixelFraction = ioneColor.item("PixelFraction")
        Colors_Color = ioneColor.item("Color")
  

        set ioneColor = Nothing

        sqlStr = "db_etcmall.[dbo].[usp_Ten_ColorImage_Que_Result_SET_ROW] "&queidx&","&i+1&",'"&Colors_Score&"',"&Colors_PixelFraction&",'"&Colors_Color&"'"
        dbget.Execute sqlStr
        i=i+1
     next

    Set oJSON = Nothing
    
    response.write "OK"
else
    response.write "TTT"
End If



%>
<!-- #include virtual="/lib/db/dbclose.asp" -->