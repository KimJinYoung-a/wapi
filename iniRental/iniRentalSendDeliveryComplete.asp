<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% session.CodePage = "65001" %>
<% Server.ScriptTimeOut = 1200 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/aspJSON1.17.asp" -->
<%
	Dim query1, retCount, iniMid, iniConfirmUrl, strData, xmlHttp, oJSON, vquery
    Dim resultCode, postdata
    Dim resultMsg, resultDate, resultTime, tid

    IF (application("Svr_Info")	= "Dev") Then
        iniMid = "teenxtest1"
        iniConfirmUrl = "https://inirt.inicis.com/apis/v1/rental/confirm"
    Else
        iniMid="teenxteenr"
        iniConfirmUrl = "https://inirt.inicis.com/apis/v1/rental/confirm"
    End If

    query1 = "" & vbcrlf
    query1 = query1 & " SELECT m.orderserial, m.userid, m.subtotalprice, m.regdate, m.paygatetid " & vbcrlf
    query1 = query1 & "     , d.songjangno, d.songjangdiv, d.dlvfinishdt, r.idx, r.resultcode, r.resultmsg, r.tid, r.regdate " & vbcrlf
    query1 = query1 & " FROM db_order.dbo.tbl_order_master m WITH(NOLOCK) " & vbcrlf
    query1 = query1 & " INNER JOIN db_order.dbo.tbl_order_detail d WITH(NOLOCK) " & vbcrlf
    query1 = query1 & "     ON m.orderserial = d.orderserial " & vbcrlf
    query1 = query1 & " LEFT JOIN db_order.dbo.tbl_iniRentalDeliveryCompleteSendData r WITH(NOLOCK) " & vbcrlf
    query1 = query1 & "     ON m.orderserial = r.orderserial " & vbcrlf
    query1 = query1 & " WHERE d.itemid NOT IN (0,100) " & vbcrlf
    query1 = query1 & "     AND m.jumundiv = 8 " & vbcrlf
    query1 = query1 & "     AND m.accountdiv = '150' " & vbcrlf
    query1 = query1 & "     AND m.ipkumdiv = 8 " & vbcrlf
    query1 = query1 & "     AND m.cancelyn = 'N' " & vbcrlf
    query1 = query1 & "     AND m.sitename = '10x10' " & vbcrlf
    query1 = query1 & "     AND ISNULL(d.dlvfinishdt,'') <> '' " & vbcrlf
    query1 = query1 & "     AND ISNULL(d.songjangno,'') <> '' " & vbcrlf
    query1 = query1 & "     AND ISNULL(r.resultcode,'') NOT IN ('00','ERR217') "
	rsget.CursorLocation = adUseClient
	rsget.Open query1,dbget, adOpenForwardOnly, adLockReadOnly

    retCount = rsget.recordcount

	If Not(rsget.bof Or rsget.eof) Then
        Do Until rsget.eof

            postdata = "mid="&CStr(iniMid)
            postdata = postdata&"&type=Confirm"
            postdata = postdata&"&clientIp="&CStr(request.ServerVariables("LOCAL_ADDR"))
            postdata = postdata&"&timestamp="&DateDiff("s", "1970-01-01 09:00:00", now)*1000+clng(timer)
            postdata = postdata&"&tid="&Cstr(rsget("paygatetid"))
            If Trim(rsget("songjangno")) = "삼성로지텍직배송" or Trim(rsget("songjangno")) = "삼성전자물류직배송" or Trim(rsget("songjangno")) = "삼성전자물류배송" Then
                postdata = postdata&"&invoiceNum="&CStr("samsunglogitech")
            Else
                postdata = postdata&"&invoiceNum="&CStr(Trim(rsget("songjangno")))
            End If
            postdata = postdata&"&deliveryConfirmDt="&CStr(replace(Left(Trim(rsget("dlvfinishdt")),10),"-",""))
            postdata = postdata&"&courierCode="&CStr(deliveryCompanyCodeMatchInicis(Trim(rsget("songjangdiv"))))
            Set xmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
            xmlHttp.open "POST",iniConfirmUrl, False
            xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=utf-8"  ''UTF-8 charset 필요.
            xmlHttp.setTimeouts 90000,90000,90000,90000 ''2013/03/14 추가
            xmlHttp.Send postdata	'post data send
            strData = BinaryToText(xmlHttp.responseBody, "UTF-8")
            Set xmlHttp = nothing

            Set oJSON = New aspJSON
            oJSON.loadJSON(strData)
            resultCode = oJSON.data("resultCode")
            resultMsg = oJSON.data("resultMsg")
            resultDate = oJSON.data("resultDate")
            resultTime = oJSON.data("resultTime")
            tid = oJSON.data("tid")
            Set oJSON = Nothing

            If rsget("idx") <> "" Then
                vquery = " UPDATE db_order.dbo.tbl_iniRentalDeliveryCompleteSendData "
                vquery = vquery & " SET resultcode='"&resultCode&"' "
                vquery = vquery & " , resultmsg='"&resultMsg&"' "
                vquery = vquery & " , resultdate='"&resultDate&"' "
                vquery = vquery & " , resulttime='"&resultTime&"' "
                vquery = vquery & " , tid='"&tid&"' "
                vquery = vquery & " , lastupdate = getdate() "
                vquery = vquery & " WHERE idx="&rsget("idx")
                dbget.Execute vquery
            Else
                vquery = " INSERT INTO db_order.dbo.tbl_iniRentalDeliveryCompleteSendData "
                vquery = vquery &" (orderserial, userid, resultcode, resultmsg, resultdate, resulttime, tid, regdate, lastupdate) VALUES "
                vquery = vquery &" ('"&rsget("orderserial")&"', '"&rsget("userid")&"', '"&resultCode&"', '"&resultMsg&"', '"&resultDate&"', '"&resultTime&"' "
                vquery = vquery &" , '"&tid&"', getdate(), getdate())"
                dbget.Execute vquery
            End If
        rsget.movenext
        Loop

    End If
	rsget.close

    response.write "ok"
    response.end

	Function deliveryCompanyCodeMatchInicis(dcode)
        Select Case dcode
            case "1"
                deliveryCompanyCodeMatchInicis = "hanjin"
            case "2"
                deliveryCompanyCodeMatchInicis = "hyundai"
            case "3"
                deliveryCompanyCodeMatchInicis = "korex"
            case "4"
                deliveryCompanyCodeMatchInicis = "cjgls"
            case "8"
                deliveryCompanyCodeMatchInicis = "EPOST"
            case "9"
                deliveryCompanyCodeMatchInicis = "kgbps"
            case "18"
                deliveryCompanyCodeMatchInicis = "kgb"
            case "21"
                deliveryCompanyCodeMatchInicis = "kdexp"
            case "26"
                deliveryCompanyCodeMatchInicis = "ilyang"
            case "31"
                deliveryCompanyCodeMatchInicis = "chunil"
            case "33"
                deliveryCompanyCodeMatchInicis = "honam"
            case "34"
                deliveryCompanyCodeMatchInicis = "daesin"
            case "37"
                deliveryCompanyCodeMatchInicis = "hdexp"
            case "42"
                deliveryCompanyCodeMatchInicis = "cvsnet"
            case else
                deliveryCompanyCodeMatchInicis = "9999"
        End Select
	End Function

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

    Function StringChk(Str)
        dim i, strTemp
        For i = 1 To Len(Str) '문자열을 하나씩 짤라서 체크 
            strTemp = Asc(Mid(Str, i, 1)) 

            '음수(-) 가 나오면 한글 
            If strTemp < 0 Then Exit For ' 한글이 존재 할경우 종료 
        Next 
        StringChk = strTemp 
    End Function

    
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->