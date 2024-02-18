<%

function FnGetRecvCheckURL(songjangDiv, songjangNo)
    Dim sqlStr
	dim targetURL

	sqlStr = " select top 1 findurl from db_order.[dbo].tbl_songjang_div where divcd = '" + CStr(songjangDiv) + "' "

	targetURL = ""
    rsget.Open sqlStr, dbget
	If Not(rsget.EOF or rsget.BOF) Then
		targetURL = db2html(rsget("findurl"))
	End If
	rsget.close

	FnGetRecvCheckURL = ""
	if (targetURL <> "") then
		FnGetRecvCheckURL = targetURL + CStr(songjangNo)

		if (CStr(songjangDiv) = "8") then
			'// 우체국
			FnGetRecvCheckURL = "http://trace.epost.go.kr/xtts/servlet/kpl.tts.common.svl.SttSVL?target_command=kpl.tts.tt.epost.cmd.RetrieveOrderConvEpostPoCMD&sid1=" + CStr(songjangNo)
		elseif (CStr(songjangDiv) = "18") then
			'// 로젠
			FnGetRecvCheckURL = "https://www.ilogen.com/iLOGEN.Web.New/TRACE/TraceDetail.aspx?gubun=type2&slipno=" + CStr(songjangNo) + "&invoiceNum=" + CStr(songjangNo)
		elseif (CStr(songjangDiv) = "13") then
			'// 엘로우
			FnGetRecvCheckURL = "https://www.kgyellowcap.co.kr/iframe-delivery.html?delivery=" + CStr(songjangNo)
		else
			'//
		end if
	end if

end function

function FnGetRecvCheckHTML(checkURL)
	dim xmlHTTP, resultHTTP

	set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	xmlHTTP.open "GET", checkURL, false
	xmlHTTP.send ""
	FnGetRecvCheckHTML = xmlHTTP.responseText
	set xmlHTTP = nothing

end function

'// 한글이 깨지는 경우 사용
function FnGetRecvCheckHTML_EUCKR(checkURL)
	dim xmlHTTP, resultHTTP
	dim responseStrm

	set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	xmlHTTP.open "GET", checkURL, false
	xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlHTTP.send ""

	Set responseStrm = CreateObject("ADODB.Stream")

	responseStrm.Open
	responseStrm.Position = 0
	responseStrm.Type = 1
	responseStrm.Write xmlHTTP.responseBody
	responseStrm.Position = 0
	responseStrm.Type = 2
	responseStrm.Charset = "euc-kr"
	FnGetRecvCheckHTML_EUCKR = responseStrm.ReadText
	responseStrm.close
	Set responseStrm = Nothing

	set xmlHTTP = nothing

end function

function FnCheckRecvState(songjangDiv, targetHTML)
	dim arrSongjangDiv(9), arrSearchText(9)
	dim i

	arrSongjangDiv(0) = "1"
	arrSearchText(0) = "배송완료"

	arrSongjangDiv(1) = "2"
	''arrSearchText(1) = "배달 완료하였습니다"
	arrSearchText(1) = "물품을 받으셨습니다"

	arrSongjangDiv(2) = "8"
	arrSearchText(2) = "배달완료"

	arrSongjangDiv(3) = "9"
	arrSearchText(3) = "배송완료. 이용해주셔서 감사합니다"

	arrSongjangDiv(4) = "13"
	arrSearchText(4) = "물품을 받으셨습니다"

	arrSongjangDiv(5) = "18"
	arrSearchText(5) = "배송완료"

	arrSongjangDiv(6) = "21"
	arrSearchText(6) = "서명처리가 완료되었습니다"

	arrSongjangDiv(7) = "28"
	arrSearchText(7) = "배송완료"

	arrSongjangDiv(8) = "31"
	arrSearchText(8) = "배송완료"

	FnCheckRecvState = "X"
	for i = 0 to UBound(arrSongjangDiv)
		if (CStr(songjangDiv) = CStr(arrSongjangDiv(i))) then
			if (InStr(targetHTML, arrSearchText(i)) > 0) then
				'// 있다.
				FnCheckRecvState = "T"
				exit function
			end if
			'
		end if
		'
	next

	FnCheckRecvState = "F"

end function

function FnCheckNSaveRecvState(songjangDiv, songjangNo, byRef errCode, byRef errMSG)
    Dim sqlStr
	dim targetURL, targetHTML, parseResult
	dim IsDebugMode

	'// 디버깅이 필요할 때 True 전환
	IsDebugMode = False
	''IsDebugMode = True

	FnCheckNSaveRecvState = False
	errCode = "0"
	errMSG = ""

	targetURL = FnGetRecvCheckURL(songjangDiv, songjangNo)
	if (IsDebugMode) then
		response.write "URL : " & targetURL & "<br>"
		response.write "택배사코드 : " & songjangDiv & "<br><br>"
	end if

	if (targetURL = "") then
		errCode = "100"
		errMSG = "택배사 오류"
		exit function
	end if

	if (CStr(songjangDiv) = "9") or (CStr(songjangDiv) = "21") or (CStr(songjangDiv) = "34") then
		'// KGB(9), 엘로우(13), 경동합동(21), 대신화물택배(34)
		targetHTML = FnGetRecvCheckHTML_EUCKR(targetURL)
	else
		targetHTML = FnGetRecvCheckHTML(targetURL)
	end if

	if (IsDebugMode) then
		response.write "<!-- " & targetHTML & "--><br><br>"
	end if

	if (targetHTML = "") then
		errCode = "101"
		errMSG = "택배조회 URL 오류"
		exit function
	end if

	parseResult = FnCheckRecvState(songjangDiv, targetHTML)
	if (IsDebugMode) then
		response.write "parseResult : " & parseResult & "<br><br>"
	end if

	if (parseResult = "T") then
		FnCheckNSaveRecvState = True
	end if

end function

%>
