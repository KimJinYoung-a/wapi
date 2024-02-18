<!-- #include virtual="/lib/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<% if NOT (IsAutoScript) then %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<% end if %>
<%

''임시 중지
'response.write "일시중지 2016/11/25"
'dbget.close():
'response.end


dim entrId      : entrId="10X10"                    ''고정값 : 제휴업체_ID
dim ordclmNo    : ordclmNo=request("ordclmNo")      ''인터파크 주문번호
dim ordSeq      : ordSeq=request("ordSeq")          ''인터파크 주문순번
dim delvDt      : delvDt=request("delvDt")          ''YYYYMMDD 출고완료일자
dim delvEntrNo  : delvEntrNo=request("delvEntrNo")  ''택배사코드
dim invoNo      : invoNo=request("invoNo")          ''운송장번호 숫자만 가능함.
dim optPrdTp    : optPrdTp=request("optPrdTp")      '' 옵션상품유형	01 (일반단품상품)  02 (추가옵션이포함된부모상품)
dim optOrdSeqList  : optOrdSeqList=request("optOrdSeqList") ''주문순번리스트


'2013/02/28 진영추가
dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ordclmNo")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ordSeq")&"'"
	sqlStr = sqlStr & "	and sellsite='interpark'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If
'2013/02/28 진영추가 끝



'''01 일반단붐상품일 경우 해당 주문순번  02 추가옵션이포함된부모상품일 경우 부모상품의 주문순번 및 추가옵션상품의 주문순번
''* 주문순번이 하나 이상일 경우 구분자(“|”)로 구문한다.
''ex) sc.optOrdSeqList=1|2|3|4


invoNo= trim(replace(invoNo,"-",""))

	'// 변수선언
	dim strSql, actCnt, lp
    dim AssignedCNT, paramAdd

	actCnt = 0			'실갱신건수

'Dim ordclmNo : ordclmNo = ord_no
'ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
'ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
'ord_no = replace(ord_no,"_2","")
'ord_no = replace(ord_no,"_3","")
'ord_no = replace(ord_no,"_4","")



paramAdd = "&sc.entrId="+entrId
if (Right(ordclmNo,2)="_1") then
    paramAdd = paramAdd + "&sc.ordclmNo="+replace(ordclmNo,"_1","")
else
    paramAdd = paramAdd + "&sc.ordclmNo="+ordclmNo
end if
paramAdd = paramAdd + "&sc.ordSeq="+ordSeq
paramAdd = paramAdd + "&sc.delvDt="+delvDt
paramAdd = paramAdd + "&sc.delvEntrNo="+delvEntrNo
paramAdd = paramAdd + "&sc.invoNo="+invoNo
paramAdd = paramAdd + "&sc.optPrdTp="+optPrdTp
paramAdd = paramAdd + "&sc.optOrdSeqList="+optOrdSeqList


Dim iResult, iMessage
Dim iParkURL, iParams, replyXML, ErrMsg

''iParkURL = "http://www.interpark.com/order/OrderClmAPI.do"
iParkURL = "https://joinapi.interpark.com/order/OrderClmAPI.do"   '''실서버는 https://joinapi.interpark.com
iParams  = "_method=delvCompForComm" & paramAdd


'rw "delvEntrNo="&delvEntrNo
''rw iParkURL &"?"& iParams
''response.end


replyXML = SendReqGet(iParkURL, iParams)

    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

    if (ErrMsg<>"") then
        if (IsAutoScript) then
            rw "인터파크 송장입력중  오류가 발생했습니다. "&ordclmNo&" "&ordclmNo&"_"&ordSeq
        else
    	    Response.Write "<script language=javascript>alert('인터파크 송장입력중  오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
    	ENd IF
    	dbget.Close: Response.End

    else

    	'XML을 담을 DOM 객체 생성
    	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
 ''response.end

        ''결과가 빈값 (20120702) 으로 왔음.. :: 성공시?
        IF (Trim(replyXML)="") THEN
            rw "결과 빈값"
            iResult="1"
        ELSE
        	xmlDOM.LoadXML replyXML

        	if Err<>0 then
        		Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
        		dbget.Close: Response.End
        	end if
            On Error Resume next
        	iResult		= Trim(xmlDOM.getElementsByTagName("CODE").item(0).text)			'결과
            iMessage    = xmlDOM.getElementsByTagName("MESSAGE").Item(0).Text
        	On Error Goto 0
        END IF
    '	'// 수정
        IF (iResult="000") THEN
        	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=1"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
            strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
            strSql = strSql & "	and sellsite='interpark'"
            strSql = strSql & "	and matchstate in ('O')"
       ''rw strSql
        	dbget.Execute strSql,AssignedCNT
        	actCnt = actCnt+AssignedCNT

        	IF (AssignedCNT>0) then
        	    if (IsAutoScript) then
        	        rw "OK|"&ordclmNo&"_"&ordSeq
        	    ELSE
            	    response.write "OK"
            	ENd IF
            ENd IF
        ELSE
            ''이미 송장이 입력된경우.// 무조건 N번 재시도로 변경.
            ''송장 늦게 입력 : 주문일로부터 30일이 경과하여 발송완료 불가합니다. 고객센터에 문의하시기 바랍니다. 문의처) 02-3289-3236
            ''취소된경우 : 발송완료 처리할 주문이 존재하지 않습니다.
            ''Maybe 고객수령 : 현재 주문내역상태가 발송완료를 처리할 수 없습니다
            IF LEft(iMessage,Len("이미 배송처리가 되었습니다."))="이미 배송처리가 되었습니다." THEN
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendState=9"
            	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
                strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
                strSql = strSql & "	and sellsite='interpark'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')"  ''A 추가
'rw  strSql
            	dbget.Execute strSql,AssignedCNT

            	if (IsAutoScript) then
            	    rw "SKIP 처리 "&ordclmNo&" "&ordSeq
            	ELSE
            	    rw "SKIP 처리"
            	ENd IF

            ELSE
                '' 시도 회수 추가 sendReqCnt
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
                strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
                strSql = strSql & "	and sellsite='interpark'"
                strSql = strSql & "	and matchstate in ('O','C','Q')" '','A'
'rw strSql
            	dbget.Execute strSql

                iMessage = "<font color=red>"&iMessage&"</font>"
            END IF

            if (IsAutoScript) then
                rw "iMessage="&iMessage&":"&ordclmNo&" "&ordclmNo&"_"&ordSeq
            else
                rw "iMessage="&iMessage
            ENd IF

'2013/02/28 김진영 추가
'만약 에러횟수가 3회가 넘으면서 minusOrderSerial이 공백일 때 해당
'updateSendState = 901		발송처리누락 수기등록건
'updateSendState = 902		취소후 제결제건
'updateSendState = 903		반품처리건
			Dim errCount
			strSql = ""
			strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
			strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
			strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
			strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
			rsget.CursorLocation = adUseClient
			rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
			If Not rsget.Eof Then
				errCount = rsget("cnt")
			End If
			rsget.Close

			If errCount > 0 Then
				response.write  "<select name='updateSendState' id=""updateSendState"">" &_
								"	<option value=''>선택</option>" &_
								"	<option value='901'>발송처리누락 수기등록건</option>" &_
								"	<option value='902'>취소후 제결제건</option>" &_
								"	<option value='903'>반품처리건</option>" &_
								"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='완료처리' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('선택해주세요');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actInterparkSongjangInputProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
				response.write "    location.replace(uri);"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			End If
'2013/02/28 김진영 추가 끝

        END IF
    	Set xmlDOM = Nothing
    end if
%>
<% if NOT (IsAutoScript) then %>
</body>
</html>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->