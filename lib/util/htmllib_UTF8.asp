<%
'+----------------------------------------------------------------------------------------------------------------------+
'|                                               HTML 공 통   함 수 선 언                                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'|             함 수 명                      |                          기    능                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | 날짜형식을 지정된 문자형으로 변환                            |
'|                                          | 사용예 : printdate = FormatDate(now(),"0000.00.00")          |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid)  | 이미지파일의 서브 폴더명을 반환한다.                                     |
'|                                           | 사용예 : SubFolder = GetImageSubFolderByItemid(1126)                     |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| db2html(checkvalue)                       | DB의 내용을 HTML에 사용할 수 있도록 변환                                 |
'|                                           | 사용예 : Contents = db2html("DB의 내용")                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| html2db(checkvalue)                       | 사용자가 입력한 내용을 DB에 넣을 수 있도록 변환                          |
'|                                           | 사용예 : Contents = html2db("저장할 내용")                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| nl2br(checkvalue)                         | 내용의 새줄(vbCrLf)을 "<br>"태그로 치환하여 반환                         |
'|                                           | 사용예 : Contents = nl2br("내용")                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| CurrFormat(byVal v)                       | 숫자를 3자리 구분의 문자열로 변환                                        |
'|                                           | 사용예 : strNum = CurrFormat(1230) → "1,230"                            |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Format00(n,orgData)                       | 숫자를 0으로 채워진 지정된 길이의 문자열로 변환                          |
'|                                           | 사용예 : strNum = Format00(5,123) → "00123"                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatCode(itemcode)                      | 제품 일련번호를 6자리의 문자열로 변환                                    |
'|                                           | 사용예 : itemCode = FormatCode(2654) → "002654"                         |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetCurrentTimeFormat()                    | 현재시간을 문자열로 반환 (yyyymmddhhmmss)                                |
'|                                           | 사용예 : strNow = GetCurrentTimeFormat() → "20060508101833"             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetListImageUrl(byval itemid)             | 제품번호에 맞는 리스트 이미지 및 폴더 반환                               |
'|                                           | 사용예 : img = GetListImageUrl("53100") → "/image/list/L000053100.jpg"  |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| DDotFormat(byval str,byval n)             | 내용을 지정한 길이로 자른다.                                             |
'|                                           | 사용예 : strShort = DDotFormat("내용입니다.",3) → "내용입..."           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| stripHTML(strng)                          | 내용 중 HTML태그를 없앤다.                                               |
'|                                           | 사용예 : Contents = stripHTML("<b>내용</b>") → " 내용 "                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getFileExtention(strFile)                 | 파일명의 확장자를 반환한다.                                              |
'|                                           | 사용예 : ext = getFileExtention("123.jpg") → "jpg"                      |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)   		 | 숫자를 지정한 길이의 문자열로 변환한다.                      			|
'|                                   		 | 사용예 : Num2Str(425,4,"0","R") → 0425                      			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)    | like iif function                                                        |
'|                                           | 사용예 : ChkIIF(1>2,"a","b") → "b"                                       |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_return(strMSG)                      | 경고창 띄운후 이전으로 돌아간다.                            				|
'|                                           | 사용예 : Call Alert_return("뒤로 돌아갑니다.")               			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_close(strMSG)                       | 경고창 띄운후 현재창을 닫는다.                               			|
'|                                           | 사용예 : Call Alert_close("창을 닫습니다.")                  			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)              | 경고창 띄운후 지정페이지로 이동한다.                         			|
'|                                           | 사용예 : Call Alert_move("이동합니다.","/index.asp")         			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                   | 지정길이로 문자열 자르기                                                 |
'|                                           | 사용예 : chrbyte("안녕하세요",3,"Y") → 안녕...                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplex(uid,pwd)               | 비밀번호 정책의 복잡성을 만족하는지 검사하고 그 이유를 반환              |
'|                                           | 사용예 : chkPasswordComplex("kobula","abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkWord(str,patrn)                        | 문자열의 형식을 정규식으로 검사                                          |
'|                                           | 사용예 : chkWord("abcd","[^-a-zA-Z0-9/ ]") : 영어숫자만 허용             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ParsingPhoneNumber(str,patrn)             | 전화번호에 대시 추가                                                     |
'|                                           | 사용예 : ParsingPhoneNumber("0112223333") :                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracket(strng)                     | 꺽은괄호 태그로 치환('<', '>')                                           |
'|                                           | 사용예 : ReplaceBracket("<>") → &lt;&gt;                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getNumeric(strNum)                        | 문자열에서 숫자만 추출 변환                                              |
'|                                           | 사용예 : getNumeric("a45d61*124") -> 461124                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| 특수 문자 제거(' ,--)                                        				|
'|                                          | 사용예 : cont = ReplaceRequestSpecialChar(Rs("strng"))       				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | 내용에 금지된 HTML태그가 있는지 검사                         				|
'|                                          | 사용예 : checkNotValidHTML("<script...") → true             				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| minutechagehour(v)                 		| 분단위를 시간단위으로 짤라서 반환                      					|
'|                                          | 사용예 : minutechagehour(v)             									|
'+-------------------------------------------+--------------------------------------------------------------------------+

'// 날짜를 지정된 문자형으로 변환 //
function FormatDate(ddate, formatstring)
	dim s
	Select Case formatstring
		Case "0000.00.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "0000-00-00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000"
			s = CStr(year(ddate)) &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000000000"
			s = CStr(year(ddate))  &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")  &_
				Num2Str(hour(ddate),2,"0","R")  &_
				Num2Str(minute(ddate),2,"0","R") &_
				Num2Str(Second(ddate),2,"0","R")		
		Case "0000.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R")
		Case "0000.00.00-00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & "-" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00.00 00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000/00/00"
			s = CStr(year(ddate)) & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00/00"
			s = Num2Str(year(ddate),2,"0","R") & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00.00"
			s = Num2Str(year(ddate),2,"0","R") & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00"
			s = Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00"
			s = Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case Else
			s = CStr(ddate)
	End Select

	FormatDate = s
end function

function GetImageSubFolderByItemid(byval iitemid)
	IF iitemid<>"" THEN
	GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	END IF
end function

'' 기존 디비에 이전 형식 있음.. 차후 삭제
function db2html(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&amp;", "&")
    v = replace(v, "&lt;", "<")
    v = replace(v, "&gt;", ">")
    v = replace(v, "&quot;", "'")
    v = Replace(v, "", "<br>")
    v = Replace(v, "\0x5C", "\")
    v = Replace(v, "\0x22", "'")
    v = Replace(v, "\0x25", "'")
    v = Replace(v, "\0x27", "%")
    v = Replace(v, "\0x2F", "/")
    v = Replace(v, "\0x5F", "_")
    ''checkvalue = Replace(checkvalue, vbcrlf,"<br>")
    db2html = v
end function

'' 2008 03 수정 - Eastone
function html2db(checkvalue)
	html2db = Newhtml2db(checkvalue)
end function

function Newhtml2db(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function
	v = Replace(v, "'", "''")
	Newhtml2db = v
end function


function nl2br(checkvalue)
	if IsNull(checkvalue) then
		nl2br = ""
		Exit function
	end if

	checkvalue = Replace(checkvalue, vbcrlf,"<br>")
	nl2br = checkvalue
end function

'// 문자열내 CR/LF를 공백으로 치환 //
function nl2blank(v)
	if IsNull(v) then
		nl2blank = ""
		Exit function
	end if

    nl2blank = Replace(v, vbcrlf,"")
end function

function CurrFormat(byVal v)
        if ((v = "") or (isnull(v))) then
                CurrFormat = 0
        else
                CurrFormat = FormatNumber(FormatCurrency(v),0)
        end if
end function


function Format00(n,orgData)
    dim tmp
    
    if IsNULL(orgData) then Exit function
    
	if (n-Len(CStr(orgData))) < 0 then
		Format00 = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	Format00 = tmp
end function


function FormatCode(itemcode)
	FormatCode = Format00(6,itemcode)
end function


function GetCurrentTimeFormat()
	dim d
	d = now
	GetCurrentTimeFormat = replace(Left(FormatDateTime(d,2),7),"-","") + Format00(2,Day(d)) + Format00(2,Hour(d)) + Format00(2,Minute(d))  +  Format00(2,Second(d))

end function


function GetListImageUrl(byval itemid)
	GetListImageUrl = "/image/list/L" + Format00(9,itemid) + ".jpg"
end function


function DDotFormat(byval str,byval n)
	DDotFormat = str
	if Len(str)> n then
		DDotFormat = Left(str,n) + "..."
	end if
end function


function stripHTML(strng)
   Dim regEx
   Set regEx = New RegExp
   regEx.Pattern = "[<][^>]*[>]"
   regEx.IgnoreCase = True
   regEx.Global = True
   stripHTML = regEx.Replace(strng, " ")
   Set regEx = nothing
End Function

function Format00(n,orgData)
    dim tmp
    
    if IsNULL(orgData) then Exit function
    
	if (n-Len(CStr(orgData))) < 0 then
		Format00 = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	Format00 = tmp
end function

function getFileExtention(strFile)
	Dim file_length, file_point, ext_len
	
	if Not(strFile="" or isNull(strFile)) then
		file_length = LEN(strFile)
		file_point = inStrRev(strFile,".") + 1
		ext_len = file_length - file_point + 1
	
		getFileExtention = Lcase(MID(strFile,file_point,ext_len))
	end if
End Function

function adminColor(v)
	adminColor = "#FFFFFF"

	if v="menubar" then
		adminColor = "#DEDFFF"
	elseif v="menubar_left" then
		adminColor = "#CCCCCC"
	elseif v="topbar" then
		adminColor = "#F4F4F4"
	elseif v="tabletop" then
		adminColor = "#E6E6E6"
	elseif v="tablebg" then
		adminColor = "#999999"
			
	elseif v="pink" then
		adminColor = "#FFDDDD"
	elseif v="green" then
		adminColor = "#DDFFDD"
	elseif v="sky" then
		adminColor = "#DDDDFF"
	elseif v="gray" then
		adminColor = "#EEEEEE"
	elseif v="dgray" then
		adminColor = "#CCCCCC"
		
	else

	end if
end function
	
	'// 숫자를 지정한 길이의 문자열로 반환 //
	Function Num2Str(inum,olen,cChr,oalign)
		Dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""
		
		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if
		
		'결합방법에따른 결과 분기
		if oalign="L" then
			'왼쪽기준
			Num2Str = inum & strChr
		else
			'오른쪽 기준 (기본값)
			Num2Str = strChr & inum
		end if

    End Function


'// 문자열을 잘라 원하는 위치의 값을 반환 //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)
    
    if UBound(buf)<pos then Exit function
    
    SplitValue = buf(pos)
end function


'// 파라메터 길이 체크 후 Maxlen 이하로 돌려줌 Code, id 등의 Param 에 사용 //
function requestCheckVar(orgval,maxlen)
	requestCheckVar = trim(orgval)
	requestCheckVar = replace(requestCheckVar,"'","")
	requestCheckVar = replace(requestCheckVar,"--","")
	requestCheckVar = Left(requestCheckVar,maxlen)
end function


'// 값비교 후 Return 값 like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function

'// 경고문 출력후 뒤로가기 //
Sub Alert_return(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"history.back();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// 경고문 출력후 창닫기 //
Sub Alert_close(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// 경고문 출력후 지정 페이지로 이동 //
Sub Alert_move(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.location='" & targetURL & "';" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// 지정길이로 문자열 자르기 //
Function chrbyte(str,chrlen,dot)

    Dim charat, wLen, cut_len, ext_chr, cblp

    if IsNULL(str) then Exit function

    for cblp=1 to len(str)
        charat=mid(str, cblp, 1)
        if asc(charat)>0 and asc(charat)<255 then
            wLen=wLen+1
        else
            wLen=wLen+2
        end if

        if wLen >= cint(chrlen) then
           cut_len = cblp
           exit for
        end if
    next

    if len(cut_len) = 0 then
        cut_len = len(str)
    end if

	if len(str)>cut_len and dot="Y" then
		ext_chr = "..."
	else
		ext_chr = ""
	end if

    chrbyte = Trim(left(str,cut_len)) & ext_chr

end function


'// 패스워드 복잡성 검사 함수
Function chkPasswordComplex(uid,pwd)
	dim msg, i, sT, sN
	msg = ""

	'비밀번호 길이 검사
	if len(pwd)<6 then
		msg = msg & "- 비밀번호는 최소 6자리이상으로 입력해주세요.\n"
	end if

	'아이디와 동일 또는 포함하고 있는가?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- 아이디와 동일하거나 아이디를 포함하고 있는 비밀번호입니다.\n"
	end if

	'## 복잡성을 만족하는가?
	'같은문자 3번 연속 금지
	sT=""
	sN=0
	for i=1 to len(pwd)
		if st=mid(pwd,i,1) then
			sN = sN +1
		else
			sN = 0
		end if
		st = mid(pwd,i,1)
		if sN>=2 then
			msg = msg & "- 같은문자가 3번 연속으로 쓰였습니다.\n"
			exit for
		end if
	next
	'영문/숫자의 조합
	if chkWord(pwd,"[^-a-zA-Z]") or chkWord(pwd,"[^-0-9 ]") then
		msg = msg & "- 비밀번호는 반드시 알파벳과 숫자를 조합해서 만들어야합니다.\n"
	end if

	'결과 반환
	chkPasswordComplex = msg
end Function

'//정규식 문자열 검사
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp
    regEx.Pattern = patrn	' 패턴을 설정.
    regEx.IgnoreCase = True	' 대/소문자를 구분하지 않도록 .
    regEx.Global = True		' 전체 문자열을 검색하도록 설정.
    SET Matches = regEx.Execute(str)
	if 0 < Matches.count then
		chkWord= false
	Else
		chkWord= true
	end if

	'pattern0 = "[^가-힣]"  '한글만
	'pattern1 = "[^-0-9 ]"  '숫자만
	'pattern2 = "[^-a-zA-Z]"  '영어만
	'pattern3 = "[^-가-힣a-zA-Z0-9/ ]" '숫자와 영어 한글만 
	'pattern4 = "<[^>]*>"   '태그만
	'pattern5 = "[^-a-zA-Z0-9/ ]"    '영어 숫자만
End Function 

'// 전화번호에 대시 추가
function ParsingPhoneNumber(orgnum)
    dim noDashNum, PreNum, CuttedNum
    noDashNum = Replace(orgnum,"-","")
    
    ParsingPhoneNumber = noDashNum
    
    if Len(noDashNum)<7 then
        exit function
    end if
    
    if Len(noDashNum)=7 then
        ParsingPhoneNumber = Left(noDashNum,3) & "-" & Right(noDashNum,4)
        Exit function    
    end if
    
    if Len(noDashNum)=8 then
        ParsingPhoneNumber = Left(noDashNum,4) & "-" & Right(noDashNum,4)
        Exit function    
    end if
    
    if (Left(noDashNum,1)<>"0") then
        Exit function  
    end if
    
    PreNum = Left(noDashNum,2)
    if (PreNum="02") then
        CuttedNum = Mid(noDashNum,3,255)
    else
        PreNum = Left(noDashNum,3)
        if (PreNum="010") or (PreNum="011") or (PreNum="016") or (PreNum="017") or (PreNum="019") then
            CuttedNum = Mid(noDashNum,4,255)
        else
            CuttedNum = Mid(noDashNum,4,255)
        end if
    end if
    
    if Len(CuttedNum)=7 then
        ParsingPhoneNumber = PreNum & "-" & Left(CuttedNum,3) & "-" & Right(CuttedNum,4)
    elseif Len(CuttedNum)=8 then
        ParsingPhoneNumber = PreNum & "-" & Left(CuttedNum,4) & "-" & Right(CuttedNum,4)
    else
        exit function
    end if
end function


'''''==================  2009 추가

' response.write 함수
Function rw(ByVal str)
	response.write str & "<br>"
End Function 

' Null을 공백으로 치환
Function null2blank(ByVal v)
	If IsNull(v) Then 
		null2blank = ""
	Else 
		null2blank = v
	End If 
End Function 

'// 큰따옴표 input 박스 value=""에 사용할때 치환
Function doubleQuote(ByVal v)
	If IsNull(v) Then 
		doubleQuote = ""
	Else 
		doubleQuote = Replace(v, """","&quot;")
	End If 
End Function 


' request 대체 함수(파라미터명, 디폴트값)
Function req(ByVal param, ByVal value)
'	VarType Return 값
'	0 (공백)
'	1 (널)
'	2 integer
'	3 Long
'	4 Single
'	5 Double
'	6 Currency
'	7 Date
'	8 String
'	9 OLE Object
'	10 Error
'	11 Boolean
'	12 Variant
'	13 Non-OLE Object
'	17 Byte
'	8192 Array

	Dim tmpValue
	
	If VarType(value) = 2 Or VarType(value) = 3 Or VarType(value) = 4 Or VarType(value) = 5 Or VarType(value) = 6 Then 
		tmpValue = Replace(Trim(Request(param)),",","")
		If Not IsNumeric(tmpValue) Then	' 숫자가 아니면
			tmpValue = value
		End If
		tmpValue = CDbl(tmpValue)
	Else 
		tmpValue = Trim(Request(param))
		If tmpValue = "" Then			' Request값이 없으면
			tmpValue = value
		End If
	End If 
	req = tmpValue

End Function

Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage)

	'변수 선언
	Dim intCurrentPage, strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop
    
	'현재 페이지 설정
	intCurrentPage = Mid(strCurrentPage, InStr(strCurrentPage, "=")+1)		'현재 페이지 값
	strCurrentPage = Left(strCurrentPage, InStr(strCurrentPage, "=")-1)		'페이지 폼값 변수명
	
	'현재 페이지 명
	strCurrentPath = Request.ServerVariables("Script_Name")
		
	'해당페이지에 표시되는 시작페이지와 마지막페이지 설정
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage
	
	'총 페이지 수 설정
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))
	
	'폼 설정 & hidden 파라미터 설정
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>"			'현재 페이지
		
	'파라미터 값들(예: 검색어)을 hidden 파라미터로 저장한다
	strParamName = ""
	For Each strParamName In Request.Form	
		If strParamName <> strCurrentPage Then
			
			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""
	
	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then			
			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"		
		END IF	
	Next
		
	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class=a><tr align='center'><td>"

	'이전 페이지 이미지 설정
	If intStartBlock > 1 Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pageprev01.gif' border='0' style='cursor:hand' alt='이전 " & intBlockPerPage & " 페이지'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();'>"
	Else
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif' border='0' >"
	End If

	Response.Write "</td><td>&nbsp;"
	
	'페이징 출력
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For
			
			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"
			
			If Int(intLoop) = Int(intCurrentPage) Then		'현재 페이지
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'그 외 페이지
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If
		
		Next
	Else		'한 페이지만 존재 할때
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'다음 페이지 이미지 설정
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pagenext01.gif' border='0' style='cursor:hand' alt='다음 " & intBlockPerPage & " 페이지'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'>"
	Else
	    Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif' border='0' >"
	End If
	
	Response.Write "</td></tr></form></table>"

End Sub



' 등록,수정,삭제 모드 텍스트 리턴
Function getModeName(ByVal mode)
    Select Case mode
        Case "INS"	: getModeName = "등록"
        Case "UPD"	: getModeName = "수정"
        Case "DEL"	: getModeName = "삭제"
        Case "FIN"	: getModeName = "완료"
        Case Else	: getModeName = "미정"
    End Select
End Function 

'// 꺽은괄호 HTML코드로 치환 //
Function ReplaceBracket(strng)
	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function


' 정규식 함수
Function ReplaceText(str, patrn, repStr)
	Dim regEx
	Set regEx = New RegExp
	with regEx
		.Pattern = patrn
		.IgnoreCase = True
		.Global = True
	End with
	ReplaceText = regEx.Replace(str, repStr)
End Function 

Function TwoNumber(number)
	Dim vNumber
	If len(number) = 1 Then
		vNumber = "0" & number
	Else
		vNumber = number
	End If
	TwoNumber = vNumber
End Function

'// 문자열에서 숫자만 추출 변환
Function getNumeric(strNum)
	Dim lp, tmpNo, strRst
	For lp=1 to len(strNum)
		tmpNo = mid(strNum, lp, 1)
		if asc(tmpNo)>47 and asc(tmpNo)<58 then
			strRst = strRst & tmpNo
		end if
	Next
	getNumeric = strRst
End Function

Function getUserLevelCSS(iuserLevel)
    if IsNULL(iuserLevel) then
        getUserLevelCSS = "member_no"
        exit function
    end if

    Select Case CStr(iuserLevel)
		Case "5"
			getUserLevelCSS = "member_orange"
		Case "0"
			getUserLevelCSS = "member_yellow"
		Case "1"
			getUserLevelCSS = "member_green"
		Case "2"
			getUserLevelCSS = "member_blue"
		Case "3"
			getUserLevelCSS = "member_vipsilver"
            ''getUserLevelCSS = "member_vip"
		Case "4"
			getUserLevelCSS = "member_vipgold"
		Case "7"
			getUserLevelCSS = "member_staff"
		Case "6"
			getUserLevelCSS = "member_red"
		Case "8"
			getUserLevelCSS = "member_red"
		Case "9"
			getUserLevelCSS = "member_red"
		Case Else
			getUserLevelCSS = "member_orange"
	end Select
end function

'//문자열내 특수문자 제거
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//올림 함수
function ceil(Pnanum,nanum)
Dim result1, result2, variant_return

 result1 = Pnanum/nanum
 result2 = round(Pnanum/nanum)

 if result1 <> result2 then
  variant_return = fix(result1) + 1
 else
  variant_return = result1
 end if
ceil = variant_return
end function
 
'//올림 함수
function ceilValue(iValue) 
 if iValue <>  round(iValue) then
  ceilValue = fix(iValue) + 1
 else
  ceilValue = iValue
 end if 
end function

'// 지정수만큼 지정한 문자로 바꿈)
Function printUserId(strID,lng,chr)
	dim le, te

	le = len(strID)
	if(le<lng) Then
		printUserId = String(lng, le)
		Exit Function
	end if

	te = left(strID,le-lng) & String(lng, chr)
	printUserId = te

End Function

'// 내용에 금지된 HTML태그가 있는지 검사 //
function checkNotValidHTML(ostr)
	checkNotValidHTML = false

	dim LcaseStr
	LcaseStr = Lcase(ostr)
	LcaseStr = Replace(LcaseStr," ","")

	if InStr(LcaseStr,"<script")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"<object")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"</iframe>")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"<iframe>")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"iframe")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"imgsrc")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"ahref")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,".wmf")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,".js")>0 then
		checkNotValidHTML = true
	end if
end function

'// 경고문 출력후 창닫고 오픈창 리로드 -2011.02.23 정윤정추가 //
Sub Alert_closenreload(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.reload();"& vbCrLf &_
			"self.close();" & vbCrLf &_ 
			"</script>"
	Response.Write strTemp
End Sub

'// 경고문 출력후 창닫고 오픈창 타겟주소로 이동 -2011.02.23 정윤정추가 //
Sub Alert_closenmove(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.href ='" & targetURL & "';" & vbCrLf &_ 
			"self.close();" & vbCrLf &_ 
			"</script>"
	Response.Write strTemp
End Sub

'//분단위를 시간단위으로 짤라서 반환	'/2011.03.31 한용민 생성
function minutechagehour(v)
	dim tmpval , tmph , tmpm
		
	if v = "" or isnull(v) or v = 0 then
		minutechagehour = ""
	else	
		tmph = int(v / 60)	'시간단위
		tmpm = v - (tmph * 60)	'분단위
		
		if tmph <> 0 then tmpval = tmpval & tmph & "시간 "
		if tmpm <> 0 then tmpval = tmpval & tmpm & "분"
			
		minutechagehour = tmpval
	end if		
end function

'// 사내 접속여부
Function isTenbyTenConnect()
	Dim conIp, arrIp, tmpIp
	conIp = Request.ServerVariables("REMOTE_ADDR")
	if left(conIp,2)<>"::" then
		arrIp = split(conIp,".")
		tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")
	end if

	'121.78.103.60 : 15층 유선
	'10.10.10.36 : m2서버
	'192.168.1.x : 15층 운영,개발,인사,재무
	'192.168.6.x : 15층 일반망
	'110.11.187.233 : 15층 wireless6
	'110.93.128.x : IDC

	if tmpIp="121078103060" or tmpIp="110011187233" or (tmpIp=>"110093128001" and tmpIp<="110093128256") or (tmpIp=>"192168001001" and tmpIp<="192168001256") or (tmpIp=>"192168006001" and tmpIp<="192168006256") then
		isTenbyTenConnect = True
	else
		isTenbyTenConnect = False
	end if
End Function

'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' 서버다운
		isServerDown = false	' 서버활성화
		if isTenbyTenConnect then isServerDown = false	'사내접속 허용

	if Not(isServerDown) then exit Sub

	response.write "서비스 점검중입니다"
	response.end
End Sub

%>

