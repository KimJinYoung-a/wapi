<%
'+----------------------------------------------------------------------------------------------------------------------+
'|                                               HTML �� ��   �� �� �� ��                                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'|             �� �� ��                      |                          ��    ��                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | ��¥������ ������ ���������� ��ȯ                            |
'|                                          | ��뿹 : printdate = FormatDate(now(),"0000.00.00")          |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid)  | �̹��������� ���� �������� ��ȯ�Ѵ�.                                     |
'|                                           | ��뿹 : SubFolder = GetImageSubFolderByItemid(1126)                     |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| db2html(checkvalue)                       | DB�� ������ HTML�� ����� �� �ֵ��� ��ȯ                                 |
'|                                           | ��뿹 : Contents = db2html("DB�� ����")                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| html2db(checkvalue)                       | ����ڰ� �Է��� ������ DB�� ���� �� �ֵ��� ��ȯ                          |
'|                                           | ��뿹 : Contents = html2db("������ ����")                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| nl2br(checkvalue)                         | ������ ����(vbCrLf)�� "<br>"�±׷� ġȯ�Ͽ� ��ȯ                         |
'|                                           | ��뿹 : Contents = nl2br("����")                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| CurrFormat(byVal v)                       | ���ڸ� 3�ڸ� ������ ���ڿ��� ��ȯ                                        |
'|                                           | ��뿹 : strNum = CurrFormat(1230) �� "1,230"                            |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Format00(n,orgData)                       | ���ڸ� 0���� ä���� ������ ������ ���ڿ��� ��ȯ                          |
'|                                           | ��뿹 : strNum = Format00(5,123) �� "00123"                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatCode(itemcode)                      | ��ǰ �Ϸù�ȣ�� 6�ڸ��� ���ڿ��� ��ȯ                                    |
'|                                           | ��뿹 : itemCode = FormatCode(2654) �� "002654"                         |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetCurrentTimeFormat()                    | ����ð��� ���ڿ��� ��ȯ (yyyymmddhhmmss)                                |
'|                                           | ��뿹 : strNow = GetCurrentTimeFormat() �� "20060508101833"             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetListImageUrl(byval itemid)             | ��ǰ��ȣ�� �´� ����Ʈ �̹��� �� ���� ��ȯ                               |
'|                                           | ��뿹 : img = GetListImageUrl("53100") �� "/image/list/L000053100.jpg"  |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| DDotFormat(byval str,byval n)             | ������ ������ ���̷� �ڸ���.                                             |
'|                                           | ��뿹 : strShort = DDotFormat("�����Դϴ�.",3) �� "������..."           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| stripHTML(strng)                          | ���� �� HTML�±׸� ���ش�.                                               |
'|                                           | ��뿹 : Contents = stripHTML("<b>����</b>") �� " ���� "                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getFileExtention(strFile)                 | ���ϸ��� Ȯ���ڸ� ��ȯ�Ѵ�.                                              |
'|                                           | ��뿹 : ext = getFileExtention("123.jpg") �� "jpg"                      |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)   		 | ���ڸ� ������ ������ ���ڿ��� ��ȯ�Ѵ�.                      			|
'|                                   		 | ��뿹 : Num2Str(425,4,"0","R") �� 0425                      			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)    | like iif function                                                        |
'|                                           | ��뿹 : ChkIIF(1>2,"a","b") �� "b"                                       |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_return(strMSG)                      | ���â ����� �������� ���ư���.                            				|
'|                                           | ��뿹 : Call Alert_return("�ڷ� ���ư��ϴ�.")               			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_close(strMSG)                       | ���â ����� ����â�� �ݴ´�.                               			|
'|                                           | ��뿹 : Call Alert_close("â�� �ݽ��ϴ�.")                  			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)              | ���â ����� ������������ �̵��Ѵ�.                         			|
'|                                           | ��뿹 : Call Alert_move("�̵��մϴ�.","/index.asp")         			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                   | �������̷� ���ڿ� �ڸ���                                                 |
'|                                           | ��뿹 : chrbyte("�ȳ��ϼ���",3,"Y") �� �ȳ�...                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplex(uid,pwd)               | ��й�ȣ ��å�� ���⼺�� �����ϴ��� �˻��ϰ� �� ������ ��ȯ              |
'|                                           | ��뿹 : chkPasswordComplex("kobula","abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkWord(str,patrn)                        | ���ڿ��� ������ ���Խ����� �˻�                                          |
'|                                           | ��뿹 : chkWord("abcd","[^-a-zA-Z0-9/ ]") : ������ڸ� ���             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ParsingPhoneNumber(str,patrn)             | ��ȭ��ȣ�� ��� �߰�                                                     |
'|                                           | ��뿹 : ParsingPhoneNumber("0112223333") :                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracket(strng)                     | ������ȣ �±׷� ġȯ('<', '>')                                           |
'|                                           | ��뿹 : ReplaceBracket("<>") �� &lt;&gt;                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getNumeric(strNum)                        | ���ڿ����� ���ڸ� ���� ��ȯ                                              |
'|                                           | ��뿹 : getNumeric("a45d61*124") -> 461124                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| Ư�� ���� ����(' ,--)                                        				|
'|                                          | ��뿹 : cont = ReplaceRequestSpecialChar(Rs("strng"))       				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | ���뿡 ������ HTML�±װ� �ִ��� �˻�                         				|
'|                                          | ��뿹 : checkNotValidHTML("<script...") �� true             				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| minutechagehour(v)                 		| �д����� �ð��������� ©�� ��ȯ                      					|
'|                                          | ��뿹 : minutechagehour(v)             									|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| BinaryToText(BinaryData, CharSet)         | ���̳ʸ� ������ TEXT���·� ��ȯ                                          |
'|                                           | ��뿹 : BinaryToText(objXML.ResponseBody, "euc-kr")                     |
'+-------------------------------------------+--------------------------------------------------------------------------+

'// ��¥�� ������ ���������� ��ȯ //
function FormatDate(ddate, formatstring)
	dim s
	Select Case formatstring
		Case "0000-00-00 00:00:00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
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

'' ���� ��� ���� ���� ����.. ���� ����
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

'' 2008 03 ���� - Eastone
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

'// ���ڿ��� CR/LF�� �������� ġȯ //
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
    if isNULL(itemcode) then
        FormatCode = itemcode
        Exit function
    end if

    if (itemcode>=1000000) then
        FormatCode = Format00(8,itemcode)
    else
	    FormatCode = Format00(6,itemcode)
    end if
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

	'// ���ڸ� ������ ������ ���ڿ��� ��ȯ //
	Function Num2Str(inum,olen,cChr,oalign)
		Dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""

		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if

		'���չ�������� ��� �б�
		if oalign="L" then
			'���ʱ���
			Num2Str = inum & strChr
		else
			'������ ���� (�⺻��)
			Num2Str = strChr & inum
		end if

    End Function


'// ���ڿ��� �߶� ���ϴ� ��ġ�� ���� ��ȯ //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)

    if UBound(buf)<pos then Exit function

    SplitValue = buf(pos)
end function


'// �Ķ���� ���� üũ �� Maxlen ���Ϸ� ������ Code, id ���� Param �� ��� //
function requestCheckVar(orgval,maxlen)
	requestCheckVar = trim(orgval)
	requestCheckVar = replace(requestCheckVar,"'","")
	requestCheckVar = replace(requestCheckVar,"--","")
	requestCheckVar = Left(requestCheckVar,maxlen)
end function


'// ���� �� Return �� like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function

'// ��� ����� �ڷΰ��� //
Sub Alert_return(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"history.back();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// ��� ����� â�ݱ� //
Sub Alert_close(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// ��� ����� ���� �������� �̵� //
Sub Alert_move(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.location='" & targetURL & "';" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// �������̷� ���ڿ� �ڸ��� //
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


'// �н����� ���⼺ �˻� �Լ�
Function chkPasswordComplex(uid,pwd)
	dim msg, i, sT, sN
	msg = ""

	'��й�ȣ ���� �˻�
	if len(pwd)<6 then
		msg = msg & "- ��й�ȣ�� �ּ� 6�ڸ��̻����� �Է����ּ���.\n"
	end if

	'���̵�� ���� �Ǵ� �����ϰ� �ִ°�?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- ���̵�� �����ϰų� ���̵� �����ϰ� �ִ� ��й�ȣ�Դϴ�.\n"
	end if

	'## ���⼺�� �����ϴ°�?
	'�������� 3�� ���� ����
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
			msg = msg & "- �������ڰ� 3�� �������� �������ϴ�.\n"
			exit for
		end if
	next
	'����/������ ����
	if chkWord(pwd,"[^-a-zA-Z]") or chkWord(pwd,"[^-0-9 ]") then
		msg = msg & "- ��й�ȣ�� �ݵ�� ���ĺ��� ���ڸ� �����ؼ� �������մϴ�.\n"
	end if

	'��� ��ȯ
	chkPasswordComplex = msg
end Function

'//���Խ� ���ڿ� �˻�
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp
    regEx.Pattern = patrn	' ������ ����.
    regEx.IgnoreCase = True	' ��/�ҹ��ڸ� �������� �ʵ��� .
    regEx.Global = True		' ��ü ���ڿ��� �˻��ϵ��� ����.
    SET Matches = regEx.Execute(str)
	if 0 < Matches.count then
		chkWord= false
	Else
		chkWord= true
	end if

	'pattern0 = "[^��-�R]"  '�ѱ۸�
	'pattern1 = "[^-0-9 ]"  '���ڸ�
	'pattern2 = "[^-a-zA-Z]"  '���
	'pattern3 = "[^-��-�Ra-zA-Z0-9/ ]" '���ڿ� ���� �ѱ۸�
	'pattern4 = "<[^>]*>"   '�±׸�
	'pattern5 = "[^-a-zA-Z0-9/ ]"    '���� ���ڸ�
End Function

'// ��ȭ��ȣ�� ��� �߰�
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


'''''==================  2009 �߰�

' response.write �Լ�
Function rw(ByVal str)
	response.write str & "<br>"
End Function

' Null�� �������� ġȯ
Function null2blank(ByVal v)
	If IsNull(v) Then
		null2blank = ""
	Else
		null2blank = v
	End If
End Function

'// ū����ǥ input �ڽ� value=""�� ����Ҷ� ġȯ
Function doubleQuote(ByVal v)
	If IsNull(v) Then
		doubleQuote = ""
	Else
		doubleQuote = Replace(v, """","&quot;")
	End If
End Function


' request ��ü �Լ�(�Ķ���͸�, ����Ʈ��)
Function req(ByVal param, ByVal value)
'	VarType Return ��
'	0 (����)
'	1 (��)
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
		If Not IsNumeric(tmpValue) Then	' ���ڰ� �ƴϸ�
			tmpValue = value
		End If
		tmpValue = CDbl(tmpValue)
	Else
		tmpValue = Trim(Request(param))
		If tmpValue = "" Then			' Request���� ������
			tmpValue = value
		End If
	End If
	req = tmpValue

End Function

Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage)

	'���� ����
	Dim intCurrentPage, strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ����
	intCurrentPage = Mid(strCurrentPage, InStr(strCurrentPage, "=")+1)		'���� ������ ��
	strCurrentPage = Left(strCurrentPage, InStr(strCurrentPage, "=")-1)		'������ ���� ������

	'���� ������ ��
	strCurrentPath = Request.ServerVariables("Script_Name")

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))

	'�� ���� & hidden �Ķ���� ����
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>"			'���� ������

	'�Ķ���� ����(��: �˻���)�� hidden �Ķ���ͷ� �����Ѵ�
	strParamName = ""
	For Each strParamName In Request.Form
		If strParamName <> strCurrentPage Then

			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""

	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then
			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"
		END IF
	Next

	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class=a><tr align='center'><td>"

	'���� ������ �̹��� ����
	If intStartBlock > 1 Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pageprev01.gif' border='0' style='cursor:hand' alt='���� " & intBlockPerPage & " ������'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();'>"
	Else
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif' border='0' >"
	End If

	Response.Write "</td><td>&nbsp;"

	'����¡ ���
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For

			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"

			If Int(intLoop) = Int(intCurrentPage) Then		'���� ������
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'�� �� ������
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If

		Next
	Else		'�� �������� ���� �Ҷ�
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'���� ������ �̹��� ����
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pagenext01.gif' border='0' style='cursor:hand' alt='���� " & intBlockPerPage & " ������'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'>"
	Else
	    Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif' border='0' >"
	End If

	Response.Write "</td></tr></form></table>"

End Sub



' ���,����,���� ��� �ؽ�Ʈ ����
Function getModeName(ByVal mode)
    Select Case mode
        Case "INS"	: getModeName = "���"
        Case "UPD"	: getModeName = "����"
        Case "DEL"	: getModeName = "����"
        Case "FIN"	: getModeName = "�Ϸ�"
        Case Else	: getModeName = "����"
    End Select
End Function

'// ������ȣ HTML�ڵ�� ġȯ //
Function ReplaceBracket(strng)
	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function


' ���Խ� �Լ�
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

'// ���ڿ����� ���ڸ� ���� ��ȯ
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

'//���ڿ��� Ư������ ����
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//�ø� �Լ�
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

'//�ø� �Լ�
function ceilValue(iValue)
 if iValue <>  round(iValue) then
  ceilValue = fix(iValue) + 1
 else
  ceilValue = iValue
 end if
end function

'// ��������ŭ ������ ���ڷ� �ٲ�)
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

'// ���뿡 ������ HTML�±װ� �ִ��� �˻� //
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

'// ��� ����� â�ݰ� ����â ���ε� -2011.02.23 �������߰� //
Sub Alert_closenreload(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.reload();"& vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// ��� ����� â�ݰ� ����â Ÿ���ּҷ� �̵� -2011.02.23 �������߰� //
Sub Alert_closenmove(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.href ='" & targetURL & "';" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'//�д����� �ð��������� ©�� ��ȯ	'/2011.03.31 �ѿ�� ����
function minutechagehour(v)
	dim tmpval , tmph , tmpm

	if v = "" or isnull(v) or v = 0 then
		minutechagehour = ""
	else
		tmph = int(v / 60)	'�ð�����
		tmpm = v - (tmph * 60)	'�д���

		if tmph <> 0 then tmpval = tmpval & tmph & "�ð� "
		if tmpm <> 0 then tmpval = tmpval & tmpm & "��"

		minutechagehour = tmpval
	end if
end function

'//���̳ʸ� ������ TEXT���·� ��ȯ
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'���� ������ Ÿ��
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' ��ȯ�� ������ ĳ���ͼ�
	 BinaryStream.CharSet = CharSet

	'��ȯ�� ������ ��ȯ
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function

'2012-09-19 ������ �״������� print_r2�� ASPȭ ������
Function r2(ByVal method)
	Dim i, pLength, msg, pValue, pJump, pLine, pTab
	pLine = "<br />"
	pTab = vbTab
	pTab = "    "
	msg = ""
	If IsArray(method) Then
		pLength = Ubound(method)
		For i=0 to pLength
			pValue = method(i)
			pJump = "&nbsp;&nbsp;" & pTab & "[" &  i & "]" & " => "
			If not IsArray(pValue) Then
				msg = msg & pJump & pValue & pLine
			Else
				msg = msg & pJump & r2(pValue)
			End If
		Next
		msg = "Array" & pLine & "(" & pLine & msg & ")" & pLine
		response.write "<span style='font-size:10pt; line-height:12px'>"&msg&"</span>"
	Else
		Dim key
		method = LCase(method)
		response.write "<table width=750 border=1 bordercolor='#cccccc' style='border-collapse:collapse;font:10pt'>" + vbcrlf
		response.write "<tr>" + vbcrlf
		response.write "	<td align='center' bgcolor='F1F1E5'>name</td>" + vbcrlf
		response.write "	<td align='center'>value</td>" + vbcrlf
		response.write "</tr>" + vbcrlf
		Select Case method
			Case "post" :
				For Each key in Request.Form
					response.write "<tr>" + vbcrlf
					response.write "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.Form(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.Form(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next

			Case "get" :
				For Each key in Request.QueryString
					response.write  "<tr>" + vbcrlf
					response.write  "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.QueryString(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.QueryString(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next
			Case "server" :
				For Each key in Request.ServerVariables
					response.write  "<tr>" + vbcrlf
					response.write  "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.ServerVariables(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.ServerVariables(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next
				response.write  method
			Case Else : response.write method
		End Select
	End If
END function

'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Sub serverupdate_underconstruction()
	Response.write "<html>"
	Response.write "<head><title>���� �������Դϴ�</title></head>"
	Response.write "<body>"
	Response.write "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
	Response.write "<tr>"
	Response.write "	<td align='center' valign='middle'><img src='http://fiximage.10x10.co.kr/10x10_underconstruction.jpg' width='673' border='0' ></td>"
	Response.write "</tr>"
	Response.write "</table>"
	Response.write "</body>"
	Response.write "</html>"
	response.End
End Sub
%>

