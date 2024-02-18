<%
Function StrToHex(byVal Str)
    Dim Length
    Dim Max
    Dim strHex
    Max = Len(Str)
    For Length = 1 To Max
        strHex = strHex & Right("0" & Hex(Asc(Mid(Str, Length, 1))), 2)
    Next
    StrToHex = strHex
End function

Function HexToStr(byVal strHex)
    Dim Length
    Dim Max
    Dim Str
    Max = Len(strHex)
    For Length = 1 To Max Step 2
        Str = Str & Chr("&h" & Mid(strHex, Length, 2))
    Next
    HexToStr = Str
End function

Function StrToHexB(byVal Str)
    Dim Length
    Dim Max
    Dim strHex
    Max = Len(Str)
    For Length = 1 To Max
   '' response.write Asc(Mid(Str, Length, 1)) & ","
        strHex = strHex & Right("000" & Hex(Asc(Mid(Str, Length, 1))), 4)
    Next
    StrToHexB = strHex
End function

Function HexToStrB(byVal strHex)
    Dim Length
    Dim Max
    Dim Str
    Max = Len(strHex)
    For Length = 1 To Max Step 4
        Str = Str & Chr("&h" & Mid(strHex, Length, 4))
    Next
    HexToStrB = Str
End function



function tenEnc(byVal sPlain)
	tenEnc = tenEncW(sPlain)
    'sPlain = StrToHexB(sPlain)
    'dim lLength, lCount, sTemp, buf
    'dim mul, mul1, mul2, mul3
    'lLength = Len(sPlain)
    'For lCount = 1 To lLength
    '    mul1 = lCount mod 9
    '    mul2 = lCount mod 4
    '    mul3 = lCount mod 30
    '    if (lCount mod 2)=0 then mul3 = 30-mul3

    '    mul  = mul1-mul2+mul3  '' Max Add 38 == F(70) + 38 = l(108),  Min -3 = 44(,)
    '    buf  = Asc(Mid(sPlain,lCount,1))+(mul)
    '    sTemp = sTemp & CHR(buf)
    'Next
    'tenEnc = sTemp
end function

function tenDec(byVal sPlainEnc)
	tenDec = tenDecW(sPlainEnc)
    'dim lLength, lCount, sTemp, buf
    'dim mul, mul1, mul2, mul3
    'lLength = Len(sPlainEnc)
    'For lCount = 1 To lLength
    '    mul1 = lCount mod 9
    '    mul2 = lCount mod 4
    '    mul3 = lCount mod 30
    '    if (lCount mod 2)=0 then mul3 = 30-mul3
    '    mul  = mul1-mul2+mul3
    '    buf = Asc(Mid(sPlainEnc, lCount, 1))-mul
    '    sTemp = sTemp & Chr(buf)
    'Next

    'tenDec = HexToStrB(sTemp)
end function


''Require base64unicode===================================
function TBTEncrypt(byVal sPlainEnc)
    TBTEncrypt = Base64encodeUnicode(tenEncW(sPlainEnc))
end function

function TBTDecrypt(byVal sEncryptedEnc)
    TBTDecrypt = tenDecW(Base64decodeUnicode(sEncryptedEnc))
end function
''========================================================

'''==========วัฑÛฟ๋==========
Function StrToHexW(byVal Str)
    Dim Length
    Dim Max
    Dim strHex
    Max = Len(Str)
    For Length = 1 To Max
        'response.write AscW(Mid(Str, Length, 1)) & "," & "<br>"
        strHex = strHex & Right("000" & Hex(AscW(Mid(Str, Length, 1))), 4)
    Next
    StrToHexW = strHex
End function

Function HexToStrW(byVal strHex)
    Dim Length
    Dim Max
    Dim Str
    Max = Len(strHex)
    For Length = 1 To Max Step 4
        Str = Str & ChrW("&h" & Mid(strHex, Length, 4))
    Next
    HexToStrW = Str
End function

function tenEncW(byVal sPlain)
    sPlain = StrToHexW(sPlain)
    dim lLength, lCount, sTemp, buf
    dim mul, mul1, mul2, mul3
    lLength = Len(sPlain)
    For lCount = 1 To lLength
        mul1 = lCount mod 9
        mul2 = lCount mod 4
        mul3 = lCount mod 30
        if (lCount mod 2)=0 then mul3 = 30-mul3

        mul  = mul1-mul2+mul3  '' Max Add 38 == F(70) + 38 = l(108),  Min -3 = 44(,)
        buf  = AscW(Mid(sPlain,lCount,1))+(mul)
        sTemp = sTemp & CHRW(buf)
    Next
    tenEncW = sTemp
end function

function tenDecW(byVal sPlainEnc)
    dim lLength, lCount, sTemp, buf
    dim mul, mul1, mul2, mul3
    lLength = Len(sPlainEnc)
    For lCount = 1 To lLength
        mul1 = lCount mod 9
        mul2 = lCount mod 4
        mul3 = lCount mod 30
        if (lCount mod 2)=0 then mul3 = 30-mul3
        mul  = mul1-mul2+mul3
        buf = AscW(Mid(sPlainEnc, lCount, 1))-mul
        sTemp = sTemp & ChrW(buf)
    Next

    tenDecW = HexToStrW(sTemp)
end function

%>