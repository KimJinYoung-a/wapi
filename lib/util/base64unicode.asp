<%

'BASE64 유니코드버전(한글을 처리할 수 있다.)
'소스 사용설명은 맨 아래 적혀있다.

Dim ssBASE_64_CHARACTERS, sBASE_64_CHARACTERSansi

ssBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
sBASE_64_CHARACTERSansi	= strUnicode2Ansi(ssBASE_64_CHARACTERS)



Private Function strUnicodeLen(asContents)
	Dim asContents1	: asContents1 ="a"	& asContents
	Dim Len1		: Len1=Len(asContents1)
	Dim K		: K=0
	Dim I, Asc1

	For I=1 To Len1
		Asc1	= asc(mid(asContents1,I,1))
		IF Asc1 < 0 Then Asc1	= 65536 + Asc1
		IF Asc1 > 255 Then
			K	= K + 2
		ELSE
			K	= K + 1
		End IF
	Next

	strUnicodeLen = K - 1
End Function



Private Function strUnicode2Ansi(asContents)
	Dim Len1			: Len1 = Len(asContents)
	Dim I, VarCHAR, VarASC, VarHEX, VarLOW, VarHIGH

	strUnicode2Ansi	= ""

	For I = 1 to Len1
		VarCHAR	= Mid(asContents,I,1)
		VarASC	= Asc(VarCHAR)
		IF VarASC < 0 Then VarASC = VarASC + 65536
		IF VarASC > 255 Then
			VarHEX		= Hex(VarASC)
			VarLOW		= Left(VarHEX,2)
			VarHIGH		= Right(VarHEX,2)
			strUnicode2Ansi	= strUnicode2Ansi & ChrB("&H" & VarLOW ) & ChrB("&H" & VarHIGH )
		Else
			strUnicode2Ansi	= strUnicode2Ansi & ChrB(VarASC)
		End IF
	Next
End Function



Private Function strAnsi2Unicode(asContents)
	Dim Len1			: Len1		= LenB(asContents)
	Dim VarCHAR, VarASC, I

	strAnsi2Unicode	= ""

	IF Len1=0 Then	Exit Function

	For I=1 To Len1
		VarCHAR	= MidB(asContents,I,1)
		VarASC	= AscB(VarCHAR)
		IF VarASC > 127 Then
			strAnsi2Unicode	= strAnsi2Unicode & Chr(AscW(MidB(asContents,I+1,1) & VarCHAR))
			I		= I + 1
		Else
			strAnsi2Unicode	= strAnsi2Unicode & Chr(VarASC)
		End IF
	Next

End function



Private Function PrivateBase64encode(asContents)
	Dim lnPosition
	Dim lsResult
	Dim Char1
	Dim Char2
	Dim Char3
	Dim Char4
	Dim Byte1
	Dim Byte2
	Dim Byte3
	Dim SaveBits1
	Dim SaveBits2
	Dim lsGroupBinary
	Dim lsGroup64
	Dim M3, M4, Len1, Len2

	Len1			=LenB(asContents)

	IF Len1 < 1 Then
		PrivateBase64encode	= ""
		Exit Function
	End IF

	M3=Len1 Mod 3

	IF M3 > 0 Then asContents = asContents & String(3 - M3, ChrB(0))

	IF m3 > 0 Then
		Len1	= Len1 + (3 - M3)
		Len2	= Len1 - 3
	Else
		Len2	= Len1
	End IF

	lsResult	= ""

	For lnPosition = 1 To Len2 Step 3
		lsGroup64	= ""
		lsGroupBinary	= MidB(asContents, lnPosition, 3)

		Byte1		= AscB(MidB(lsGroupBinary, 1, 1))	: SaveBits1	= Byte1 And 3
		Byte2		= AscB(MidB(lsGroupBinary, 2, 1))	: SaveBits2	= Byte2 And 15
		Byte3		= AscB(MidB(lsGroupBinary, 3, 1))

		Char1		= MidB(sBASE_64_CHARACTERSansi, ((Byte1 And 252) \ 4) + 1,1)
		Char2		= MidB(sBASE_64_CHARACTERSansi, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
		Char3		= MidB(sBASE_64_CHARACTERSansi, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
		Char4		= MidB(sBASE_64_CHARACTERSansi, (Byte3 And 63) + 1, 1)
		lsGroup64	= Char1 & Char2 & Char3 & Char4

		lsResult		= lsResult & lsGroup64
	Next

	IF M3 > 0 Then
		lsGroup64	= ""
		lsGroupBinary	= MidB(asContents, Len2 + 1, 3)

		Byte1		= AscB(MidB(lsGroupBinary, 1, 1))	: SaveBits1	= Byte1 And 3
		Byte2		= AscB(MidB(lsGroupBinary, 2, 1))	: SaveBits2	= Byte2 And 15
		Byte3		= AscB(MidB(lsGroupBinary, 3, 1))

		Char1		= MidB(sBASE_64_CHARACTERSansi, ((Byte1 And 252) \ 4) + 1,1)
		Char2		= MidB(sBASE_64_CHARACTERSansi, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
		Char3		= MidB(sBASE_64_CHARACTERSansi, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)

		IF M3=1 Then
			lsGroup64	= Char1 & Char2 & ChrB(61) & ChrB(61)
		Else
			lsGroup64	= Char1 & Char2 & Char3 & ChrB(61)
		End IF

		lsResult		= lsResult & lsGroup64
	End IF

	PrivateBase64encode = lsResult
End Function

Private Function PrivateBase64decode(asContents)
	Dim lsResult
	Dim lnPosition
	Dim lsGroup64, lsGroupBinary
	Dim Char1, Char2, Char3, Char4
	Dim Byte1, Byte2, Byte3
	Dim M4, Len1, Len2

	Len1	= LenB(asContents)
	M4	= Len1 Mod 4

	IF Len1 < 1 Or M4 > 0 Then
		PrivateBase64decode = ""
		Exit Function
	End IF

	IF MidB(asContents, Len1, 1) = ChrB(61) Then	M4 = 3
	IF MidB(asContents, Len1-1, 1) = ChrB(61) Then	M4 = 2

	IF M4 = 0 Then
		Len2	= Len1
	Else
		Len2	= Len1 - 4
	End IF

	For lnPosition = 1 To Len2 Step 4
		lsGroupBinary	= ""
		lsGroup64	= MidB(asContents, lnPosition, 4)

		Char1		= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 1, 1)) - 1
		Char2		= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 2, 1)) - 1
		Char3		= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 3, 1)) - 1
		Char4		= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 4, 1)) - 1

		Byte1		= ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
		Byte2		= lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
		Byte3		= ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
		lsGroupBinary	= Byte1 & Byte2 & Byte3

		lsResult		= lsResult & lsGroupBinary
	Next

	IF M4 > 0 Then
		lsGroupBinary	= ""
		lsGroup64	= MidB(asContents, Len2 + 1, M4) & ChrB(65)
		IF M4=2 Then
			lsGroup64	= lsGroup64 & chrB(65)
		End IF
		Char1	= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 1, 1)) - 1
		Char2	= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 2, 1)) - 1
		Char3	= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 3, 1)) - 1
		Char4	= InStrB(sBASE_64_CHARACTERSansi, MidB(lsGroup64, 4, 1)) - 1

		Byte1	= ChrB(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
		Byte2	= lsGroupBinary & ChrB(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
		Byte3	= ChrB((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))

		IF M4=2 Then
			lsGroupBinary	= Byte1
		elseIF M4=3 Then
			lsGroupBinary	= Byte1 & Byte2
		end IF

		lsResult			= lsResult & lsGroupBinary
	End IF

	PrivateBase64decode			= lsResult
End Function

Public Function Base64EncodeUnicode(asContents)

	Base64EncodeUnicode = strAnsi2Unicode(PrivateBase64encode(strUnicode2Ansi(asContents)))

End Function

Public Function Base64decodeUnicode(asContents)

	Base64decodeUnicode = strAnsi2Unicode(PrivateBase64decode(strUnicode2Ansi(asContents)))

End Function


'################################## 사용 샘플 ##################################
'Dim ORIGNvalue	: ORIGNvalue	= "하완어환오라ㅓ고ㅓ엉ㅁㄹ22ㅁㅁ@@ㅇㅁㄹ;ㄴㅇ"
'Dim EncodeA	: EncodeA		= Base64EncodeUnicode(ORIGNvalue)
'Dim DecodeA	: DecodeA		= Base64decodeUnicode(EncodeA)

'response.write "------------------------------------------------------------------------<BR>"
'response.write "lsResult : " & ORIGNvalue			& "<BR>"
'response.write "lsResult : " & EncodeA				& "<BR>"
'response.write "lsResult : " & DecodeA				& "<BR>"
'response.write "------------------------------------------------------------------------<BR>"

%>
