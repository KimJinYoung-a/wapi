<%
Class CTnGridSubItem
    public FRowNum
    public FColValue
    public FColName
    
    Private Sub Class_Initialize()
		
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

Class CTnGridData
    public FData
    public FCnt
    
    public FPageSize
    public FCurrPage
    public FTotalCount
    
    public FTotalSum
    public FAvgSum
    
    public Fdummi1
    public Fdummi2
    public Fdummi3
    public Fdummi4
    public Fdummi5
    
    public sub AddData(iRowNum, iColValue, iColName)
        dim iData
        set iData = New CTnGridSubItem
        
        iData.FRowNum   = iRowNum
        iData.FColValue = iColValue
        iData.FColName  = iColName
        
        if (FCnt>=UBound(FData)) then
            redim preserve FData(FCnt)
        end if
        
        set FData(FCnt) = iData
        FCnt = FCnt + 1
    end sub

    Private Sub Class_Initialize()
		redim FData(0)
		FCnt = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

''복호화
function DecMsg(byval msg)
    DecMsg = msg
    if IsNULL(msg) or (msg="") then Exit function
	DecMsg = URLDecode(MyEnDecrypt(URLDecode(msg),chr(4)&chr(6)&chr(3)&chr(5)))
end function

''암호화
function EncMsg(byval msg)
    EncMsg = msg
    if IsNULL(msg) or (msg="") then Exit function
    
	EncMsg = server.UrlEncode(MyEnDecrypt(server.UrlEncode(msg),chr(4)&chr(6)&chr(3)&chr(5)))
end function

''Function MakeObjVal(byRef iObj, iRowNum, iColName, iVal)
    
''end function

Function URLDecode(expr)
    Dim strSource, strTemp, strResult, strchr
    Dim lngPos, AddNum, IFKor
    strSource = Replace(expr, "+", " ")
    For lngPos = 1 To Len(strSource)
        AddNum = 2
        strTemp = Mid(strSource, lngPos, 1)
        If strTemp = "%" Then
            If lngPos + AddNum < Len(strSource) + 1 Then
                strchr = CInt("&H" & Mid(strSource, lngPos + 1, AddNum))
                If strchr > 130 Then
                    AddNum = 5
                    IFKor = Mid(strSource, lngPos + 1, AddNum)
                    IFKor = Replace(IFKor, "%", "")
                    strchr = CInt("&H" & IFKor )
                End If
                strResult = strResult & Chr(strchr)
                lngPos = lngPos + AddNum
            End If
        Else
            strResult = strResult & strTemp
        End If
    Next
    URLDecode = strResult
End Function


Public Function MyEnDecrypt(ENCRSTRING,strPWD )
	Dim I
	Dim strNewPwd
	Dim J , K, c
	
	MyEnDecrypt = ENCRSTRING
	if IsNULL(ENCRSTRING) or (ENCRSTRING="") then Exit Function
	
    J = Len(strPWD)
    For I = 1 To Len(ENCRSTRING)
        If K = J Then K = 1 Else K = K + 1
        c = Chr(Asc(Mid(ENCRSTRING, I, 1)) Xor Asc(Mid(strPWD, K, 1)))
        strNewPwd = strNewPwd + c
    Next
    MyEnDecrypt = strNewPwd
End Function


function returnXMLResult(byval cmd ,byval enc, byval result ,byval resultMsg, byval rows, byval ArrType)
    dim xmlMsg, intCols, intRows, i, j
    xmlMsg = "<?xml version='1.0' encoding='EUC-KR' ?>" + VbCrlf
    xmlMsg = xmlMsg + "<root>" + VbCrlf
    xmlMsg = xmlMsg + "<cmd>" + cmd + "</cmd>" + VbCrlf
    xmlMsg = xmlMsg + "<result>" + result + "</result>" + VbCrlf
    xmlMsg = xmlMsg + "<resultmsg>" + resultMsg + "</resultmsg>" + VbCrlf
    xmlMsg = xmlMsg + "<enc>" + enc + "</enc>" + VbCrlf
    if IsArray(rows) then
        if (ArrType=2) then
            ''2차원 Array
            intCols = UBound(rows,1)
            intRows = UBound(rows,2)
            for i=0 to intRows
                xmlMsg = xmlMsg + "<rowData>" + VbCrlf
                for j=0 to intCols
                    if (enc="Y") then
                        xmlMsg = xmlMsg + "    <col" & j & "><![CDATA[" & (EncMsg(rows(j,i))) & "]]></col" & j & ">" + VbCrlf
                    else
                        xmlMsg = xmlMsg + "    <col" & j & "><![CDATA[" & rows(j,i) & "]]></col" & j & ">" + VbCrlf
                    end if
                next
                xmlMsg = xmlMsg + "</rowData>" + VbCrlf
            Next
        elseif (ArrType=1) then
            ''1차원 Array
            intCols = UBound(rows,1)
            xmlMsg = xmlMsg + "<rowData>" + VbCrlf
            for j=0 to intCols
                if (enc="Y") then
                    xmlMsg = xmlMsg + "    <col" & j & "><![CDATA[" & (EncMsg(rows(j))) & "]]></col" & j & ">" + VbCrlf
                else
                    xmlMsg = xmlMsg + "    <col" & j & "><![CDATA[" & rows(j) & "]]></col" & j & ">" + VbCrlf
                end if
            next
            xmlMsg = xmlMsg + "</rowData>" + VbCrlf
        end if
    elseif (rows<>"") then
        xmlMsg = xmlMsg + "<rowData>" + VbCrlf
        if (enc="Y") then
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & (EncMsg(rows)) & "]]></col0>" + VbCrlf
        else
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & rows & "]]></col0>" + VbCrlf
        end if
        xmlMsg = xmlMsg + "</rowData>" + VbCrlf
    end if
    
    xmlMsg = xmlMsg + "</root>" + VbCrlf
    
    returnXMLResult = xmlMsg
end function


function returnXMLResultWithColName(byval cmd ,byval enc, byval result ,byval resultMsg, byval rows, byval ArrType, byval ArrColName)
    dim xmlMsg, intCols, intRows, i, j
    xmlMsg = "<?xml version='1.0' encoding='EUC-KR' ?>" + VbCrlf
    xmlMsg = xmlMsg + "<root>" + VbCrlf
    xmlMsg = xmlMsg + "<cmd>" + cmd + "</cmd>" + VbCrlf
    xmlMsg = xmlMsg + "<result>" + result + "</result>" + VbCrlf
    xmlMsg = xmlMsg + "<resultmsg>" + resultMsg + "</resultmsg>" + VbCrlf
    xmlMsg = xmlMsg + "<enc>" + enc + "</enc>" + VbCrlf
    if IsArray(rows) then
        if (ArrType=2) then
            ''2차원 Array
            intCols = UBound(rows,1)
            intRows = UBound(rows,2)
            
            for i=0 to intRows
                xmlMsg = xmlMsg + "<rowData>" + VbCrlf
                for j=0 to intCols
                    if (enc="Y") then
                        xmlMsg = xmlMsg + "    <" & ArrColName(j) & "><![CDATA[" & (EncMsg(rows(j,i))) & "]]></" & ArrColName(j) & ">" + VbCrlf
                    else
                        xmlMsg = xmlMsg + "    <" & ArrColName(j) & "><![CDATA[" & rows(j,i) & "]]></" & ArrColName(j) & ">" + VbCrlf
                    end if
                next
                xmlMsg = xmlMsg + "</rowData>" + VbCrlf
            Next
        elseif (ArrType=1) then
            ''1차원 Array
            intCols = UBound(rows,1)
            xmlMsg = xmlMsg + "<rowData>" + VbCrlf
            for j=0 to intCols
                if (enc="Y") then
                    xmlMsg = xmlMsg + "    <" & ArrColName(j) & "><![CDATA[" & (EncMsg(rows(j))) & "]]></" & ArrColName(j) & ">" + VbCrlf
                else
                    xmlMsg = xmlMsg + "    <" & ArrColName(j) & "><![CDATA[" & rows(j) & "]]></" & ArrColName(j) & ">" + VbCrlf
                end if
            next
            xmlMsg = xmlMsg + "</rowData>" + VbCrlf
        end if
    elseif (rows<>"") then
        xmlMsg = xmlMsg + "<rowData>" + VbCrlf
        if (enc="Y") then
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & (EncMsg(rows)) & "]]></col0>" + VbCrlf
        else
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & rows & "]]></col0>" + VbCrlf
        end if
        xmlMsg = xmlMsg + "</rowData>" + VbCrlf
    end if
    
    xmlMsg = xmlMsg + "</root>" + VbCrlf
    
    returnXMLResultWithColName = xmlMsg
end function



function returnXMLResultObjArr(byval cmd ,byval enc, byval result ,byval resultMsg, byval iObj)
    dim xmlMsg, cnt, i, preRowNum
    preRowNum = -1
    
    xmlMsg = "<?xml version='1.0' encoding='EUC-KR' ?>" + VbCrlf
    xmlMsg = xmlMsg + "<root>" + VbCrlf
    xmlMsg = xmlMsg + "<cmd>" + cmd + "</cmd>" + VbCrlf
    xmlMsg = xmlMsg + "<result>" + result + "</result>" + VbCrlf
    xmlMsg = xmlMsg + "<resultmsg>" + resultMsg + "</resultmsg>" + VbCrlf
    xmlMsg = xmlMsg + "<enc>" + enc + "</enc>" + VbCrlf
    if IsObject(iObj) then
        ''1차원 Array
        cnt = iObj.FCnt
        
        xmlMsg = xmlMsg + "<pagesize>" & iObj.FPageSize & "</pagesize>" + VbCrlf
        xmlMsg = xmlMsg + "<currpage>" & iObj.FCurrPage & "</currpage>" + VbCrlf
        xmlMsg = xmlMsg + "<totalcnt>" & iObj.FTotalCount & "</totalcnt>" + VbCrlf
        xmlMsg = xmlMsg + "<totalsum>" & iObj.FTotalSum & "</totalsum>" + VbCrlf
        xmlMsg = xmlMsg + "<avgsum>" & iObj.FAvgSum & "</avgsum>" + VbCrlf
        xmlMsg = xmlMsg + "<dummi1>" & iObj.Fdummi1 & "</dummi1>" + VbCrlf
        xmlMsg = xmlMsg + "<dummi2>" & iObj.Fdummi2 & "</dummi2>" + VbCrlf
        xmlMsg = xmlMsg + "<dummi3>" & iObj.Fdummi3 & "</dummi3>" + VbCrlf
        xmlMsg = xmlMsg + "<dummi4>" & iObj.Fdummi4 & "</dummi4>" + VbCrlf
        xmlMsg = xmlMsg + "<dummi5>" & iObj.Fdummi5 & "</dummi5>" + VbCrlf
        
        for i=0 to cnt-1
            if Not IsEmpty(iObj.FData(i)) then
                if (preRowNum<>iObj.FData(i).FRowNum) then
                    if (preRowNum<>-1) then
                        xmlMsg = xmlMsg + "</rowData>" + VbCrlf
                    end if
                    xmlMsg = xmlMsg + "<rowData>" + VbCrlf
                end if
                
                if (enc="Y") then
                    xmlMsg = xmlMsg + "    <" & iObj.FData(i).FColName & "><![CDATA[" & (EncMsg(iObj.FData(i).FColValue)) & "]]></" & iObj.FData(i).FColName & ">" + VbCrlf
                else
                    xmlMsg = xmlMsg + "    <" & iObj.FData(i).FColName & "><![CDATA[" & iObj.FData(i).FColValue & "]]></" & iObj.FData(i).FColName & ">" + VbCrlf
                end if
                preRowNum = iObj.FData(i).FRowNum
                
                if (i=cnt-1) then xmlMsg = xmlMsg + "</rowData>" + VbCrlf
            end if
            
        next
        
    else
        xmlMsg = xmlMsg + "<rowData>" + VbCrlf
        if (enc="Y") then
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & (EncMsg(iObj)) & "]]></col0>" + VbCrlf
        else
            xmlMsg = xmlMsg + "    <col0><![CDATA[" & iObj & "]]></col0>" + VbCrlf
        end if
        xmlMsg = xmlMsg + "</rowData>" + VbCrlf
    end if
    
    xmlMsg = xmlMsg + "</root>" + VbCrlf
    
    returnXMLResultObjArr = xmlMsg
end function
%>