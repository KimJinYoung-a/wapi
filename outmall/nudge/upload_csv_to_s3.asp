<script language="javascript" runat="server">
function GMTNow(){return new Date().toGMTString()}
</script>
<%
on error resume next
Server.ScriptTimeOut = 9000

'Const AWS_BUCKETNAME = "jujubtown"
'Const AWS_ACCESSKEY = "AKIAIAHGR5BYU2DZCXHQ"
'Const AWS_SECRETKEY = "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"
Const AWS_BUCKETNAME = "10x10"
Const AWS_ACCESSKEY = "AKIAJCGAW54UTGTBT7ZA"
Const AWS_SECRETKEY = "nkfd7FwJgaMHAnrvjo5WwsCi4ap/baNZK0n7npRo"

LocalFile = Server.Mappath("/outmall/nudge/nudge_"&replace(FormatDateTime(Now(),2),"-","")&".csv")
''LocalFile = Server.Mappath("/outmall/nudge/nudge_"&replace(FormatDateTime(Now(),2),"-","")&"_1.csv") ''TEST

Set fs = Server.CreateObject("Scripting.FileSystemObject")
if not fs.FileExists(LocalFile) Then
	Set fs = nothing
	response.ContentType = "application/json"
	Response.write "{""result"":""fail"",""message"":""file not exist""}"
    response.end 
end if 

Dim sRemoteFilePath
sRemoteFilePath = "/"&replace(FormatDateTime(Now(),2),"-","")&"_data.csv" 'Remote Path, note that AWS paths (in fact they aren't real paths) are strictly case sensitive
''sRemoteFilePath = "/"&replace(FormatDateTime(Now(),2),"-","")&"_data_1.csv" ''TEST

Dim strNow
    strNow = GMTNow() ' GMT Date String

Dim StringToSign
    StringToSign = Replace("PUT\n\ntext/csv\n\nx-amz-date:" & strNow & "\n/"& AWS_BUCKETNAME & sRemoteFilePath, "\n", vbLf)

Dim Signature
    Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))

Dim Authorization
    Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature

Dim AWSBucketUrl
    AWSBucketUrl = "https://" & AWS_BUCKETNAME & ".s3.amazonaws.com"

''With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
With Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	.setTimeouts 15000, 15000, 900000, 900000                     ''2015/09/23
    .open "PUT", AWSBucketUrl & sRemoteFilePath, False
    .setRequestHeader "Authorization", Authorization
    .setRequestHeader "Content-Type", "text/csv"
    .setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
    .setRequestHeader "x-amz-date", strNow
    .send GetBytes(LocalFile) 'Get bytes of local file and send
    If .status = 200 Then ' successful
        'Response.Write "<a href="& AWSBucketUrl & sRemoteFilePath &" target=_blank>Uploaded File</a>"
        response.ContentType = "application/json"
        Response.write "{""result"":""success"",""fileurl"":""" & AWSBucketUrl & sRemoteFilePath & """}"
    Else ' an error ocurred, consider xml string of error details
        'Response.ContentType = "text/xml"
        'Response.Write .responseText
        response.ContentType = "application/json"
        Response.write "{""result"":""fail"",""message"":""" & Server.UrlEncode(.responseText) & """}"
    End If
    
    ''if fs.FileExists(LocalFile) Then	fs.DeleteFile(LocalFile)
    Set fs = nothing
End With

Function GetBytes(sPath)
    With Server.CreateObject("Adodb.Stream")
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile sPath
        .Position = 0
        GetBytes = .Read
        .Close
    End With
End Function

Function BytesToBase64(varBytes)
    With Server.CreateObject("MSXML2.DomDocument").CreateElement("b64")
        .dataType = "bin.base64"
        .nodeTypedValue = varBytes
        BytesToBase64 = .Text
    End With
End Function

Function HMACSHA1(varKey, varValue)
    With Server.CreateObject("System.Security.Cryptography.HMACSHA1")
        .Key = UTF8Bytes(varKey)
        HMACSHA1 = .ComputeHash_2(UTF8Bytes(varValue))
    End With
End Function

Function UTF8Bytes(varStr)
    With Server.CreateObject("System.Text.UTF8Encoding")
        UTF8Bytes = .GetBytes_4(varStr)
    End With
End Function
%>