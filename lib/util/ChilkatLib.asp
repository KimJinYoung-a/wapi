<%
function UniToHanbyChilkat(oStr)
    dim cks, buf
    buf = oStr
    set cks = Server.CreateObject("Chilkat.String")
    cks.Str = buf
    cks.HtmlEntityDecode
    buf = cks.Str
    set cks=Nothing
    UniToHanbyChilkat = buf
end function
%>