<%

function db2xml(checkvalue)
        dim v
        v = checkvalue
        if Isnull(v) then Exit function

        On Error resume Next
        v = replace(v, "&quot;", "'")
        v = Replace(v, "", "<br>")
        v = Replace(v, "\0x5C", "\")
        v = Replace(v, "\0x22", "'")
        v = Replace(v, "\0x25", "'")
        v = Replace(v, "\0x27", "%")
        v = Replace(v, "\0x2F", "/")
        v = Replace(v, "\0x5F", "_")
        db2xml = v
end function

function html2xml(checkvalue)
        dim v
        v = checkvalue
        if Isnull(v) then Exit function

        v = replace(v, "&", "&amp;")
        v = replace(v, "<", "&lt;")
        v = replace(v, ">", "&gt;")
        html2xml = v
end function

%>
