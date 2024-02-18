<%

function FormatDate(ddate, formatstring)
        dim s
        if (formatstring = "0000.00.00") then
                s = CStr(year(ddate)) + "."
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "."
                else
        	        s = s + CStr(month(ddate)) + "."
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "0000-00-00") then
                s = CStr(year(ddate)) + "-"
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "-"
                else
        	        s = s + CStr(month(ddate)) + "-"
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "00000000") then
                s = CStr(year(ddate))
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate))
                else
        	        s = s + CStr(month(ddate))
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "0000.00") then
                s = CStr(year(ddate)) + "."
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate))
                else
        	        s = s + CStr(month(ddate))
        	end if
        elseif (formatstring = "0000.00.00-00:00:00") then
                s = CStr(year(ddate)) + "-"
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "-"
                else
        	        s = s + CStr(month(ddate)) + "-"
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        	if (Hour(ddate) < 10) then
        	        s = s + "0" + CStr(Hour(ddate))
                else
        	        s = s + CStr(Hour(ddate))
        	end if
        	if (Minute(ddate) < 10) then
        	        s = s + "0" + CStr(Minute(ddate))
                else
        	        s = s + CStr(Minute(ddate))
        	end if
        	if (Second(ddate) < 10) then
        	        s = s + "0" + CStr(Second(ddate))
                else
        	        s = s + CStr(Second(ddate))
        	end if
        elseif (formatstring = "0000.00.00 00:00:00") then
                s = CStr(year(ddate)) + "."
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "."
                else
        	        s = s + CStr(month(ddate)) + "."
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate)) + " "
                else
        	        s = s + CStr(day(ddate)) + " "
        	end if
        	if (Hour(ddate) < 10) then
        	        s = s + "0" + CStr(Hour(ddate)) + ":"
                else
        	        s = s + CStr(Hour(ddate)) + ":"
        	end if
        	if (Minute(ddate) < 10) then
        	        s = s + "0" + CStr(Minute(ddate)) + ":"
                else
        	        s = s + CStr(Minute(ddate)) + ":"
        	end if
        	if (Second(ddate) < 10) then
        	        s = s + "0" + CStr(Second(ddate))
                else
        	        s = s + CStr(Second(ddate))
        	end if
        elseif (formatstring = "0000/00/00") then
                s = CStr(year(ddate)) + "/"
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "/"
                else
        	        s = s + CStr(month(ddate)) + "/"
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "00/00/00") then
                s = Right(CStr(year(ddate)), 2) + "/"
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "/"
                else
        	        s = s + CStr(month(ddate)) + "/"
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "00.00.00") then
                s = Right(CStr(year(ddate)), 2) + "."
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "."
                else
        	        s = s + CStr(month(ddate)) + "."
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "00/00") then
                s = ""
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "/"
                else
        	        s = s + CStr(month(ddate)) + "/"
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        elseif (formatstring = "00.00") then
                s = ""
        	if (month(ddate) < 10) then
        	        s = s + "0" + CStr(month(ddate)) + "."
                else
        	        s = s + CStr(month(ddate)) + "."
        	end if
        	if (day(ddate) < 10) then
        	        s = s + "0" + CStr(day(ddate))
                else
        	        s = s + CStr(day(ddate))
        	end if
        else
                FormatDate = CStr(ddate)
        end if

        FormatDate = s
end function

%>