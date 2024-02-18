<%@Language=VBScript%>
<% Option Explicit %>
<%
	'날짜 변수명 받기
	Dim chrDateName, chrForm, vInDomain
	chrDateName = Request.QueryString("DN")
	chrForm = Request.QueryString("FN")
	vInDomain = Request.QueryString("in_domain")
%>
<!-- #include virtual="/lib/popheader.asp"-->
<script>
<!--
resizeTo(260,280);

/******** 환경설정 부분 *******************************************************/

var giStartYear = 1973;
var giEndYear = 2073;
var giCellWidth = 16;
var gMonths = new Array("01","02","03","04","05","06","07","08","09","10","11","12");
var gcOtherDay = "gray";
var gcToggle = "yellow";
var gcBG = "white";
var gcTodayBG = "#dddddf";
var gcFrame = "orange";
var gcHead = "white";
var gcWeekend = "red";
var gcWeekend1 = "blue";
var gcWorkday = "black";
var gcCalBG = "lightblue";
//-->

var gcTemp = gcBG;
var gdCurDate = new Date();
var giYear = gdCurDate.getFullYear();
var giMonth = gdCurDate.getMonth()+1;
var giDay = gdCurDate.getDate();
var tbMonSelect, tbYearSelect;
var gCellSet = new Array;

<% If vInDomain = "o" Then %>
document.domain = "10x10.co.kr";
<% End If %>

function fSetDate(iYear, iMonth, iDay){
	if(iDay < 10)
	iDay = '0'+iDay;

     <% if chrForm="" then %>
     eval("window.opener.document.all.<%=chrDateName%>").value = iYear+"-"+gMonths[iMonth-1]+"-"+iDay;
     <% else %>
     eval("window.opener.document.<%=chrForm%>.<%=chrDateName%>").value = iYear+"-"+gMonths[iMonth-1]+"-"+iDay;
     <% end if %>

  self.close();
}

function fSetSelected(aCell){
  var iOffset = 0;
  var iYear = parseInt(tbSelYear.value);
  var iMonth = parseInt(tbSelMonth.value);

  aCell.bgColor = gcBG;
  with (aCell.firstChild){
  	var iDate = parseInt(innerHTML);
  	if (style.color==gcOtherDay)
		iOffset = (id<10)?-1:1;
	iMonth += iOffset;
	if (iMonth<1) {
		iYear--;
		iMonth = 12;
	}else if (iMonth>12){
		iYear++;
		iMonth = 1;
	}
  }

  fSetDate(iYear, iMonth, iDate);
}

function fBuildCal(iYear, iMonth) {
  var aMonth=new Array();
  for(i=1;i<7;i++)
  	aMonth[i]=new Array(i);

  var dCalDate=new Date(iYear, iMonth-1, 1);
  var iDayOfFirst=dCalDate.getDay();
  var iDaysInMonth=new Date(iYear, iMonth, 0).getDate();
  var iOffsetLast=new Date(iYear, iMonth-1, 0).getDate()-iDayOfFirst+1;
  var iDate = 1;
  var iNext = 1;

  for (d = 0; d < 7; d++)
	aMonth[1][d] = (d<iDayOfFirst)?-(iOffsetLast+d):iDate++;
  for (w = 2; w < 7; w++)
  	for (d = 0; d < 7; d++)
		aMonth[w][d] = (iDate<=iDaysInMonth)?iDate++:-(iNext++);
  return aMonth;
}

function fDrawCal(iCellWidth) {


 var WeekDay = new Array("일","월","화","수","목","금","토");
 var styleTD = " width='"+iCellWidth+"' ";


with (document) {
write("<tr>")
write("<td><table width='220' border='0' cellpadding='0' cellspacing='1' bgcolor='#BEBEBE' class='a'>")
write("<tr align='center' bgcolor='#B9D4E6'>")
write("<td height='19'>일</td>")
write("<td height='19'>월</td>")
write("<td height='19'>화</td>")
write("<td height='19'>수</td>")
write("<td height='19'>목</td>")
write("<td height='19'>금</td>")
write("<td height='19'>토</td>")
write("</tr>")

  	for (w = 1; w < 7; w++) {
		write("<tr align='center' bgcolor='#FFFFFF'>");
		for (d = 0; d < 7; d++) {
			write("<td height='19' "+styleTD+" class='red' onMouseOver='gcTemp=this.bgColor;this.bgColor=gcToggle;this.bgColor=gcToggle' onMouseOut='this.bgColor=gcTemp;this.bgColor=gcTemp' onclick='fSetSelected(this)'>");
			write("<A href='#null' onfocus='this.blur();'>00</A></td>")
		}
		write("</tr>");
	}
  }
}

function fUpdateCal(iYear, iMonth) {
  myMonth = fBuildCal(iYear, iMonth);
  var i = 0;
  var iDate = 0;
  for (w = 0; w < 6; w++)
	for (d = 0; d < 7; d++)
		with (gCellSet[(7*w)+d]) {
			id = i++;
			if (myMonth[w+1][d]<0) {
				style.color =gcBG;
				innerHTML = -myMonth[w+1][d];
				iDate = 0;
			}else{

			   if(d==0)
				style.color = gcWeekend;
			   else if(d==6)
			    style.color = gcWeekend1;
			   else if(d!=6)
			    style.color = gcWorkday;
				innerHTML = myMonth[w+1][d];
				iDate++;
			}

			parentNode.bgColor = ((iYear==giYear)&&(iMonth==giMonth)&&(iDate==giDay))?gcTodayBG:gcBG;
			parentNode.bgColor = parentNode.bgColor;
		}
}

function fSetYearMon(iYear, iMon){
  tbSelMonth.options[iMon-1].selected = true;
  if (iYear>giEndYear) iYear=giEndYear;
  if (iYear<giStartYear) iYear=giStartYear;
  tbSelYear.options[iYear-giStartYear].selected = true;
  fUpdateCal(iYear, iMon);
}

function fPrevMonth(){
	var iMon = tbSelMonth.value;
	var iYear = tbSelYear.value;

	if (--iMon<1) {
		iMon = 12;
		iYear--;
	}

	fSetYearMon(iYear, iMon);
}

function fNextMonth(){
	var iMon = tbSelMonth.value;
	var iYear = tbSelYear.value;

	if (++iMon>12) {
		iMon = 1;
		iYear++;
	}

	fSetYearMon(iYear, iMon);
}

function fPrevYear(){
	var iMon = tbSelMonth.value;
	var iYear = tbSelYear.value;

	iYear--;

	fSetYearMon(iYear, iMon);
}

function fNextYear(){
	var iMon = tbSelMonth.value;
	var iYear = tbSelYear.value;

	iYear++;

	fSetYearMon(iYear, iMon);
}


with (document) {

write("<table width='230' border='0' cellspacing='0' cellpadding='0' bgcolor='#FFFFFF' >")
write("<tr>")
write("<td align='center' background='/WebManager/Image/Calendar/calendar_02.gif'><table width='220' border='0' cellspacing='0' cellpadding='0'>")
write("<tr>")
write("<td height='20' align='center'><table border='0' cellspacing='0' cellpadding='0' class='a'>")
write("<tr>")
//write("<td width='20'><img src='/images/icon_arrow_left.gif'  border='0' onclick='fPrevMonth()'></td>"))
write("<td width='45'><input type=button class='icon' value='<<' onclick='fPrevYear()'> <input type=button class='icon'  value='<' onclick='fPrevMonth()'></td>")
write("<td class='blue'><SELECT id='tbYearSelect' class='HeadBox' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)' Victor='Won'>");
 for(i=giStartYear;i<=giEndYear;i++)
	write("<OPTION value='"+i+"'>"+i+"</OPTION>");
	write("</SELECT><strong>년</strong>&nbsp;&nbsp;");
write("<select id='tbMonSelect' class='HeadBox' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)' Victor='Won'><strong>");
for (i=0; i<12; i++)
	write("<option value='"+(i+1)+"'>"+gMonths[i]+"</option>");
	write("</SELECT><strong>월</strong></TD>") ;

//write("<td width='20' align='right'><img src='/images/icon_arrow_right.gif'  border='0' onclick='fNextMonth()'></td>")
write("<td width='45' align='right'><input type=button class='icon'  value='>' onclick='fNextMonth()'> <input type=button class='icon'  value='>>' onclick='fNextYear()'></td>")
write("</tr>")
write("</table></td>")
write("</tr>")
write("<tr>")
write("<td height='3'> </td>")
write("</tr>")

 tbSelMonth = getElementById("tbMonSelect");
 tbSelYear = getElementById("tbYearSelect");
  fDrawCal(giCellWidth);
	gCellSet = getElementsByTagName("A") ;
  fSetYearMon(giYear, giMonth);
write("</table></td>");
write("</tr>");
write("</table></td>");
write("</tr>");
write("<tr><td height='15'></td></tr>");
write("</table>");

}
// -->
</script>
</body>
</html>