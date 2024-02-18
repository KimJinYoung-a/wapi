<%
Dim infoLoop, infoDivValue
%>
<script language='javascript'>
function checkComp(comp){
	if ((comp.name=="bestOrd")||(comp.name=="bestOrdMall")){
		if ((comp.name=="bestOrd")&&(comp.checked)){
			comp.form.bestOrdMall.checked=false;
		}
		if ((comp.name=="bestOrdMall")&&(comp.checked)){
			comp.form.bestOrd.checked=false;
		}
	}
}
</script>
텐바이텐 :
<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>&nbsp;
판매
<select name="sellyn" class="select">
	<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >전체
	<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >판매
	<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >품절
</select>&nbsp;
한정
<select name="limityn" class="select">
	<option value="">전체
	<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >한정
	<option value="N" <%= CHkIIF(limityn="N","selected","") %> >일반
</select>&nbsp;
세일
<select name="sailyn" class="select">
	<option value="">전체
	<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >세일Y
	<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >세일N
</select>&nbsp;
기준마진(<%= CMAXMARGIN %>%)
<select name="onlyValidMargin" class="select">
	<option value="">전체
	<option value="Y" <%= CHkIIF(onlyValidMargin="Y","selected","") %> >마진이상
	<option value="N" <%= CHkIIF(onlyValidMargin="N","selected","") %> >마진이하
</select>&nbsp;
주문제작
<select name="isMadeHand" class="select">
	<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
</select>&nbsp;
옵션
<select name="isOption" class="select">
	<option value="" <%= CHkIIF(isOption="","selected","") %> >전체
	<option value="optAll" <%= CHkIIF(isOption="optAll","selected","") %> >옵션전체
	<option value="optaddpricey" <%= CHkIIF(isOption="optaddpricey","selected","") %> >추가금액Y
	<option value="optaddpricen" <%= CHkIIF(isOption="optaddpricen","selected","") %> >추가금액N
	<option value="optN" <%= CHkIIF(isOption="optN","selected","") %> >단품
</select>&nbsp;
품목
<select name="infodiv" class="select">
	<option value="" <%= CHkIIF(infoDiv="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >입력
	<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >미입력
<%
	For infoLoop = 1 To 35
		If infoLoop < 10 Then
			infoDivValue = "0"&infoLoop
		Else
			infoDivValue = infoLoop
		End If
%>
	<option value="<%=infoDivValue%>" <%= CHkIIF(CStr(infodiv) = CStr(infoDivValue),"selected","") %> ><%= infoDivValue %>
	<% Next %>
	<option value="47" <%= CHkIIF(CStr(infodiv) = "47","selected","") %> >47
	<option value="48" <%= CHkIIF(CStr(infodiv) = "48","selected","") %> >48
</select>
<br>
제휴몰 &nbsp;&nbsp; :
<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>베스트순</b>&nbsp;
판매
<select name="extsellyn" class="select">
	<option value="" <%= CHkIIF(extsellyn="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >판매
	<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >품절
	<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >종료
	<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >종료제외
</select>&nbsp;
오류
<select name="failCntExists" class="select">
	<option value="" <%= CHkIIF(failCntExists="","selected","") %> >전체
	<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >등록수정오류1회이상
	<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >등록수정오류0회
</select>&nbsp;
