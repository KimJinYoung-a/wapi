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
�ٹ����� :
<input type="checkbox" name="bestOrd" <%= ChkIIF(bestOrd="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
�Ǹ�
<select name="sellyn" class="select">
	<option value="A" <%= CHkIIF(sellyn="A","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(sellyn="Y","selected","") %> >�Ǹ�
	<option value="N" <%= CHkIIF(sellyn="N","selected","") %> >ǰ��
</select>&nbsp;
����
<select name="limityn" class="select">
	<option value="">��ü
	<option value="Y" <%= CHkIIF(limityn="Y","selected","") %> >����
	<option value="N" <%= CHkIIF(limityn="N","selected","") %> >�Ϲ�
</select>&nbsp;
����
<select name="sailyn" class="select">
	<option value="">��ü
	<option value="Y" <%= CHkIIF(sailyn="Y","selected","") %> >����Y
	<option value="N" <%= CHkIIF(sailyn="N","selected","") %> >����N
</select>&nbsp;
���ظ���(<%= CMAXMARGIN %>%)
<select name="onlyValidMargin" class="select">
	<option value="">��ü
	<option value="Y" <%= CHkIIF(onlyValidMargin="Y","selected","") %> >�����̻�
	<option value="N" <%= CHkIIF(onlyValidMargin="N","selected","") %> >��������
</select>&nbsp;
�ֹ�����
<select name="isMadeHand" class="select">
	<option value="" <%= CHkIIF(isMadeHand="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(isMadeHand="Y","selected","") %> >Y
	<option value="N" <%= CHkIIF(isMadeHand="N","selected","") %> >N
</select>&nbsp;
�ɼ�
<select name="isOption" class="select">
	<option value="" <%= CHkIIF(isOption="","selected","") %> >��ü
	<option value="optAll" <%= CHkIIF(isOption="optAll","selected","") %> >�ɼ���ü
	<option value="optaddpricey" <%= CHkIIF(isOption="optaddpricey","selected","") %> >�߰��ݾ�Y
	<option value="optaddpricen" <%= CHkIIF(isOption="optaddpricen","selected","") %> >�߰��ݾ�N
	<option value="optN" <%= CHkIIF(isOption="optN","selected","") %> >��ǰ
</select>&nbsp;
ǰ��
<select name="infodiv" class="select">
	<option value="" <%= CHkIIF(infoDiv="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(infoDiv="Y","selected","") %> >�Է�
	<option value="N" <%= CHkIIF(infoDiv="N","selected","") %> >���Է�
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
���޸� &nbsp;&nbsp; :
<input type="checkbox" name="bestOrdMall" <%= ChkIIF(bestOrdMall="on","checked","") %> onClick="checkComp(this)"><b>����Ʈ��</b>&nbsp;
�Ǹ�
<select name="extsellyn" class="select">
	<option value="" <%= CHkIIF(extsellyn="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(extsellyn="Y","selected","") %> >�Ǹ�
	<option value="N" <%= CHkIIF(extsellyn="N","selected","") %> >ǰ��
	<option value="X" <%= CHkIIF(extsellyn="X","selected","") %> >����
	<option value="YN" <%= CHkIIF(extsellyn="YN","selected","") %> >��������
</select>&nbsp;
����
<select name="failCntExists" class="select">
	<option value="" <%= CHkIIF(failCntExists="","selected","") %> >��ü
	<option value="Y" <%= CHkIIF(failCntExists="Y","selected","") %> >��ϼ�������1ȸ�̻�
	<option value="N" <%= CHkIIF(failCntExists="N","selected","") %> >��ϼ�������0ȸ
</select>&nbsp;
