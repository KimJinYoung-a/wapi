-- Quick �˻� / ��� / --
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >��ϰ��� ��ǰ
<br><br>
-- Quick �˻� / ���� / --
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>������</font>��ǰ���� (MaxMagin : <%= CMAXMARGIN %>%) (Homeplus �Ǹ���)
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus ����<�ٹ����� �ǸŰ�</font>��ǰ����
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>���ݻ���</font>��ü����
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusYes10x10No" <%= ChkIIF(HomeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus�Ǹ���&�ٹ�����ǰ��</font>��ǰ����
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusNo10x10Yes" <%= ChkIIF(HomeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplusǰ��&�ٹ������ǸŰ���</font>(�Ǹ���,����>=10) ��ǰ����
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>�������</font>��ǰ���� (����������Ʈ�� ����)
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>ǰ��ó�����</font>��ǰ���� (���޸� �����Ե�)