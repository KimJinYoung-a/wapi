-- Quick 검색 / 등록 / --
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqREG"  >등록가능 상품
<br><br>
-- Quick 검색 / 수정 / --
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="showminusmagin" <%= ChkIIF(showminusmagin="on","checked","") %> ><font color=red>역마진</font>상품보기 (MaxMagin : <%= CMAXMARGIN %>%) (Homeplus 판매중)
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="expensive10x10" <%= ChkIIF(expensive10x10="on","checked","") %> ><font color=red>Homeplus 가격<텐바이텐 판매가</font>상품보기
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="diffPrc" <%= ChkIIF(diffPrc="on","checked","") %> ><font color=red>가격상이</font>전체보기
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusYes10x10No" <%= ChkIIF(HomeplusYes10x10No="on","checked","") %> ><font color=red>Homeplus판매중&텐바이텐품절</font>상품보기
&nbsp;
<input onClick="checkQuickClick(this)" type="checkbox" name="HomeplusNo10x10Yes" <%= ChkIIF(HomeplusNo10x10Yes="on","checked","") %> ><font color=red>Homeplus품절&텐바이텐판매가능</font>(판매중,한정>=10) 상품보기
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqEdit" <%= ChkIIF(reqEdit="on","checked","") %> ><font color=red>수정요망</font>상품보기 (최종업데이트일 기준)
<br>
<input onClick="checkQuickClick(this)" type="checkbox" name="reqExpire" <%= ChkIIF(reqExpire="on","checked","") %> ><font color=red>품절처리요망</font>상품보기 (제휴몰 사용안함등)