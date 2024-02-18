<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  메일 오픈율 그래프
' History : 2007.08.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailopenresultclass/mailopenclass.asp"-->

<?xml version='1.0' encoding='EUC-KR' ?>
<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' showDivLineValue='0' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >
<categories>
	<category name='/' showName='1' showLine='1' />
	<category name='/event/eventmain.asp' showName='1' showLine='1' />
	<category name='/search/search_result.asp' showName='1' showLine='1' />
	<category name='/shoppingtoday/shoppingchance_saleitem.asp' showName='1' showLine='1' />
	<category name='/shopping/category_list.asp?cdl=10%26cdm=35' showName='1' showLine='1' />
</categories>
<dataset seriesName='오픈통수' color='006990' showValues='0' >
	<set value='406080' />
	<set value='307951' />
	<set value='289489' />
	<set value='128776' />
	<set value='71127' />
</dataset>
<dataset seriesName='성공발송통수' color='FF6F00' showValues='0' >
	<set value='174667' />
	<set value='93112' />
	<set value='44782' />
	<set value='12228' />
	<set value='11281' />
</dataset>
<dataset seriesName='실제발송통수' color='FF6F00' showValues='0' >
	<set value='174667' />
	<set value='93112' />
	<set value='44782' />
	<set value='12228' />
	<set value='11281' />
</dataset>
<trendLines></trendLines>
<styles>
	<definition>
		<style name='shadow215' type='shadow' angle='215' distance='3'/>
		<style name='shadow45' type='shadow' angle='45' distance='3'/>
	</definition>
	<application>
		<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />
		<apply toObject='DATAPLOTLINE' styles='shadow215' />
		<apply toObject='DATAPLOT' styles='shadow45' />
	</application>
</styles>
</chart>

<!-- #include virtual="/lib/db/dbclose.asp" -->