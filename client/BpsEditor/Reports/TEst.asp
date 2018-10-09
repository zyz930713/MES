<%server.ScriptTimeout=999999
response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/BpsEditor/Reports/Charts/FusionCharts.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>AMS Testing Data</title>
<script src="charts/FusionCharts.js" type="text/javascript"></script>
<script src="charts/FusionChartsExportComponent.js" type="text/javascript"></script>
<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
	background:#FFFFFF;
}
table tr td{
	text-align:center;
}
@media print {#btn{display: none;}}
</style>
</head>
<body style="font-size:14px;background-color:#FFFFFF;">
<%
Set conn_ams = Server.CreateObject("ADODB.Connection")
conn_ams.Open "driver={SQL Server};server=KEW-SQL2008;uid=KEWAMS;pwd=KEWAMS;database=TSDHV"
set rs = Server.CreateObject("adodb.recordset")
Getdata = request("Getdata")
Sdatetime = request("Sdatetime")
Edatetime = request("Edatetime")
MCode = request("MCode")

if MCode = "" then MCode = "26000171BJGAS01B"
%>
<center>
<form method="post" action="TEst.asp">
	<table width="962" border="1" cellspacing="0" cellpadding="0">
		<tr style="height:30px;line-height:30px;">
			<td align="center">
				Test Machine Name:<input type="" name="MCode" value="<%=MCode%>"/>
				Start Time:<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 07:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Sdatetime" value="<%=Sdatetime%>" size="20" />&nbsp;
				End Time:<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 19:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Edatetime" value="<%=Edatetime%>" size="20" />&nbsp;<input type="submit" name="Getdata" value="seacher" />
			</td>
		</tr>
	</table>
</form>
<center>
<table width="100%">

<%
sql = "EXEC [dbo].[SP_GetData] @TestAmsName = N'"&MCode&"',@StartTime = N'"&Sdatetime&"',@EndTime = N'"&Edatetime&"'"
'response.write sql & "</br>"
'response.End()

strXMLds=""
rs.open sql,conn_ams,1,1

if not rs.eof then
while not rs.eof  
strXMLds=strXMLds&"<set value='"&rs("QtyInput")&"' name='Input' /> "&"<set value='"&rs("QtyClassA")&"' name='Class A' /> "&"<set value='"&rs("QtyClassB")&"' name='Class B' /> "&"<set value='"&rs("QtyBad")&"' name='Defect' /> "
rs.movenext
wend
end if
rs.close
%>

    <tr>
      <td colspan="15">
<div id="dv2"></div>
<% 
strXML = "<graph caption='WDT Report' xAxisName='Test Data' yAxisName='Qty' decimalPrecision='0' formatNumberScale='0' slantLabels='1' exportEnabled='1' exportAtClient='1' exportHandler='fcExporter1' showPercentageValues='1' baseFontSize='20' BaseFontColor='000000' baseFontColor='0000CD'>"
	
strXML = strXML &strXMLds & "</graph>"
	Call renderChart("charts/Column3D.swf", "", strXML, "myNext", 500, 500, false,true)
%>


      </td>
            <td>
<%
sql = "EXEC [dbo].[SP_GetDataTest] @TestAmsName = N'"&MCode&"',@StartTime = N'"&Sdatetime&"',@EndTime = N'"&Edatetime&"'"
'response.write sql & "</br>"
'response.End()
strXMLMSL=""
strXMLds=""
rs.open sql,conn_ams,1,1

if not rs.eof then
	while not rs.eof
	strXMLds=strXMLds&" <set label='" & Replace(rs("TestName"),"C_","") & "' issliced='1' value='" & rs("BadQty") & "'/>"
	rs.movenext
	wend
end if
rs.close

strXMLMSL = "<chart caption='Defect Code'  baseFontSize='20' outCnvBaseFontColor='000000' baseFontColor='0000CD' palette='2' animation='1' formatnumberscale='1' decimals='2' pieslicedepth='60' startingangle='-50' showPercentValues='1' showPercentInToolTip='0' showValues='1' labelSepChar='~'>"&strXMLds& "</chart>"
	Call renderChart("charts/Pie3D.swf?ChartNoDataText=No data to display. &amp;#xA;", "", strXMLMSL, "myNext00", "800", "500", false,true)
%>

 	</td>
    </tr>
</table>
</body>
</html>
