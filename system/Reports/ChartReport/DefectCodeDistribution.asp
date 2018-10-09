<%@  language="VBSCRIPT" codepage="936" %>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/fusioncharts/FusionCharts.asp" -->
<!--#include virtual="/Functions/GetDefectCode.asp" -->
<!--#include virtual="/Language/Reports/ChartReport/Lan_DefectCodeDistribution.asp" -->
<%

fromdate=request("fromdate")
todate=request("todate")
defectcode=request("defectcode")
action=request("Action")
if fromdate="" then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
if todate="" then
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if

if action="1" then
	if len(defectcode)>0 then
		charttype="1"'colunm图
		SQL="SELECT TO_CHAR(S.CLOSE_TIME,'YYYY-MM-DD') AS CLOSE_TIME,SUM(d.defect_quantity) AS DEFECT_QUANTITY "
		SQL=SQL+"FROM JOB_DEFECTCODES D LEFT JOIN JOB_STATIONS S ON D.STATION_ID=S.STATION_ID WHERE 1=1 "	
		if fromdate<>"" then
			SQL=SQL+" and S.CLOSE_TIME>=to_date('" & fromdate &  " 00:00:00','yyyy-mm-dd hh24:mi:ss')"
		end if 
		if todate<>"" then
			SQL=SQL+" and S.CLOSE_TIME<=to_date('" & todate &  " 23:59:59','yyyy-mm-dd hh24:mi:ss')"
		end if
		if defectcode<>"" then
			'SQL=SQL+" and D.DEFECT_CODE_ID='"&defectcode&"'"
			SQL=SQL+" and D.DEFECT_CODE_ID IN（select NID from DEFECTCODE WHERE MOTHER_DEFECT_ID= '"&defectcode&"'）"
		end if	
		SQL=SQL+" GROUP BY TO_CHAR(S.CLOSE_TIME,'YYYY-MM-DD') ORDER BY TO_CHAR(S.CLOSE_TIME,'YYYY-MM-DD')"
	else
		charttype="2"'stack colunm图
		SQL="SELECT * FROM ( "
		SQL=sql+" SELECT  TO_CHAR(JS.CLOSE_TIME,'YYYY-MM-DD') AS CLOSE_TIME,D.DEFECT_NAME,SUM(Jd.defect_quantity) AS DEFECT_QUANTITY, "
		SQL=SQL+" row_number() over ( PARTITION BY TO_CHAR(JS.CLOSE_TIME,'YYYY-MM-DD') ORDER BY TO_CHAR(JS.CLOSE_TIME,'YYYY-MM-DD'),SUM(Jd.defect_quantity) desc) AS RN "
		SQL=SQL+" FROM JOB_DEFECTCODES JD LEFT JOIN JOB_STATIONS JS ON JD.STATION_ID=JS.STATION_ID left join DEFECTCODE D on D.NID=JD.DEFECT_CODE_ID WHERE 1=1" 
		if fromdate<>"" then
			SQL=SQL+" and JS.CLOSE_TIME>=to_date('" & fromdate &  " 00:00:00','yyyy-mm-dd hh24:mi:ss')"
		end if 
		if todate<>"" then
			SQL=SQL+" and JS.CLOSE_TIME<=to_date('" & todate &  " 23:59:59','yyyy-mm-dd hh24:mi:ss')"
		end if
		if defectcode<>"" then
			SQL=SQL+" and JD.DEFECT_CODE_ID='"&defectcode&"'"
		end if		
		SQL=SQL+" group by TO_CHAR(JS.CLOSE_TIME,'YYYY-MM-DD'),D.DEFECT_NAME) T WHERE T.RN<6"
	end if
	'response.Write(SQL)
	'response.End()
	rs.open SQL,conn,1,3	
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Components/sniffer.js" type="text/javascript"></script>
<script language="javascript" src="/Components/dynCalendar.js" type="text/javascript"></script>
<script src="/Components/fusioncharts/FusionCharts.js" type="text/javascript"></script>
<script src="/Components/fusioncharts/FusionChartsExportComponent.js" type="text/javascript"></script>
<script type="text/javascript">
	function GenerateReport() {
		form1.action = "DefectCodeDistribution.asp?Action=1"
		form1.submit();
	}		
</script>
</head>
<body onLoad="language();">
<form action="DefectCodeDistribution.asp" method="post" name="form1" target="_self" id="form1">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" style="z-index:100">
    <tr>
      <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_DefectCodeDistributionReport"></span></td>
    </tr>
    <tr align="center">
		<td width="80"><span id="inner_DefectCode"></span></td>
		<td width="80">
		  <select name="defectcode" id="defectcode">
			<option value="" ></option>
			<%=getDefectCode_New("OPTION",defectcode,"","","")%>
		  </select>
		</td>
		<td width="80"><span id="inner_Date"></span>&nbsp;<span id="inner_SearchFrom"></span> </td>
		<td width="100">
		  <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
		  <script language="JavaScript" type="text/javascript">
				function calendar1Callback(date, month, year) {
					document.all.fromdate.value = year + '-' + month + '-' + date
				}
				calendar1 = new dynCalendar('calendar1', 'calendar1Callback', document.all.fromdate.value);
			</script>
		</td>
		<td width="30"><span id="inner_SearchTo"></span> </td>
		<td width="100">
		  <input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
		  <script language="JavaScript" type="text/javascript">
				function calendar2Callback(date, month, year) {
					document.all.todate.value = year + '-' + month + '-' + date
				}
				calendar2 = new dynCalendar('calendar2', 'calendar2Callback', document.all.todate.value);
			</script>
		<td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle"style="cursor: hand" onclick="GenerateReport()" > 
		</td>        
    </tr>
  </table>
</form>  
<%if action="1" then%>  
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" style="z-index:0px">
    <tr>
      <td>
<%  if charttype="1" then
	strXML = "<graph caption='Defect Code Distribution Chart' xAxisName='DATE' yAxisName='DEFECT QUANTITY' decimalPrecision='0' formatNumberScale='0' slantLabels='1' exportEnabled='1' exportAtClient='1' exportHandler='fcExporter1' showPercentageValues='1'>"

	if not rs.eof then
		while not rs.eof    
			strXML=strXML&"<set name='"&rs("CLOSE_TIME")&"' value='"&rs("DEFECT_QUANTITY")&"' /> "
		  rs.movenext
		wend
	end if
	rs.close
	set rs=nothing
	
	strXML = strXML  & "</graph>"
	Call renderChart("/Components/fusioncharts/charts/Column3D.swf", "", strXML, "myNext", 1000, 400, false,true)

  elseif charttype="2" then

	strXML = "<graph "
	if session("language")="0" then
		strXML=strXML&"caption='Defect Code Distribution Chart' xAxisName='DATE' yAxisName='DEFECT QUANTITY' "
	else
		strXML=strXML&"caption='缺陷代码分布' xAxisName='日期' yAxisName='缺陷数量' "
	end if
	strXML=strXML&" decimalPrecision='0' formatNumberScale='0' slantLabels='1' exportEnabled='1' exportAtClient='1' exportHandler='fcExporter1' showPercentageValues='1'>"

	if not rs.eof then
		timecount=""
		defectcount=""		
		strXML=strXML&"<categories>"
		while not rs.eof
			if instr(timecount,rs("CLOSE_TIME"))=0 then
				timecount=timecount&rs("CLOSE_TIME")&";"
				strXML=strXML&"<category label='"&rs("CLOSE_TIME")&"' />"
			end if
			rs.movenext
		wend
		strXML=strXML&"</categories>"
	
		rs.MoveFirst
		arr=split(left(timecount,len(timecount)-1),";")
			
		while not rs.eof
			strXML=strXML&"<dataset seriesName='"&rs("DEFECT_NAME")&"'>"
			for i=0 to ubound(arr)
				if rs("CLOSE_TIME")=arr(i) then
					strXML=strXML&"<set value='"&rs("DEFECT_QUANTITY")&"' />"
				else
					strXML=strXML&"<set value='' />"	
				end if
			next		
			strXML=strXML&"</dataset>"
			rs.movenext
		wend
	end if
	rs.close
	set rs=nothing

	strXML = strXML  & "</graph>"
	response.Write(strXML)
	Call renderChart("/Components/fusioncharts/charts/StackedColumn3D.swf", "", strXML, "myNext", 1000, 400, false,true)
  end if
%>
        <div id="fcexpDiv" align="center">FusionCharts Export Handler Component</div>
        <script type="text/javascript">   
   var myExportComponent = new FusionChartsExportObject("fcExporter1", "/Components/FusionChart/charts/Column2D.swf");
	myExportComponent.Render("fcexpDiv");
</script>
      </td>
    </tr>
  </table>
 <%end if%>

</body>
</html>
<!--#include virtual="/Components/CopyRight.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
