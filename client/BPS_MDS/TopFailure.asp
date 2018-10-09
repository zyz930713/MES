<!--#include file="conn.asp"-->
<!--#include virtual="/include/function.asp" -->
<!--#include virtual="/include/Charts/FusionCharts.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="300" charset="utf-8" />
<title>列表</title>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<script language="JavaScript" type="text/javascript" src="../include/Charts/FusionCharts.js"></script>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<center>
<form name="form1" method="post" action="Output.asp" >
	<center>
		<input type="submit" name="refresh" value="刷新" />
	</center>
</form>
<%
	response.write "待开发....."
	response.end
	NowDate = FormatTime(now(),2)
	TemTime2 = cstr(NowDate) & " 06:30:00"
	TemTime1 = dateadd("h",-12,cdate(TemTime2))
	TemTime3 = dateadd("h",12,cdate(TemTime2))
	TemTime4 = dateadd("h",12,cdate(TemTime3))

	Nt = now()
	if datediff("s",TemTime1,Nt) > 0 and datediff("s",Nt,TemTime2) > 0 then 
		'response.write "A"
		Sdatetime = TemTime1
		Edatetime = TemTime2
	elseif datediff("s",TemTime2,Nt) > 0 and datediff("s",Nt,TemTime3) > 0 then 
		'response.write "B"
		Sdatetime = TemTime2
		Edatetime = TemTime3
	elseif datediff("s",TemTime3,Nt) > 0 and datediff("s",Nt,TemTime4) > 0 then 
		'response.write "C"
		Sdatetime = TemTime3
		Edatetime = TemTime4
	end if
	
	'写Cookies默认缓存数据'
	LineName = request.Cookies("mds")("LineName")
	if LineName = "" then
		response.Cookies("mds")("LineName") = "AF01"
	end if
	'写Cookies默认缓存数据'

	dim ProductName,LineName,rs,sql
	xAxisNameStr = ""
	PrBadXml = ""
	PrGoodXml = ""
	PrFORXml = ""
	
	set RsL = Server.CreateObject("adodb.recordset")
	Sql2 = "EXEC GetOutputFOR @StartTime = N'"&Sdatetime&"', @EndTime = N'"&Edatetime&"'"
	'response.write Sql2 & "</BR>"
	RsL.open Sql2,conn,1,1
	if not RsL.eof then
		do while not RsL.eof		
			LineName = RsL("LineName")
			LineBad = CSNG(RsL("LineBad"))
			LineGood = CSNG(RsL("LineGood"))
			LineTotal = CSNG(RsL("LineTotal"))
			'LineGood = LineTotal - LineBad
			if LineTotal = 0 then
				LineFOR = 0
			else
				LineFOR = LineBad / LineTotal * 100
			end if
			xAxisNameStr = xAxisNameStr & "<category name='"&LineName&"' />"
			PrBadXml = PrBadXml & "<set value='"&LineBad&"' />"
			PrGoodXml = PrGoodXml & "<set value='"&LineGood&"' />"
			PrAllXml = PrAllXml & "<set value='"&LineTotal&"' />"
			PrFORXml = PrFORXml & "<set value='"&LineFOR&"' />"
			RsL.movenext
		loop
		ChartXml = ""
		ChartXml = ChartXml & "<graph caption='" & Sdatetime & " to " & Edatetime & "' PYAxisName='Quantity (PCS)' SYAxisName='FOR(％)' xAxisName='Line' showValues='1' decimalPrecision='2' formatNumberScale='0' bgcolor='f3f3f3' bgAlpha='70' showColumnShadow='1' divlinecolor='c5c5c5' divLineAlpha='60' showAlternateHGridColor='1' alternateHGridColor='f8f8f8' alternateHGridAlpha='60' useRoundEdges='0'>"
		ChartXml = ChartXml & "<categories>"
		ChartXml = ChartXml & xAxisNameStr
		ChartXml = ChartXml & "</categories>"
		ChartXml = ChartXml & "<dataset seriesName='Bad' parentYAxis='P' color='FF8080' >"
		ChartXml = ChartXml & PrBadXml
		ChartXml = ChartXml & "</dataset>"
		ChartXml = ChartXml & "<dataset seriesName='Input' parentYAxis='P' color='8080FF' >"
		ChartXml = ChartXml & PrAllXml
		ChartXml = ChartXml & "</dataset>"
		ChartXml = ChartXml & "<dataset seriesName='Good' parentYAxis='P' color='80FF80' >"
		ChartXml = ChartXml & PrGoodXml
		ChartXml = ChartXml & "</dataset>"
		ChartXml = ChartXml & "<dataset seriesName='FOR(％)' parentYAxis='S' color='FF0000' numberSuffix='%25' anchorSides='20' anchorRadius='6' anchorBorderColor='009900' >"
		ChartXml = ChartXml & PrFORXml
		ChartXml = ChartXml & "</dataset>"
		ChartXml = ChartXml & "</graph>"
		'response.write ChartXml
		Call renderChart("../include/charts/MSCombiDY2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", ChartXml, "myNext01", "1024", "540", false,true)
	else
		response.write "没有找到数据"		
	end if
	RsL.close
	set RsL = nothing
%>
</center>
</body>
</html>
