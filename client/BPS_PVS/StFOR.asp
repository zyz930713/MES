<!--#include file="conn.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="300" charset="utf-8" />
<title>StationFOR</title>
<script type="text/javascript" language="javascript" src="../include/AnyChart/js/AnyChart.js"></script>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<center>
<%
' Response.buffer=true '开启缓存
' Response.Write "<div id='loading' style='width:960px;height:480px;background:#FFFFFF url(images/loading.gif) center no-repeat;'>" '先写入提示
' Response.Write "</div>"
' Response.flush()
' Response.clear()


set Rs = Server.CreateObject("adodb.recordset")
Sql = "EXEC [dbo].[GetStationFOR] @LineName = N'"&LineName&"', @StartTime = N'"&Sdatetime&"', @EndTime = N'"&Edatetime&"',@SortFOR = N'"&SortFOR&"'"
'response.write sql
Rs.open sql,Conn,1,1
if not rs.eof then
	dim xAxisName,PrBad1Xml,PrBad2Xml
	xAxisName = ""
	GoodXmlContent = ""
	PrBad1Xml = ""
	PrBad2Xml = ""
	do while not rs.eof
		StID = rs("StID")
		StName = rs("StName")
		LR = rs("PrPosition")
		PrBad = rs("PrBad")
		PrAll = rs("PrAll")
		LRC = Replace(LR,"0","L")
		LRC = Replace(LRC,"1","R")
		
		xAxisName = xAxisName & "<category name='" & StName & " - " & LRC & "  " & PrBad & "/" & PrAll & "' />"
		PrFOR1 = rs("PrFOR1")
		if PrFOR1 = 0 then
			PrBad1Xml = PrBad1Xml & "<set value=''/>"
		else
			PrBad1Xml = PrBad1Xml & "<set value='" & PrFOR1 & "' link='P-NewWin,width=600,height=500,toolbar=no,scrollbars=no,resizable=no-StFORDeatil.asp?LineName="&LineName&"&StID="&StID&"&LR="&LR&"&Edatetime="&Edatetime&"' />"
		end if
		
		PrFOR2 = rs("PrFOR2")
		PrBad2 = rs("PrBad2")
		if PrFOR2 = 0 then
			PrBad2Xml = PrBad2Xml & "<set value=''/>"
		else
			PrBad2Xml = PrBad2Xml & "<set value='" & PrFOR2 & "' link='P-NewWin,width=600,height=500,toolbar=no,scrollbars=no,resizable=no-StFORDeatil.asp?LineName="&LineName&"&StID="&StID&"&LR="&LR&"&Edatetime="&Edatetime&"'/>"
		end if
		rs.movenext
	loop

	ChartXml = ""
	ChartXml = ChartXml & "<graph Caption='Process FOR' plotFillAlpha='100' baseFontSize='10' numberSuffix='%25' decimalPrecision='2' formatNumberScale='0' rotateNames='0' showValues='1'>"
	ChartXml = ChartXml & "<categories>"
	ChartXml = ChartXml & xAxisName
	ChartXml = ChartXml & "</categories>"
	
	ChartXml = ChartXml & "<dataset color='1A7AFF' >"
	ChartXml = ChartXml & PrBad1Xml
	ChartXml = ChartXml & "</dataset>"
	
	ChartXml = ChartXml & "<dataset color='Fad35e' >"
	ChartXml = ChartXml & PrBad2Xml
	ChartXml = ChartXml & "</dataset>"
	
	ChartXml = ChartXml & GoodXmlContent
	ChartXml = ChartXml & "</graph>"
	
	response.write "<script>"
	response.write "var oLoading=document.getElementById('loading');"
	response.write "var parent=oLoading.parentNode;"
	response.write "parent.removeChild(oLoading);"
	response.write "</script>"
Call renderChart("../include/charts/MSBar2D.swf?ChartNoDataText=No data to display.&amp;#xA;", "", ChartXml, "myNext01", "1150", "1000", false,true)
end if
rs.close
set rs = nothing
%>
</center>
</body>
</html>
