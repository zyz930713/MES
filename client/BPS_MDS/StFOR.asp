<!--#include file="conn.asp"-->
<!--#include virtual="/include/function.asp" -->
<!--#include virtual="/include/Charts/FusionCharts.asp" -->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="300" charset="utf-8" />
<title>StationFOR</title>
<script language="JavaScript" type="text/javascript" src="../include/Charts/FusionCharts.js"></script>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<center>
<%
	LineName = request("LineName")
	SortFOR = request("SortFOR")
	if SortFOR = "on" then SortSelect = "checked"
	'取Cookies缓存数据'
	if LineName = "" then
		LineName = request.Cookies("mds")("LineName")
	else
		response.Cookies("mds")("LineName") = LineName
	end if
	
	if Itime = "" then
		Itime = request.Cookies("mds")("Itime")
	else
		response.Cookies("mds")("Itime") = Itime
	end if
	'取Cookies缓存数据'
	
	LinePC = request("LinePC")
	Itime = request("Itime")
	if Itime = "" then Itime = 1
	Edatetime = now()
	Sdatetime = dateadd("h",(0-Itime),Edatetime)
	'response.write LineName & " - " & Sdatetime & " - " & Edatetime
%>
<form name="form1" method="post" action="StFOR.asp" >
线号：<select name="LineName">
				<option value="AF01" <%if LineName = "AF01" then response.write "selected"%> >AF01</option>
				<option value="AF02" <%if LineName = "AF02" then response.write "selected"%> >AF02</option>
				<option value="AF03" <%if LineName = "AF03" then response.write "selected"%> >AF03</option>
				<option value="AF04" <%if LineName = "AF04" then response.write "selected"%> >AF04</option>
				<option value="AF05" <%if LineName = "AF05" then response.write "selected"%> >AF05</option>
			</select>
　间隔：<select name="Itime">
        <option value="1" <%if Itime = "1" then response.write "selected"%>> 1 </option>
        <option value="2" <%if Itime = "2" then response.write "selected"%>> 2 </option>
        <option value="3" <%if Itime = "3" then response.write "selected"%>> 3 </option>
        <option value="4" <%if Itime = "4" then response.write "selected"%>> 4 </option>
        <option value="5" <%if Itime = "5" then response.write "selected"%>> 5 </option>
        <option value="6" <%if Itime = "6" then response.write "selected"%>> 6 </option>
        <option value="7" <%if Itime = "7" then response.write "selected"%>> 7 </option>
        <option value="8" <%if Itime = "8" then response.write "selected"%>> 8 </option>
        <option value="9" <%if Itime = "9" then response.write "selected"%>> 9 </option>
        <option value="10" <%if Itime = "10" then response.write "selected"%>> 10 </option>
        <option value="11" <%if Itime = "11" then response.write "selected"%>> 11 </option>
        <option value="12" <%if Itime = "12" then response.write "selected"%>> 12 </option>
    </select>小时
　<input type="checkbox" name="SortFOR" <%=SortSelect%>/>排序显示
　<input type="submit" name="refresh" value="刷新" />
</form>
<%
Response.buffer=true '开启缓存
Response.Write "<div id='loading' style='width:960px;height:480px;background:#FFFFFF url(images/loading.gif) center no-repeat;'>" '先写入提示
'Response.Write "<img src='images/loading.gif' />"
Response.Write "</div>"
Response.flush()
Response.clear()

dim LineName,rs,sql
ProductName = "Marigold"
set Rs = Server.CreateObject("adodb.recordset")
Sql = "EXEC [dbo].[GetStationFOR] @productName = N'Marigold',@LineName = N'"&LineName&"', @StartTime = N'"&Sdatetime&"', @EndTime = N'"&Edatetime&"',@SortFOR = N'"&SortFOR&"'"
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
			PrBad1Xml = PrBad1Xml & "<set value='" & PrFOR1 & "' link='P-NewWin,width=600,height=500,toolbar=no,scrollbars=no,resizable=no-StFORDeatil.asp?productName="&productName&"&LineName="&LineName&"&StID="&StID&"&LR="&LR&"&Edatetime="&Edatetime&"' />"
		end if
		
		PrFOR2 = rs("PrFOR2")
		PrBad2 = rs("PrBad2")
		if PrFOR2 = 0 then
			PrBad2Xml = PrBad2Xml & "<set value=''/>"
		else
			PrBad2Xml = PrBad2Xml & "<set value='" & PrFOR2 & "' link='P-NewWin,width=600,height=500,toolbar=no,scrollbars=no,resizable=no-StFORDeatil.asp?productName="&productName&"&LineName="&LineName&"&StID="&StID&"&LR="&LR&"&Edatetime="&Edatetime&"'/>"
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
