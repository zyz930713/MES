<!--#include virtual="/BpsEditor/include/function.asp" -->
<!--#include virtual="/BpsEditor/Conn_open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="../../styles/basic.css" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<title>Weekly Reporting</title>
<style type="text/css">
body{
font:normal 16px/16px "Arial";
}
table{
border-collapse:collapse;
border-spacing:0;
border:1px solid black;
}

a{
color:#fff;
text-decoration:none;
}
</style>
</head>
<body style="background-color:#FFFFFF;">
<%
ProductName = request("ProductName")
Getdata = request("Getdata")

Ayear = request("Ayear")
AWeek = request("AWeek")

if Ayear = "" then Ayear = Year(now())
if Aweek = "" then AWeek = GetWeekNo(date())

'response.write Aweek
'response.end
%>
<br/>
<center>
<form method="post" action="WeeklyReportHv.asp">
产品：<select name="ProductName" id="TmProduct"></select>
		<select name="LineName" id="TmLine" style="display:none" ></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
		<script type="text/javascript">
			MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
		</script>&nbsp;
　年：<input type="text" name="Ayear" value="<%=Ayear%>" size="3"/>　周：<input type="text" name="AWeek" value="<%=AWeek%>" size="3"/>&nbsp;
<input type="submit" name="Getdata" value="刷新" />
</form>
<center>

<%
if Getdata = "刷新" or Getdata = "ok" then
	set rs = server.createobject("adodb.recordset")
	sql = "SELECT top 1 [Year],[WeekNr],[Month],[MonthInt],[Startdate],[Enddate] FROM [dbo].[YearWeek] where [Year] = '"&Ayear&"' and WeekNr = '"&AWeek&"'"
	rs.open sql,ConnSql,1,1
	if not rs.eof then
		Startdate = rs("Startdate")
		EndDate = rs("Enddate")
	end if

	'制作Table表头
	TableStr = "<table width='960' border='1' cellpadding='0' cellspacing='0'>"
	TableStr = TableStr & "<tr bgcolor='#B9E9FF'><th>Prod<br/>Line</th><th>DAY<br/>Night</th><th>Prod<br/>Time</th><th>Down<br/>Time</th><th>Target<br/>CC</th>"
	if ProductName = "RA" or ProductName = "Slim" then 
		TableStr = TableStr & "<th>Weekly<br/>Quantity<br/>Cell 5</th><th>Output<br/>Cell 5</th><th>FOR<br/>Cell 5</th><th>FOR<br/>Cell 4</th>"
	else
		TableStr = TableStr & "<th>Weekly<br/>Quantity<br/>Cell 3</th><th>Output<br/>Cell 3</th>"
	end if
	TableStr = TableStr & "<th>FOR<br/>Cell 3</th><th>FOR<br/>Cell 2</th><th>FOR<br/>Cell 1</th><th>FOR<br/>Transfer</th><th>OEE</th><th>TEEP</th></tr><br/>"
	
	%>
	<table width="960" border="0" cellspacing="0" cellpadding="0" bgcolor="#7BB1DB">
	  <tr>
		<td width="20%">&nbsp;</td>
		<td width="60%" rowspan="2"><p style="font-size:20px;color:#FFF;text-align:center;"><strong><%=ProductName%></strong></p></td>
		<td width="20%" colspan="3">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;">
			<a href="WeeklyReportHv.asp?Getdata=ok&ProductName=<%=ProductName%>&Ayear=<%=Ayear%>&Aweek=<%=Aweek-1%>"> << </a>
			　<%=Ayear%>年<%=Aweek%>周　
			<a href="WeeklyReportHv.asp?Getdata=ok&ProductName=<%=ProductName%>&Ayear=<%=Ayear%>&Aweek=<%=Aweek+1%>"> >> </a>
		</td>
		<td width="3%" rowspan="2">&nbsp;</td>
		<td width="14%" rowspan="2"><img src="../Images/Logo.jpg" width="124" height="52" align="middle"></td>
		<td width="3%" rowspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;">&nbsp;</td>
		<td><p style="font-size:20px;color:#FFF;text-align:center;">Daily Report</p></td>
	  </tr>
	  <tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="3">&nbsp;</td>
	  </tr>
	</table>
<%
	'第一次循环
	set RsF = Server.CreateObject("adodb.recordset")
	set RsS = Server.CreateObject("adodb.recordset")
	set RsT = Server.CreateObject("adodb.recordset")
	LoopStr = ""
	LineAllTargetCC = 0
	SqlF = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME FROM report_hv_data RHD where '"&Startdate&"' < = adate and adate < '"&EndDate&"' and ProductName like '"&ProductName&"' group by PRODUCTNAME,LINENAME order by PRODUCTNAME,LINENAME"
	'response.write SqlF & "<br/><br/>"
	RsF.open SqlF,conn,1,1
	if not RsF.eof then
		do while not RsF.eof
			ProductNameF = RsF("PRODUCTNAME")
			LineNameF = RsF("LINENAME")
			'response.write ProductNameF & " - " & LineNameF & "<br/>"
			
			'第二次循环开始
			SqlS = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME,DNAME FROM report_hv_data RHD where '"&Startdate&"' < = adate and adate < '"&EndDate&"' and ProductName like '"&ProductName&"' and LineName = '"&LineNameF&"' group by PRODUCTNAME,LINENAME,DNAME order by PRODUCTNAME,LINENAME"
			RsS.open SqlS,conn,1,1
			if not RsS.eof then
				do while not RsS.eof
					DNameS = RsS("DNAME")
					'response.write ProductNameF & " - " & LineNameF & " - " & DNameS & "<br/>"
					
					'第三次循环开始
					SqlT = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME,DNAME,sum(NVL(OPENTIME,0))/60 OPENTIME,sum(NVL(STOPTIME,0))/60 STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B FROM report_hv_data RHD"
					SqlT = SqlT & " where '"&Startdate&"' < = adate and adate < '"&EndDate&"' and ProductName like '"&ProductName&"' and LineName = '"&LineNameF&"' and Dname = '"&DNameS&"' GROUP BY PRODUCTNAME,LINENAME,DNAME"
					'response.write SqlT & "<br/>"
					LoopStr = LoopStr & LoopWrite(SqlT)
					'第三次循环结束
				RsS.movenext
				loop
			'计算SubTotal
				SqlSubTotal = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME,'SubTotal' DName,sum(NVL(OPENTIME,0))/60 OPENTIME,sum(NVL(STOPTIME,0))/60 STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B FROM report_hv_data RHD"
				SqlSubTotal = SqlSubTotal & " where '"&Startdate&"' < = adate and adate < '"&EndDate&"' and ProductName like '"&ProductName&"' and LineName = '"&LineNameF&"' GROUP BY PRODUCTNAME,LINENAME"
				'response.write SqlSubTotal & "<br/>"
				'response.write "计算SubTotal<br/><br/><hr/>"
				LoopStr = LoopStr & LoopWrite(SqlSubTotal)
				
			'计算SubTotal
			end if
			RsS.close
			'第二次循环结束
			
		RsF.movenext
		loop
	'计算AllTotal
		SqlAllTotal = "SELECT count(Productname) CountDay,PRODUCTNAME,'All' LINENAME,'All Lines' DName,sum(NVL(OPENTIME,0))/60 OPENTIME,sum(NVL(STOPTIME,0))/60 STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B FROM report_hv_data RHD"
		SqlAllTotal = SqlAllTotal & " where '"&Startdate&"' < = adate and adate < '"&EndDate&"' and ProductName like '"&ProductName&"' GROUP BY PRODUCTNAME"
		'response.write SqlAllTotal & "<br/>"
		'response.write "计算AllTotal<br/>"
		LoopStr = LoopStr & LoopWrite(SqlAllTotal)
	'计算AllTotal

	end if
	RsF.close
	set RsF = nothing
	'第一次循环结束

	response.write TableStr & LoopStr & "</table><br/>"
end if

function LoopWrite(SqlStr)
	SqlStr = SqlStr
	set LoopRs = Server.CreateObject("adodb.recordset")
	LoopRs.open SqlStr,conn,1,1
	if not LoopRs.eof then
		
		LINENAME = LoopRs("LINENAME")
		DNAME = LoopRs("DNAME")
		CountDay = CSNG(LoopRs("CountDay"))
		OPENTIME = CSNG(LoopRs("OPENTIME"))
		STOPTIME = CSNG(LoopRs("STOPTIME"))
		'response.write LINENAME & "<br/>"
		
		'取出Target
		set TarRs = server.createobject("adodb.recordset")
		OraTarSql = "select OutPCS,nvl(T1,0) T1,NVL(C1,0) C1,NVL(C2,0) C2,NVL(C3,0) C3,NVL(C4,0) C4,NVL(C5,0) C5,OEE from report_hv_target"
		OraTarSql = OraTarSql & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' and Tdate < = '"&Startdate&"' ORDER BY Tdate desc"
		'response.write PRODUCTNAME & " - " & LINENAME & "<br/>"
		'response.write OraTarSql & "<br/>"
		TarRs.open OraTarSql,conn,1,1
		TargetCC = 0
		T1T = 100
		C1T = 100
		C2T = 100
		C3T = 100
		C4T = 100
		C5T = 100
		OEE = 100
		if not TarRs.eof then
			TargetCC = csng(TarRs("OutPCS"))
			T1T = CSNG(TarRs("T1"))
			C1T = CSNG(TarRs("C1"))
			C2T = CSNG(TarRs("C2"))
			C3T = CSNG(TarRs("C3"))
			C4T = CSNG(TarRs("C4"))
			C5T = CSNG(TarRs("C5"))
			OEE = TarRs("OEE")
		end if
		TarRs.close
		Set TarRs = nothing
		'取出Target

		T1G = CSNG(LoopRs("T1G"))
		T1B = CSNG(LoopRs("T1B"))
		T1A = T1G + T1B
		IF T1A = 0 THEN
			T1F = 0
		ELSE
			T1F = T1B / T1A * 100
		END IF
		
		C1G = CSNG(LoopRs("C1G"))
		C1B = CSNG(LoopRs("C1B"))
		C1A = C1G + C1B
		IF C1A = 0 THEN
			C1F = 0
		ELSE
			C1F = C1B / C1A * 100
		END IF
		
		C2G = CSNG(LoopRs("C2G"))
		C2B = CSNG(LoopRs("C2B"))
		C2A = C2G + C2B
		IF C2A = 0 THEN
			C2F = 0
		ELSE
			C2F = C2B / C2A * 100
		END IF
		
		C3G = CSNG(LoopRs("C3G"))
		C3B = CSNG(LoopRs("C3B"))
		C3A = C3G + C3B
		IF C3A = 0 THEN
			C3F = 0
		ELSE
			C3F = C3B / C3A * 100
		END IF
		
		C4G = CSNG(LoopRs("C4G"))
		C4B = CSNG(LoopRs("C4B"))
		C4A = C4G + C4B
		IF C4A = 0 THEN
			C4F = 0
		ELSE
			C4F = C4B / C4A * 100
		END IF
		'response.write C4A
		
		C5G = CSNG(LoopRs("C5G"))
		C5B = CSNG(LoopRs("C5B"))
		C5A = C5G + C5B
		IF C5A = 0 THEN
			C5F = 0
		ELSE
			C5F = C5B / C5A * 100
		END IF
		
		'C3GW = CSNG(LoopRs("C3GW"))
		'C5GW = CSNG(LoopRs("C5GW"))
		
		if T1F > T1T then
			T1FStr = "<font color='red'>" & Formatnumber(T1F,2,-1) & "%</font>"
		else
			T1FStr = "<font color='green'>" & Formatnumber(T1F,2,-1) & "%</font>"
		end if
		
		if C1F > C1T then
			C1FStr = "<font color='red'>" & Formatnumber(C1F,2,-1) & "%</font>"
		else
			C1FStr = "<font color='green'>" & Formatnumber(C1F,2,-1) & "%</font>"
		end if
		
		if C2F > C2T then
			C2FStr = "<font color='red'>" & Formatnumber(C2F,2,-1) & "%</font>"
		else
			C2FStr = "<font color='green'>" & Formatnumber(C2F,2,-1) & "%</font>"
		end if

		if C3F > C3T then
			C3FStr = "<font color='red'>" & Formatnumber(C3F,2,-1) & "%</font>"
		else
			C3FStr = "<font color='green'>" & Formatnumber(C3F,2,-1) & "%</font>"
		end if
		
		if C4F > C4T then
			C4FStr = "<font color='red'>" & Formatnumber(C4F,2,-1) & "%</font>"
		else
			C4FStr = "<font color='green'>" & Formatnumber(C4F,2,-1) & "%</font>"
		end if
		
		if C5F > C5T then
			C5FStr = "<font color='red'>" & Formatnumber(C5F,2,-1) & "%</font>"
		else
			C5FStr = "<font color='green'>" & Formatnumber(C5F,2,-1) & "%</font>"
		end if

		if ProductName = "RA" or ProductName = "Slim" then
			COG = Formatnumber(C5G,0,-1)
		else
			COG = Formatnumber(C3G,0,-1)
		end if
		COW = COG
		'response.write COW & "<br/>"

		Oee = Formatnumber(COG / (OPENTIME * 60),2,-1)
		Teep = Formatnumber(COG / CountDay / 720,2,-1)
		
		TargetCC = TargetCC * CountDay
		
		if DNAME = "D" then
			Shift = "Day"
			TrStyle = "style='background:#FFFFFF;text-align:center;'"
		elseif DNAME = "N" then
			Shift = "Night"
			TrStyle = "style='background:#EBEBEB;text-align:center;'"
		elseif DNAME = "SubTotal" then
			Shift = "SubTotal"
			TrStyle = "style='background:#FFE293;text-align:center;'"
			LineAllTargetCC = LineAllTargetCC + TargetCC
		elseif DNAME = "All Lines" then
			Shift = "All Lines"
			TargetCC = LineAllTargetCC
			TrStyle = "style='background:#FFAF93;text-align:center;'"
		end if

		LoopWrite = "<tr " & TrStyle & "><td>Line "&LineName&"</td><td>"& Shift & "</td><td>" & OPENTIME & "h</td><td>" & STOPTIME & "h</td><td>" & TargetCC & "</td><td>" & COW & "</td><td>" & COG & "</td>"
			if ProductName = "RA" or ProductName = "Slim" then LoopWrite = LoopWrite & "<td>" & C5FStr & "</td><td>" & C4FStr & "</td>"
		LoopWrite = LoopWrite & "<td>" & C3FStr & "</td><td>" & C2FStr & "</td><td>" & C1FStr & "</td><td>" & T1FStr & "</td><td>" & Oee & "%</td><td>" & Teep & "%</td></tr>"
		'response.write LoopWrite
	end if
	LoopRs.close
	set LoopRs = nothing
end function
%>
</body>
</html>
