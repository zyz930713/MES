<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/ChangeLineName.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<title>Daily Reporting</title>
<style type="text/css">
body{
font:normal 16px/16px "Arial";
}
table{
border-collapse:collapse;
border-spacing:0;
border:1px solid black;
}
table th{
border-color:black;
font:normal 16px/16px "Arial";
}
a{
color:#fff;
text-decoration:none;
}
body,td,th {
	font-family: Arial;
}
</style>
</head>
<body>
<%
Adate = request("Adate")
ProductName = request("ProductName")
Getdata = request("Getdata")

if Adate = "" then
	Adate = FormatTime((date()-1),2)
end if

WeekFirstDay = DateAdd("d",0 - Weekday(Adate,2) + 1, CDate(Adate))
'response.end
%>
<br>
<center>
<form method="post" action="DailyReportHvTest.asp">
产品：<select name="ProductName" id="TmProduct"></select>
		<select name="LineName" id="TmLine" style="display:none" ></select>
		<select name="CellName" id="TmCell" style="display:none" ></select>
		<script type="text/javascript">
			MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
		</script>
　日期：<input type="text" name="Adate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />
　<input type="submit" name="Getdata" value="刷新" />
</form>
<center>
<%
if Getdata = "刷新" or Getdata = "ok" then
	'Adate = Cdate(Adate)
%>
	<table width="1020" border="0" cellspacing="0" cellpadding="0" bgcolor="#7BB1DB">
	  <tr>
		<td width="20%">&nbsp;</td>
		<td width="60%" rowspan="2"><p style="font-size:20px;color:#FFF;text-align:center;"><strong><%=ProductName%></strong></p></td>
		<td width="20%" colspan="3">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><a href="DailyReportHvTest.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",-1,Adate)%>"><<</a>　<%=Adate%>　<a href="DailyReportHvTest.asp?Getdata=ok&ProductName=<%=ProductName%>&Adate=<%=dateadd("d",1,Adate)%>">>></a></td>
		<td width="3%" rowspan="2">&nbsp;</td>
		<td width="14%" rowspan="2"><img src="Images/Logo.jpg" width="124" height="52" align="middle"></td>
		<td width="3%" rowspan="2">&nbsp;</td>
	  </tr>
	  <tr>
		<td style="color:#fff;text-align:center;"><%=FormatTime(Adate,9)%></td>
		<td><p style="font-size:20px;color:#FFF;text-align:center;">Daily Report</p></td>
	  </tr>
	  <tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td colspan="3">&nbsp;</td>
	  </tr>
	</table>
	<br/>
	
	<table width="1020" border="1" cellpadding="0" cellspacing="0">
	<tr style="background:#B9E9FF;font-width:bold;">
		<th>Prod<br>Line</th>
		<th>DAY<br>Night</th>
		<th>Prod<br>Time</th>
		<th>Down<br>Time</th>
		<th>Target<br>CC</th>
		<%if ProductName = "RA" or ProductName = "Slim" then%>
			<th>Weekly<br>Quantity<br>Cell 5</th>
			<th>Output<br>Cell 5</th>
			<th>FOR<br>Cell 5</th>
			<th>FOR<br>Cell 4</th>
		<%else%>
			<th>Weekly<br>Quantity<br>Cell 3</th>
			<th>Output<br>Cell 3</th>
		<%end if%>
		<th>FOR<br>Cell 3</th>
		<th>FOR<br>Cell 2</th>
		<th>FOR<br>Cell 1</th>
		<th>FOR<br>Transfer</th>
		<th>OEE</th>
		<th>TEEP</th>
	</tr>
	<%
	dim TableStr,TableStrSum,RemarkStr,RemarkStrSum
	TableStr = ""
	TableStrSum = ""
	RemarkStrSum = ""
	
	set RsL = Server.CreateObject("adodb.recordset")
	set RsD = Server.CreateObject("adodb.recordset")
	SqlL = "select Linename from report_hv_data where adate = '"&Adate&"' and ProductName like '"&ProductName&"' GROUP BY Linename order by Linename"
	'response.write SqlL & "<br/><br/>"
	RsL.open SqlL,conn,1,1
	while not RsL.eof
		LineName = RsL("Linename")
		
		SqlD = "select dname from report_hv_data where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' order by dname"
		'response.write SqlD & "<br/><br/>"
		RsD.open SqlD,conn,1,1
		while not RsD.eof
			DNAME = RsD("DNAME")
			OraLoopSql = "SELECT PRODUCTNAME,LINENAME,DNAME,ShiftName,1 CountDay,NVL(OPENTIME,0) OPENTIME,NVL(STOPTIME,0) STOPTIME,NVL(T1G,0) T1G,NVL(T1B,0) T1B,NVL(C1G,0) C1G,NVL(C1B,0) C1B,NVL(C2G,0) C2G,NVL(C2B,0) C2B,NVL(C3G,0) C3G,NVL(C3B,0) C3B,NVL(C4G,0) C4G,NVL(C4B,0) C4B,NVL(C5G,0) C5G,NVL(C5B,0) C5B"
			OraLoopSql = OraLoopSql & ",NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0) C3GW"
			OraLoopSql = OraLoopSql & ",NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0) C5GW"
			OraLoopSql = OraLoopSql & " FROM report_hv_data RHD"
			OraLoopSql = OraLoopSql & " where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' and DNAME = '"&DNAME&"'"
			'response.write OraLoopSql & "<br/><br/>"
			TableStr = WriteTableStr(OraLoopSql)
			TableStrSum = TableStrSum & TableStr
			
			RsD.movenext
		Wend
		RsD.Close
		
		OraLoopSql = "SELECT PRODUCTNAME,LINENAME,'SubTotal' DNAME,'SubTotal' ShiftName,count(productname) CountDay,sum(NVL(OPENTIME,0)) OPENTIME,sum(NVL(STOPTIME,0)) STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B"
		OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C3GW"
		OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C5GW"
		OraLoopSql = OraLoopSql & " FROM report_hv_data RHD"
		OraLoopSql = OraLoopSql & " where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' group by PRODUCTNAME,LINENAME"
		'response.write OraLoopSql & "<br/><br/>"
		TableStrSum = TableStrSum & WriteTableStr(OraLoopSql)
		
		RsL.movenext
	Wend
	RsL.Close
	
	OraLoopSql = "SELECT PRODUCTNAME,'All' LINENAME,DNAME,'AllShift' ShiftName,count(productname) CountDay,sum(NVL(OPENTIME,0)) OPENTIME,sum(NVL(STOPTIME,0)) STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C3GW"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C5GW"
	OraLoopSql = OraLoopSql & " FROM report_hv_data RHD"
	OraLoopSql = OraLoopSql & " where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND DNAME = 'D' group by PRODUCTNAME,DNAME"
	'response.write OraLoopSql & "<br/><br/>"
	TableStrSum = TableStrSum & WriteTableStr(OraLoopSql)
	
	OraLoopSql = "SELECT PRODUCTNAME,'All' LINENAME,DNAME,'AllShift' ShiftName,count(productname) CountDay,sum(NVL(OPENTIME,0)) OPENTIME,sum(NVL(STOPTIME,0)) STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C3GW"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C5GW"
	OraLoopSql = OraLoopSql & " FROM report_hv_data RHD"
	OraLoopSql = OraLoopSql & " where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND DNAME = 'N' group by PRODUCTNAME,DNAME"
	'response.write OraLoopSql & "<br/><br/>"
	TableStrSum = TableStrSum & WriteTableStr(OraLoopSql)
	
	OraLoopSql = "SELECT PRODUCTNAME,'All' LINENAME,'All Lines' DNAME,'AllShift' ShiftName,count(productname) CountDay,sum(NVL(OPENTIME,0)) OPENTIME,sum(NVL(STOPTIME,0)) STOPTIME,sum(NVL(T1G,0)) T1G,sum(NVL(T1B,0)) T1B,sum(NVL(C1G,0)) C1G,sum(NVL(C1B,0)) C1B,sum(NVL(C2G,0)) C2G,sum(NVL(C2B,0)) C2B,sum(NVL(C3G,0)) C3G,sum(NVL(C3B,0)) C3B,sum(NVL(C4G,0)) C4G,sum(NVL(C4B,0)) C4B,sum(NVL(C5G,0)) C5G,sum(NVL(C5B,0)) C5B"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c3g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C3GW"
	OraLoopSql = OraLoopSql & ",SUM(NVL((select sum(c5g) from report_hv_data where '"&WeekFirstDay&"' < = adate and adate < = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' AND dname = rhd.dname),0)) C5GW"
	OraLoopSql = OraLoopSql & " FROM report_hv_data RHD"
	OraLoopSql = OraLoopSql & " where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' group by PRODUCTNAME"
	'response.write OraLoopSql & "<br/><br/>"
	TableStrSum = TableStrSum & WriteTableStr(OraLoopSql)
	
	response.write TableStrSum
	response.write "</table>"
	response.write "<br/>"

	response.write "<table width='1020' border='1' cellpadding='0' cellspacing='0'>"
	response.write "<tr style='background:#C9C9C9;height:26px;font-size:20px;'><th><strong>Remarks:</strong></th></tr>"
	response.write "<tr><td align='left'>" & RemarkStrSum & "</td></tr></table><br>"
end if



'制作表格函数
function WriteTableStr(SqlStr)
	SqlStr = SqlStr
	set LoopRs = Server.CreateObject("adodb.recordset")
	LoopRs.open SqlStr,conn,1,1
	if not LoopRs.eof then
		
		LINENAME = LoopRs("LINENAME")
		DNAME = LoopRs("DNAME")
		ShiftName = LoopRs("ShiftName")
		CountDay = CSNG(LoopRs("CountDay"))
		OPENTIME = CSNG(LoopRs("OPENTIME"))
		STOPTIME = CSNG(LoopRs("STOPTIME"))
		OPENTIME = CSNG(LoopRs("OPENTIME")) / 60
		STOPTIME = CSNG(LoopRs("STOPTIME")) / 60
		'response.write LINENAME & "<br/>"
		
		'取出Target
		TargetCC = 0
		T1T = 100
		C1T = 100
		C2T = 100
		C3T = 100
		C4T = 100
		C5T = 100
		OEE = 100
		set TarRs = server.createobject("adodb.recordset")
		set TarRsL = server.createobject("adodb.recordset")
		if LINENAME = "All" then
			'response.write DNAME & "<br/>"
			TarLoop = 0
			if DNAME = "All Lines" then
				SqlTarger = "select linename FROM report_hv_data RHD where adate = '"&Adate&"' AND ProductName = '"&ProductName&"'"
			else
				SqlTarger = "select linename FROM report_hv_data RHD where adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND DNAME = '"&DNAME&"'"
			end if
			
			response.write SqlTarger & "<br/><br/>"
			
			TarRsL.open SqlTarger,conn,1,1
			While not TarRsL.eof
				TarLineName = TarRsL("linename")
				OraTarSql = "select OutPCS,nvl(T1,0) T1,NVL(C1,0) C1,NVL(C2,0) C2,NVL(C3,0) C3,NVL(C4,0) C4,NVL(C5,0) C5,OEE from report_hv_target"
				OraTarSql = OraTarSql & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&TarLineName&"' and Tdate < = '"&Adate&"' ORDER BY Tdate desc"
				TarRs.open OraTarSql,conn,1,1
				if not TarRs.eof then
					TargetCC = TargetCC + csng(TarRs("OutPCS"))
					T1T = T1T + CSNG(TarRs("T1"))
					C1T = C1T + CSNG(TarRs("C1"))
					C2T = C2T + CSNG(TarRs("C2"))
					C3T = C3T + CSNG(TarRs("C3"))
					C4T = C4T + CSNG(TarRs("C4"))
					C5T = C5T + CSNG(TarRs("C5"))
					OEE = OEE + TarRs("OEE")
				end if
				TarRs.close
				TarRsL.movenext
				TarLoop = TarLoop + 1
			wend
			
			response.write TargetCC & " / "
			response.write TarLoop & " = "
			
			TargetCC = TargetCC / TarLoop
			T1T = T1T / TarLoop
			C1T = C1T / TarLoop
			C2T = C2T / TarLoop
			C3T = C3T / TarLoop
			C4T = C4T / TarLoop
			C5T = C5T / TarLoop
			OEE = OEE / TarLoop
			response.write TargetCC & "<br/><br/>"
		else
			OraTarSql = "select OutPCS,nvl(T1,0) T1,NVL(C1,0) C1,NVL(C2,0) C2,NVL(C3,0) C3,NVL(C4,0) C4,NVL(C5,0) C5,OEE from report_hv_target"
			OraTarSql = OraTarSql & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' and Tdate < = '"&Adate&"' ORDER BY Tdate desc"
			TarRs.open OraTarSql,conn,1,1
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
		end if
		Set TarRsL = nothing
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
		elseif DNAME = "All Lines" then
			Shift = "All Lines"
			TrStyle = "style='background:#FFAF93;text-align:center;'"
		end if

		WriteTableStr = "<tr " & TrStyle & "><td>Line "&ReplaceCcl(LineName)&"</td><td>"& Shift & "</td><td>" & OPENTIME & "h</td><td>" & STOPTIME & "h</td><td>" & TargetCC & "</td><td>" & COW & "</td><td>" & COG & "</td>"
			if ProductName = "RA" or ProductName = "Slim" then WriteTableStr = WriteTableStr & "<td>" & C5FStr & "</td><td>" & C4FStr & "</td>"
		WriteTableStr = WriteTableStr & "<td>" & C3FStr & "</td><td>" & C2FStr & "</td><td>" & C1FStr & "</td><td>" & T1FStr & "</td><td>" & Oee & "%</td><td>" & Teep & "%</td></tr>"
		'response.write WriteTableStr
		
		'取出Remark
		if LineName <> "All" then
			if DNAME <> "SubTotal" then
				RemarkTableStr = "<table width='1018' border='1' cellpadding='0' cellspacing='0'>" & WriteTableStr & "</table>"
				RemarkStrSum = RemarkStrSum & RemarkTableStr & WriteRemarkStr(Adate,ProductName,LineName,DName,ShiftName)
			end if
		end if
		'取出Remark
		
	end if
	LoopRs.close
	set LoopRs = nothing
end function
'制作表格函数

'提取Remark开始
function WriteRemarkStr(ADATE,ProductName,LineName,DName,ShiftName)
	RemarkStr = ""
	RemarkStrSumTmp = ""
	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&ADATE&"'"
	'response.write sql & "<br>"
	rs.open sql,ConnSql,1,1
	StopStr = ""
	PCStr = ""
	if not rs.eof then
		PcSum = rs("PcSum")
		PcQty = rs("PcQty")
		PlanStop = rs("PlanStop")
		if PlanStop = "0" OR PlanStop = "" then PlanStop = "0"
		StopStr = "计划停机: " & PlanStop & " 小时<br><br>"
		PCStr = "在线主载体数量: " & PcSum & "　待修主载体数量: " & PcQty & "<br>"
	end if
	rs.close
	set rs = nothing

	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&ADATE&"' order by Rid"
	'response.write sql & "<br>"
	rs.open sql,ConnSql,1,1
	if not rs.eof then
		RemarkStrSumTmp = RemarkStrSumTmp & ADATE & "　" & ProductName & " " & ReplaceCcl(LineName) & "　D/N: " & rs("DName") & "　Shift: " & ShiftName & " </BR></BR>"
		RemarkStrSumTmp = RemarkStrSumTmp & StopStr & PCStr
		do while not rs.eof
			Dlossn = rs("Dlossn")
			ADATE = FormatTime(ADATE,2)
			Dhtime = rs("Dhtime")
			Detime = rs("Detime")
			Dtime = rs("Dtime")
			Dequipment = rs("Dequipment")
			DPosition = rs("DPosition")
			Dproblem = rs("Dproblem")
			DReason = rs("DReason")
			Dsolution = rs("Dsolution")
				RemarkStr = Dhtime & " - " & Detime & "　停机：" & Dtime & " 分钟<br>"
				RemarkStr = RemarkStr & "位置：" & Dequipment & " " & DPosition & "<br>"
				RemarkStr = RemarkStr & "问题：" & Dproblem & "<br>"
				RemarkStr = RemarkStr & "原因：" & DReason & "<br>"
				RemarkStr = RemarkStr & "解决：" & Dsolution & "<br>"
				RemarkStr = RemarkStr & "<br>"
			RemarkStrSumTmp = RemarkStrSumTmp & RemarkStr
		rs.movenext
		loop
	else
		RemarkStrSumTmp = RemarkStrSumTmp & "没有维护记录<br><br>"
	end if
	rs.close
	set rs = nothing
	WriteRemarkStr = RemarkStrSumTmp & RemarkStr
end function
'提取Remark结束
%>

</table>
</body>
</html>
