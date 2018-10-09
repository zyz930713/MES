<%session("Page_Role") = ",HV_Reports_Reader"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="../../styles/basic.css" />
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<title>Daily Reporting</title>
<style type="text/css">
body{
font:normal 16px/16px "Arial";
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
SearchEdit = request("SearchEdit")

if Adate = "" then
	Adate = FormatTime((date()-1),2)
end if

WeekFirstDay = DateAdd("d",0 - Weekday(Adate,2) + 1, CDate(Adate))
'response.end
%>
<br/>
<center>
<form method="post" action="DailyOutput.asp">
　日期：<input type="text" name="Adate" class="Wdate" id="d122" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('d122_1').value=$dp.cal.getP('W','W');$dp.$('d122_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />&nbsp;
　<input type="submit" name="SearchEdit" value="刷新" />
</form>
<center>
<%
'response.end
if SearchEdit = "刷新" then
	Startdate = Cdate(Adate)
	EndDate = Startdate
	
	set RsF = Server.CreateObject("adodb.recordset")
	set RsS = Server.CreateObject("adodb.recordset")
	set RsT = Server.CreateObject("adodb.recordset")
	LoopStr = ""
	LineAllTargetCC = 0
	SqlF = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME FROM report_hv_data RHD where adate = '"&EndDate&"' group by PRODUCTNAME,LINENAME order by PRODUCTNAME,LINENAME"
	'response.write SqlF & "<br/><br/>"
	RsF.open SqlF,conn,1,1
	if not RsF.eof then
		do while not RsF.eof
			ProductNameF = RsF("PRODUCTNAME")
			LineNameF = RsF("LINENAME")
			'response.write ProductNameF & " - " & LineNameF & "<br/>"
			
			'第二次循环开始
			SqlS = "SELECT count(Productname) CountDay,PRODUCTNAME,LINENAME,DNAME FROM report_hv_data RHD where adate = '"&EndDate&"' and ProductName like '"&ProductNameF&"' and LineName = '"&LineNameF&"' group by PRODUCTNAME,LINENAME,DNAME order by PRODUCTNAME,LINENAME"
			'response.write SqlS & "<br/><br/>"
			RsS.open SqlS,conn,1,1
			if not RsS.eof then
				
				do while not RsS.eof
					DNameS = RsS("DNAME")
					if DNameS = "D" then
					Shift = " Day "
					TrStyle = "style='background:#FFFFCC;text-align:center;'"
				elseif DNameS = "N" then
					Shift = "Night"
					TrStyle = "style='background:#CCFFFF;text-align:center;'"
				end if
				
					response.write "<table width='90%' border='1' cellspacing='0' cellpadding='0'" & TrStyle & ">"
					'response.write ProductNameF & " - " & LineNameF & " - " & DNameS & "<br/><br/>"
					response.write "<tr><th colspan='4'>" & ProductNameF & " - " & LineNameF & " - " & Shift & "</th></tr>"
					response.write "<tr><th width='25%'>工位名</th><th width='25%'>良品</th><th width='25%'>废品</th><th width='25%'>FOR</th></tr>"
					
					'第三次循环开始
					SqlT = "SELECT PRODUCTNAME,LINENAME,DNAME,NVL(OPENTIME,0)/60 OPENTIME,NVL(STOPTIME,0)/60 STOPTIME,NVL(T1G,0) T1G,NVL(T1B,0) T1B,NVL(C1G,0) C1G,NVL(C1B,0) C1B,NVL(C2G,0) C2G,NVL(C2B,0) C2B,NVL(C3G,0) C3G,NVL(C3B,0) C3B,NVL(C4G,0) C4G,NVL(C4B,0) C4B,NVL(C5G,0) C5G,NVL(C5B,0) C5B FROM report_hv_data RHD"
					SqlT = SqlT & " where adate = '"&EndDate&"' and ProductName like '"&ProductNameF&"' and LineName = '"&LineNameF&"' and Dname = '"&DNameS&"'"
					'response.write SqlT & "<br/><br/>"
					LoopStr = LoopWrite(SqlT)
					'第三次循环结束
					response.write LoopStr
				response.write "</tbale><br/>"
				RsS.movenext
				loop
				
			end if
			RsS.close
			'第二次循环结束
			
		RsF.movenext
		loop
	end if
	RsF.close
	set RsF = nothing
	'第一次循环结束
end if


function LoopWrite(SqlStr)
	SqlStr = SqlStr
	LoopWrite = ""
	set LoopRs = Server.CreateObject("adodb.recordset")
	LoopRs.open SqlStr,conn,1,1
	if not LoopRs.eof then
		ProductNameL = LoopRs("ProductName")
		LINENAMEL = LoopRs("LINENAME")
		DNAMEL = LoopRs("DNAME")

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
		
		C5G = CSNG(LoopRs("C5G"))
		C5B = CSNG(LoopRs("C5B"))
		C5A = C5G + C5B
		IF C5A = 0 THEN
			C5F = 0
		ELSE
			C5F = C5B / C5A * 100
		END IF
		
		set RsOnline = Server.CreateObject("adodb.recordset")
		sql = "SELECT [Rid],[QtySM],[QtyOC] ,[QtyTS],[QtyMT],[QABlock],[QARelease] FROM [dbo].[Day_Data_Online] where ProductName = '"&ProductNameL&"' and LineName = '"&LINENAMEL&"' and DName = '"&DNAMEL&"' and Adate = '"&Adate&"' order by Rid"
		RsOnline.open sql,ConnSql,1,1
			QtySM = 0
			QtyOC = 0
			QtyTS = 0
			QtyMT = 0
			QABlock = 0
			'QARelease = 0
		if not RsOnline.eof then
			QtySM = RsOnline("QtySM")
			QtyOC = RsOnline("QtyOC")
			QtyTS = RsOnline("QtyTS")
			QtyMT = RsOnline("QtyMT")
			QABlock = RsOnline("QABlock")
			'QARelease = RsOnline("QARelease")
		end if
		RsOnline.close
		set RsOnline = nothing

			LoopWrite = LoopWrite & "<tr><td>Transfer</td><td>" & T1G & "</td><td>" & T1B & "</td><td>" & Formatnumber(T1F,2,-1) & "</td></tr>"
			LoopWrite = LoopWrite & "<tr><td>Cell1</td><td>" & C1G & "</td><td>" & C1B & "</td><td>" & Formatnumber(C1F,2,-1) & "</td></tr>"
			LoopWrite = LoopWrite & "<tr><td>Cell2</td><td>" & C2G & "</td><td>" & C2B & "</td><td>" & Formatnumber(C2F,2,-1) & "</td></tr>"
			LoopWrite = LoopWrite & "<tr><td>Cell3</td><td>" & C3G & "</td><td>" & C3B & "</td><td>" & Formatnumber(C3F,2,-1) & "</td></tr>"
			if ProductNameL <> "PETRA" then
				LoopWrite = LoopWrite & "<tr><td>Cell4</td><td>" & C4G & "</td><td>" & C4B & "</td><td>" & Formatnumber(C4F,2,-1) & "</td></tr>"
				LoopWrite = LoopWrite & "<tr><td>Cell5</td><td>" & C5G & "</td><td>" & C5B & "</td><td>" & Formatnumber(C5F,2,-1) & "</td></tr>"
			end if
			LoopWrite = LoopWrite & "<tr><td colspan='4'>采样:" & QtySM & "　目测废品:" & QtyOC & "　清零废品:" & QtyTS & "　废膜/顶片:" & QtyMT & "　封存:" & QABlock & "</td></tr>"
	end if
	LoopRs.close
	set LoopRs = nothing
end function
%>

</table>
</body>
</html>
