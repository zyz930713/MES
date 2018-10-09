<!--#include file="conn.asp"-->
<!--#include virtual="/include/function.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="refresh" content="300" charset="utf-8" />
<title>StationFOR</title>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
</head>
<body style="margin:20px;font-size:14px;background-color:#FFFFFF;">
<center>
<%	
	SearchEdit = request("SearchEdit")
	Sdatetime = request("Sdatetime")
	Edatetime = request("Edatetime")
	if Sdatetime = "" or Edatetime = "" then
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
	end if
	
if SearchEdit <> "ToExcel" then
%>
<form name="form1" method="post" action="Reports.asp" >
起始时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 07:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Sdatetime" value="<%=Sdatetime%>" size="20" />&nbsp;
　终止时间：<input type="text" class="Wdate" onFocus="WdatePicker({startDate:'%y-%M-%d 19:15:00',dateFmt:'yyyy-MM-dd HH:mm:ss',alwaysUseStartDate:true})" name="Edatetime" value="<%=Edatetime%>" size="20" />&nbsp;
　<input type="submit" name="SearchEdit" value="刷新" />
　<input type="submit" name="SearchEdit" value="ToExcel" />
</form>
<%
end if

set rs = Server.CreateObject("adodb.recordset")
sql1 = "SELECT [LineName] FROM [dbo].[CfgLine]"
rs.open sql1,conn,1,1
set RsL = Server.CreateObject("adodb.recordset")
do while not rs.eof
	LineName = rs("LineName")
	Sql2 = "EXEC [dbo].[GetOutputFORDeatil] @LineName = N'"&LineName&"', @StartTime = N'"&Sdatetime&"', @EndTime = N'"&Edatetime&"'"
	response.write Sql2 & "<br/>"
	RsL.open Sql2,conn,1,1
	if not RsL.eof then
		WriteTable = "<table width='auto' border='1' cellspacing='0' cellpadding='0'>"
		WriteTableTitle = "<tr><th colspan='24'>"&LineName&"</th></tr>"
		WriteContentTitle = "<th>工位名</th>"
		
		WriteTableContentLBad = "<td>L工位</td>"
		WriteTableContentLAll = "<td>L总数</td>"
		WriteTableContentLFOR = "<td>FOR</td>"
		
		WriteTableContentRBad = "<td>R工位</td>"
		WriteTableContentRAll = "<td>R总数</td>"
		WriteTableContentRFOR = "<td>FOR</td>"
		
		WriteTableContentABad = "<td>工位</td>"
		WriteTableContentAAll = "<td>总数</td>"
		WriteTableContentAFOR = "<td>FOR</td>"
		do while not RsL.eof
			
			WriteContentTitle = WriteContentTitle & "<th>" & RsL("StName") & "</th>"
			WriteTableContentLBad = WriteTableContentLBad & "<td>" & RsL("PrL_Bad") & "</td>"
			WriteTableContentLAll = WriteTableContentLAll & "<td>" & RsL("PrL_Total") & "</td>"
			WriteTableContentLFOR = WriteTableContentLFOR & "<td>" & Formatnumber(RsL("PrL_FOR"),2,-1) & "%</td>"
			
			WriteTableContentRBad = WriteTableContentRBad & "<td>" & RsL("PrR_Bad") & "</td>"
			WriteTableContentRAll = WriteTableContentRAll & "<td>" & RsL("PrR_Total") & "</td>"
			WriteTableContentRFOR = WriteTableContentRFOR & "<td>" & Formatnumber(RsL("PrR_FOR"),2,-1) & "%</td>"
			
			WriteTableContentABad = WriteTableContentABad & "<td>" & RsL("PrA_Bad") & "</td>"
			WriteTableContentAAll = WriteTableContentAAll & "<td>" & RsL("PrA_Total") & "</td>"
			WriteTableContentAFOR = WriteTableContentAFOR & "<td>" & Formatnumber(RsL("PrA_FOR"),2,-1) & "%</td>"
			
		RsL.movenext
		loop
	end if
	RsL.close
	
	WriteContentTitle = "<tr>" & WriteContentTitle & "</tr>"
	WriteTableContentLBad = "<tr>" & WriteTableContentLBad & "</tr>"
	WriteTableContentLAll = "<tr>" & WriteTableContentLAll & "</tr>"
	WriteTableContentLFOR = "<tr>" & WriteTableContentLFOR & "</tr>"
	
	WriteTableContentRBad = "<tr>" & WriteTableContentRBad & "</tr>"
	WriteTableContentRAll = "<tr>" & WriteTableContentRAll & "</tr>"
	WriteTableContentRFOR = "<tr>" & WriteTableContentRFOR & "</tr>"
	
	WriteTableContentABad = "<tr>" & WriteTableContentABad & "</tr>"
	WriteTableContentAAll = "<tr>" & WriteTableContentAAll & "</tr>"
	WriteTableContentAFOR = "<tr>" & WriteTableContentAFOR & "</tr>"
	
	WriteTable = WriteTable & WriteTableTitle & WriteContentTitle & WriteTableContentLBad & WriteTableContentLAll & WriteTableContentLFOR & WriteTableContentRBad & WriteTableContentRAll & WriteTableContentRFOR & WriteTableContentABad & WriteTableContentAAll & WriteTableContentAFOR & "</table>"
	response.write WriteTable & "<br/>"
	
rs.movenext
loop
set RsL = nothing
rs.close
set rs = nothing
if SearchEdit = "ToExcel" then
		Response.ContentType ="application/vnd.ms-excel"  
		Response.AddHeader "Content-Disposition","attachment;filename=OutputReport.xls"
end if

%>
</center>
</body>
</html>
