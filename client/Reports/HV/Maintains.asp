<%session("Page_Role") = ",HV_Reports_Reader"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/include/Functions.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<title>Weekly Reporting</title>
<style type="text/css">
body{
font:normal 16px/16px "Arial";
}

body,td,th {
	font-family: Arial;
}
</style>
</head>
<body style="background-color:#FFFFFF;">
<%
ProductName = request("ProductName")
SearchEdit = request("SearchEdit")

Ayear = request("Ayear")
AWeek = request("AWeek")

if Ayear = "" then Ayear = Year(now())
if Aweek = "" then AWeek = GetWeekNo(date())

' EndDate = request("EndDate")
' Startdate = request("Startdate")

' if EndDate = "" then EndDate = date()
' if Startdate = "" then Startdate = DateAdd("d",-1, CDate(EndDate))


'response.write Aweek
'response.end
if SearchEdit <> "ToExcel" then
%>

<br/>
<center>
<form method="post" action="Maintains.asp">
���꣺<input type="text" name="Ayear" value="<%=Ayear%>" size="3"/>���ܣ�<input type="text" name="AWeek" value="<%=AWeek%>" size="3"/>
��<input type="submit" name="SearchEdit" value="����ˢ��" />
��<input type="submit" name="SearchEdit" value="ToExcel" />
</form>
<center>

<%
end if
if SearchEdit = "����ˢ��" or SearchEdit = "ToExcel" then
	set rs = Server.CreateObject("adodb.recordset")
	
	sql = "SELECT top 1 [Year],[WeekNr],[Month],[MonthInt],[Startdate],[Enddate] FROM [dbo].[YearWeek] where [Year] = '"&Ayear&"' and WeekNr = '"&AWeek&"'"
	'response.write sql
	rs.open sql,ConnSql,1,1
	if not rs.eof then
		Startdate = rs("Startdate")
		EndDate = rs("Enddate")
	end if
	rs.close
	'response.write Startdate & " - " & EndDate & "<br/>"
	
	HaveData = 0
	
	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where '"&Startdate&"' < = adate and adate < '"&EndDate&"' order by Adate asc,ProductName desc,LineName asc,DName"
	'response.write sql
	rs.open sql,ConnSql,1,1
		do while not rs.eof
		HaveData = 1
		Adate = FormatTime(rs("Adate"),2)
		ProductName = rs("ProductName")
		LineName = rs("LineName")
		DName = rs("DName")
		Dlossn = rs("Dlossn")
		DlossName = XlName(right(Dlossn,1))
		Dhtime = rs("Dhtime")
		Detime = rs("Detime")
		Dtime = rs("Dtime")
		Dequipment = rs("Dequipment")
		DPosition = rs("DPosition")
		Dproblem = rs("Dproblem")
		DReason = rs("DReason")
		Dsolution = rs("Dsolution")
		TableStr = TableStr & "<tr><td>" & Adate & "</td><td>" & ProductName & "</td><td>" & LineName & "</td><td>" & DName & "</td><td>" & DlossName & "</td><td>" & Dhtime & " - " & Detime & "</td><td>" & Dtime & "</td><td>" & Dequipment & "</td><td>" & DPosition & "</td><td align='left'>" & Dproblem & "</td><td>" & DReason & "</td><td>" & Dsolution & "</td></tr>"
		rs.movenext
		loop
	rs.close
	
	if SearchEdit = "ToExcel" then
		Response.ContentType ="application/vnd.ms-excel"  
		Response.AddHeader "Content-Disposition","attachment;filename=Maintains.xls"
	end if
	
	if HaveData = 1 then
		response.write "<table width='100%' border='1' cellspacing='0' cellpadding='0'>"
		response.write "<tr><th>����</th><th>��Ʒ</th><th>��</th><th>D/N</th><th>ά����Ŀ</th><th>��ֹʱ��</th><th>ʱ��</th><th>�豸</th><th>��λ</th><th>��������</th><th>����ԭ��</th><th>�������</th></tr>"
		response.write TableStr
		response.write "</table><br/>"
	else
		response.write "<div style='width:960px;height:480px;background:#FFFFFF;line-height:480px;font-size:32px;'><p>û���ҵ���������ݣ�</p></div>"
	end if
		
end if

function XlName(Xltype)
	select case Xltype
		case 1
			XlName = "�����㱸��"
		case 2
			XlName = "ά��"
		case 3
			XlName = "�豸����"
		case 4
			XlName = "�ƻ��Լ��"
		case 5
			XlName = "��λ�ӹ�ʱ��"
		case 6
			XlName = "�ٶȲ�ƥ��"
		case 7
			XlName = "Сͣ��"
		case 8
			XlName = "���̷�Ʒ"
		case 9
			XlName = "��ʼ����Ʒ"
	end select
end function
%>
</table>
</body>
</html>
