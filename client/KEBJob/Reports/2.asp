<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
Response.buffer=true 
'StartDate = formatdate(dateadd("d",-1,now()),"yyyymmdd")
EndDate = formatdate(now(),"yyyymmdd")

reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
'response.write replace(StartDate & " 07:15:00","-","/")
'response.write "</br>"
'response.write replace(EndDate & " 07:15:00","-","/")

mailTitleFromDate = formatdate(reportTime,"dd.mm.yyyy")
'reportDay = 'GetWeekDayEn(weekday(reportTime)) & "," & formatdate(now(),"dd.mm.yyyy")
reportDay = formatdate(now(),"dd.mm.yyyy")


htmls="<html><head><title></title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
htmls=htmls&"<style type='text/css'>"
htmls=htmls&"body {font:normal 12px/13px 'Arial';}"
htmls=htmls&"table {border-collapse:collapse;border-spacing:0;border:1px solid black;}"
htmls=htmls&"table th{border-color:black;font-size:14px;}"
htmls=htmls&"table td{border-color:black;}"
htmls=htmls&"img{margin:5px;}"
htmls=htmls&"</style></head><body>"


htmls=htmls&"<table width='100%' border='1' cellpadding='0' cellspacing='0' bgcolor='#7BB1DB'>"
htmls=htmls&"<tr><td><table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='#7BB1DB' align='center'>"
htmls=htmls&"<tr><td width='10%' rowspan='4'>&nbsp;</td>"
htmls=htmls&"<td colspan='2'><span style='height:40px;line-height:40px;font-size:24px;'><b>Overdate un-returend product reminding - Beijing</b></span></td>"
htmls=htmls&"<tr><td width='24%'><b>Report Day: &nbsp;</b>"&reportDay&"</td>"
htmls=htmls&"<tr><td><b>Author:&nbsp;</b>Xu YiDing</td>"
htmls=htmls&"<tr><td colspan='2'>&nbsp;</td></tr></table></td></tr></table></br>"

set rsjs1 = server.createobject("adodb.recordset")
SQL = "select SNNO, BOXNO, SENDAREA, OPERATORCODE, GETCODE, ACCEPTCODE, PTCSTATE,TXTCOMMENTS, AREA, EXPECTDATE from PTC_SN where EXPECTDATE is not NULL and SYSDATE > EXPECTDATE and SELECTTYPE <> '不归还' "
'response.write(SQL)

rsjs1.Open SQL,conn

htmls=htmls&"<table width='100%' border='1' cellpadding='5' cellspacing='0' bgcolor='#7BB1DB'>"

	htmls=htmls&"<tr>"
	for each x in rsjs1.Fields
		htmls=htmls&"<th> "& x.name &" </th>"
	next
	htmls=htmls&"</tr>"
	do until rsjs1.eof
	htmls=htmls&"<tr>"
	htmls=htmls&"<td>"&rsjs1("SNNO")&"</td>"
	htmls=htmls&"<td>"&rsjs1("BoxNo")&"</td>"
	htmls=htmls&"<td>"&rsjs1("SENDAREA")&"</td>"
	htmls=htmls&"<td>"&rsjs1("OPERATORCODE")&"</td>"
	htmls=htmls&"<td>"&rsjs1("GETCODE")&"</td>"
	htmls=htmls&"<td><a href='http://keb-hr50:8082/admin/KEB_DisplayB.asp?EmployeeNo="&rsjs1("ACCEPTCODE")&"' >"&rsjs1("ACCEPTCODE")&"</td>"
	htmls=htmls&"<td>"&rsjs1("PTCSTATE")&"</td>"	
	htmls=htmls&"<td>"&rsjs1("TXTCOMMENTS")&"</td>"
	htmls=htmls&"<td>"&rsjs1("AREA")&"</td>"
	htmls=htmls&"<td>"&rsjs1("EXPECTDATE")&"</td>"
	htmls=htmls&"</tr>"
	
	rsjs1.MoveNext
	
	loop 
	rsjs1.close 
	conn.close
	
htmls=htmls&"</table>"

mailContents = htmls & "</body></html>"

MailSwitch = 2
if MailSwitch > 0 then
	select case MailSwitch
			case 1
				mailto="chao.yao@knowles.com"
			case 2
				mailto="Young.Li@knowles.com;chao.yao@knowles.com"
			case 3
				mailto="linlin.hao@knowles.com;henry.qi@knowles.com;fanny.zhang@knowles.com"
			case else
				response.Write(mailContents)
	end select
%>
<script type="text/javascript">
	window.opener=null;
	window.open("","_self");
	window.close();
</script>
<%
	SendJMail application("MailSender"),mailTo,"return reminder - " & mailTitleFromDate,mailContents
else
	response.Write(mailContents)
end if

%>