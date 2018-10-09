<%
Dim conn
Set conn = Server.CreateObject("ADODB.Connection") 
conn.Open "driver={SQL Server};server=10.6.100.82;uid=kardexread;pwd=kardexread;database=PP5000V37KEB_PROD" 
reportTime=replace(dateadd("d",-1,now()),"/","-")
mailTitleFromDate = formatdate(reportTime,"yyyymmdd")
fromTime=formatdate(reportTime,"yyyymmdd")&" 08:15:00"
toTime=formatdate(now(),"yyyymmdd")&" 08:15:00"
htmls="<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body>"
htmls = htmls & "<table width='550' border='1' cellpadding='0' cellspacing='0' ><tr align='center' bgcolor='#9ADAFF'><td colspan='2'><b>最新零库存备件清单</b></td></tr>"
htmls = htmls & "<tr align='center' bgcolor='#9ADAFF'><td colspan='2'><b>" & fromTime & " 至 " & toTime & "</b></td></tr>"
set rs = Server.CreateObject("adodb.recordset")
sql = "SELECT b.BUCHUNGSDATUM,b.ARTIKEL,b.menge Omenge,b.BESTAND,lza.MENGE OnHand FROM [PP5000V37KEB_PROD].[dbo].[BUCHUNG] B inner JOIN LO_ZU_ART lza ON b.ARTIKEL=lza.ARTIKEL"
sql = sql & " WHERE b.FKTCODE NOT IN('11','12','17','77','88') and lza.menge < '1' AND b.BESTAND < '1' and '"&fromTime&"' < = b.buchungsdatum and b.buchungsdatum < '"&toTime&"'"
rs.open sql,conn,1,1
if not rs.eof then
	do while not rs.eof
		htmls = htmls & "<tr><td width='400'><img src='http://10.6.100.44/spimages/" & rs("artikel") & ".jpg' border='0' width='400' height='300'></td><td align='center'><a href='http://10.6.100.44/uselist.asp?ns12nc=" & rs("artikel") & "'>" & rs("artikel") & "</a></td></tr>"
	rs.movenext
	loop
else
	htmls = htmls & "<tr><td colspan='2' align='center'>昨日没有零库存备件</td></tr>"
end if

rs.close
set rs = nothing
mailContents = htmls & "</body></html>"

MailSwitch = 2
if MailSwitch > 0 then
	select case MailSwitch
			case 1
				mailto="chao.yao@knowles.com"
			case 2
				mailto="zhengkai.qiu@knowles.com;Luke.Huang@knowles.com;peili.liu@knowles.com;kardex.bj@knowles.com;chao.yao@knowles.com"
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
	SendJMail application("MailSender"),mailTo,"最新零库存备件清单  " & fromTime & " - " & toTime,mailContents
else
	response.Write(mailContents)
end if

function SendJMail(sender,reciever,subject,body)
	on error resume next
	set JMail=CreateObject("JMail.Message") 
	JMail.ContentType = "text/html"
	JMail.Charset ="utf-8"
	JMail.From = "kardex.bj@knowles.com"
	if instr(reciever,";")>1 then
		array_mail=split(reciever,";")
		For i=0 to ubound(array_mail) 
			JMail.AddRecipient array_mail(i)
		Next
	else
		JMail.AddRecipient reciever
	end if
	JMail.Subject = subject
	JMail.Body = body
	JMail.Send (application("MailServer"))
	set JMail = Nothing
end function

function formatdate(vdate,vformat)
if vdate<>"" then
	vyear=year(vdate)
	vmonth=month(vdate)
	vday=day(vdate)
	
	vhour=hour(vdate)
	vminute=minute(vdate)
	vsecond=second(vdate)
	
	select case lcase(vformat)
	case "dd.mm.yyyy"
	formatdate=repeatstring(vday,"0",2)&"."&repeatstring(vmonth,"0",2)&"."& vyear
	case "yyyymmdd"
	formatdate=vyear&"-"&vmonth&"-"&vday
	case "@yyyymmdd@"
	formatdate=vyear&"-"&repeatstring(vmonth,"0",2)&"-"&repeatstring(vday,"0",2)
	case "ddmmyyyy"
	formatdate=vday&"/"&vmonth&"/"&vyear
	case "mmddyyyy"
	formatdate=vmonth&"/"&vday&"/"&vyear
	case "yyyymmdd hhnnss"
	formatdate=vyear&"-"&vmonth&"-"&vday&" "&vhour&":"&vminute&":"&vsecond
	case "ddmmyyyy hhnnss"
	formatdate=vday&"/"&vmonth&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case "mmddyyyy hhnnss"
	formatdate=vmonth&"/"&vday&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case "hhnnss"
	formatdate=vhour&":"&vminute&":"&vsecond
	case "hhnn"
	formatdate=vhour&":"&vminute
	case "hh"
	formatdate=vhour
	case else
	formatdate=vdate
	end select
end if
end function
conn.close
set conn = nothing
%>