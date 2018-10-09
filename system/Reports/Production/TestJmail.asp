<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/SendJMail.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
reportTime=replace(dateadd("d",-1,now()),"/","-")
fromTime=formatdate(reportTime,"dd.mm.yyyy")&" 07:15"
mailTitleFromDate = formatdate(reportTime,"dd.mm.yyyy")
toTime=formatdate(now(),"dd.mm.yyyy")&" 07:14"
reportDay=longweekdayconvert(weekday(reportTime))&","&formatdate(reportTime,"dd.mm.yyyy")
htmls="<html xmlns:v='urn:schemas-microsoft-com:vml' xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns:m='http://schemas.microsoft.com/office/2004/12/omml' xmlns='http://www.w3.org/TR/REC-html40'><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><meta name=Generator content='Microsoft Word 14 (filtered medium)'><style><!--"
htmls=htmls&"/* Font Definitions */"
htmls=htmls&"@font-face"
htmls=htmls&"	{font-family:Arial;"
htmls=htmls&"	panose-1:2 1 6 0 3 1 1 1 1 1;}"
htmls=htmls&"/* Style Definitions */"
htmls=htmls&"p.MsoNormal, li.MsoNormal, div.MsoNormal"
htmls=htmls&"	{margin:0cm;"
htmls=htmls&"	margin-bottom:.0001pt;"
htmls=htmls&"	text-align:justify;"
htmls=htmls&"	text-justify:inter-ideograph;"
htmls=htmls&"	font-size:10.5pt;"
htmls=htmls&"	font-family:'Arial','sans-serif';}"
htmls=htmls&"a:link, span.MsoHyperlink"
htmls=htmls&"	{mso-style-priority:99;"
htmls=htmls&"	color:blue;"
htmls=htmls&"	text-decoration:underline;}"
htmls=htmls&"a:visited, span.MsoHyperlinkFollowed"
htmls=htmls&"	{mso-style-priority:99;"
htmls=htmls&"	color:purple;"
htmls=htmls&"	text-decoration:underline;}"
htmls=htmls&"span.EmailStyle17"
htmls=htmls&"	{mso-style-type:personal-compose;"
htmls=htmls&"	font-family:'Arial','sans-serif';"
htmls=htmls&"	color:windowtext;}"
htmls=htmls&".MsoChpDefault"
htmls=htmls&"	{mso-style-type:export-only;"
htmls=htmls&"	font-family:'Arial','sans-serif';}"
htmls=htmls&"/* Page Definitions */"
htmls=htmls&"@page WordSection1"
htmls=htmls&"	{size:612.0pt 792.0pt;"
htmls=htmls&"	margin:72.0pt 90.0pt 72.0pt 90.0pt;}"
htmls=htmls&"div.WordSection1"
htmls=htmls&"	{page:WordSection1;}"
htmls=htmls&"--></style><!--[if gte mso 9]><xml>"
htmls=htmls&"<o:shapedefaults v:ext='edit' spidmax='1026' />"
htmls=htmls&"</xml><![endif]--><!--[if gte mso 9]><xml>"
htmls=htmls&"<o:shapelayout v:ext='edit'>"
htmls=htmls&"<o:idmap v:ext='edit' data='1' />"
htmls=htmls&"</o:shapelayout></xml><![endif]--></head><body>"
'绘制页面上方页头
htmls=htmls&"<table width='100%' border='1' cellpadding='0' cellspacing='0' bgcolor='#0099FF'>"
htmls=htmls&"<tr><td height='90'><table width='100%' border='0' cellpadding='0' cellspacing='0'>"
htmls=htmls&"<tr align='center'><th>Daily Production Report - KEB IT Internal Demo Test</th><td width='30' rowspan='2' align='right'><img src='http://"&application("HostServer")&"/Images/rpt-logo.jpg'></td></tr>"
htmls=htmls&"<tr align='center' ><td><table width='80%' border='0' cellpadding='0' cellspacing='0'>"
htmls=htmls&"<tr><td><b>Report Day: &nbsp;</b>"&reportDay&"</td>"
htmls=htmls&"<td><b>Page: &nbsp;</b>1/1</td>"
htmls=htmls&"<td><b>Time:&nbsp;</b> "&fromTime&" - "&toTime&"</td></tr>"
htmls=htmls&"<tr ><td><b>Author:&nbsp;</b>Jeff Xu, Zhang Hui, Wang Binglin</td><td><b>Print Date:&nbsp;</b>"&formatdate(now(),"dd.mm.yyyy")&"</td></tr></table>"
htmls=htmls&"</td></tr></table></td></tr></table><br>Test</body></html>"


mailContents = htmls

MailSwitch = 1			'邮件发送/测试用显示开关
if MailSwitch > 0 then
	select case MailSwitch
			case 1
				mailto="chao.yao@knowles.com"
			case 2
				mailto="bl.wang@knowles.com;Jeff.Xu@knowles.com;hui.zhang@knowles.com;Xuehui.Chen@knowles.com;mark.fan@knowles.com;fanny.zhang@knowles.com;kunlun.liu@knowles.com;Vivian.Huang@knowles.com;Ivan.li@knowles.com;Young.Li@knowles.com;bo.wang@knowles.com;chao.yao@knowles.com"
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
	SendJMail application("MailSender"),mailTo,"BPS Daily Report (Demo Version) - " & mailTitleFromDate,mailContents
else
	response.Write(mailContents)
end if
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->