<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
reportname=trim(request("reportname"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by REPORT_TIME desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if reportname<>"" then
where=where&" and WIP_NAME='"&reportname&"'"
end if
if fromdate<>"" then
where=where&" and REPORT_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and REPORT_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&reportname="&reportname&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Process/WIP/WIPList.asp"
SQL="select * from WEB_REPORT.HOURLY_WIP_LIST where NID is not null "&where&order
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript">
function form2check()
{
var flag=0;
	with (document.form2)
	{
		for(var i=0;i<section.length;i++)
		{
			if (section[i].checked)
			{
			flag=1;
			}
		}
	}
	if (flag==0)
	{
	alert("You must select one section");
	return false;
	}
}
</script>
</head>

<body>
<form name="form1" method="post" action="/Reports/Process/PartWIP/PartWIPList.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span>Search WIP Reports </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Report Name </span> </td>
    <td height="20"><input name="reportname" type="text" id="reportname" value="<%=reportname%>"></td>
    <td>Report Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate2" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="48%">Browse WIP Report </td>
        <td width="52%"><div align="right">User:
          <% =session("User") %>
        </div></td>
      </tr>
    </table>    </td>
</tr>
		<%if admin=true then%>
		<form action="/Reports/Process/WIP/GenerateWIP.asp" method="post" name="form2" target="_self" onSubmit="return form2check()">
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy">
		  Section type: <%FactoryRight "S."%><%=getSection("RADIO","",factorywhereoutside,"","")%>
          <br>
          <input name="Generate" type="submit" id="Generate" value="Generate WIP report at once">
</td>
</tr>
		  </form>
		  <%end if%>
<tr>
  <td height="20" colspan="5"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Delete</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=REPORT_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WIP_NAME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=WIP_NAME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">Creator</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:location.href='DeleteWIPReport.asp?id=<%=rs("NID")%>&WIP_name=<%=rs("WIP_NAME")%>&path=<%=path%>&query=<%=query%>'">Delete</span></div></td>
    <td height="20"><div align="center"><% =formatdate(rs("REPORT_TIME"),application("longdateformat"))%></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('WIPReport.asp?WIP_id=<%=rs("NID")%>&WIP_name=<%=rs("WIP_NAME")%>&WIP_report_time=<%=rs("REPORT_TIME")%>&section_id=<%=rs("SECTION_ID")%>')"><%= rs("WIP_NAME") %></span>&nbsp;</div></td>
    <td><div align="center"><%= rs("CREATOR_CODE") %></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="5"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->