<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Process/ProcessCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
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
pagename="/Reports/WIP/WIPList.asp"
SQL="select * from FAMILY_OUTPUT_LIST where NID is not null "&where&order
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
	with (document.form2)
	{
		if (factory.selectedIndex==0)
		{
		alert("You must select one factory");
		return false;
		}
		if (section.selectedIndex==0)
		{
		alert("You must select one section");
		return false;
		}
	}
}
</script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<form name="form1" method="post" action="/Reports/Process/FamilyOutput/FamilyOutputList.asp" onSubmit="form2check()">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span>Search Family Output Reports </span></td>
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
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
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
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#FFFFFF">
  <%if admin=true then%>
  <form action="/Reports/Process/FamilyOutput/GenerateFamilyOutput.asp" method="post" name="form2" target="_self" onSubmit="return form2check()">
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy">Generating conditions</td>
    </tr>
    <tr>
      <td>Factory: </td>
      <td><select name="factory" id="factory">
        <option value="">Select</option>
        <%=getFactory("OPTION","",factorywhereinside,"","")%>
      </select></td>
      <td>Section</td>
      <td><select name="section" id="section">
        <option value="">Select</option>
        <%=getSection("OPTION","",factorywhereinside,"","")%>
      </select></td>
      <td>Production   Time</td>
      <td>From
        <input name="jfromdate" type="text" id="jfromdate" value="<%=dateadd("d",-1,date())%>" size="10">
          <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.jfromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
          <input name="jfromhour" type="text" id="jfromhour" value="08:00" size="5">
        &nbsp;to
        <input name="jtodate" type="text" id="jtodate" value="<%=date()%>" size="10">
        <script language=JavaScript type=text/javascript>function calendar4Callback(date, month, year)
	{
	document.all.jtodate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
                        </script>
        &nbsp;
        <input name="jtohour" type="text" id="jtohour" value="08:00" size="5"></td>
    </tr>
    <tr>
      <td height="20" colspan="6"><input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Generate Family Output Report at once"></td>
    </tr>
  </form>
  <%end if%>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="47%">Browse Family Output Report </td>
        <td width="53%"><div align="right">User:
          <% =session("User") %>
        </div></td>
      </tr>
    </table></td>
</tr>
		<%if admin=true then%>
          <form action="/Reports/Process/FamilyFamily Output/GenerateFamilyFamily Output.asp" method="post" name="form2" target="_self" onSubmit="return form2check()">
		  </form>
		  <%end if%><tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><div align="center">Delete</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Report Time </div></td>
  <td class="t-t-Borrow"><div align="center">Report Name</div></td>
  <td class="t-t-Borrow"><div align="center">From Time </div></td>
  <td class="t-t-Borrow"><div align="center">To Time </div></td>
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
    <td><div align="center"><span class="red" style="cursor:hand" onClick="javascript:location.href='DeleteFamily OutputReport.asp?id=<%=rs("NID")%>&output_name=<%=rs("FAMILY_OUTPUT_NAME")%>&path=<%=path%>&query=<%=query%>'">Delete</span></div></td>
    <td height="20"><div align="center"><% =formatdate(rs("REPORT_TIME"),application("longdateformat"))%></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('FamilyOutputReport.asp?family_output_id=<%=rs("NID")%>&family_output_name=<%=rs("FAMILY_OUTPUT_NAME")%>&report_time=<%=rs("REPORT_TIME")%>&from_time=<%=rs("FROM_TIME")%>&to_time=<%=rs("TO_TIME")%>&factory_id=<%=rs("FACTORY_ID")%>&section_id=<%=rs("SECTION_ID")%>')"><%= rs("FAMILY_OUTPUT_NAME") %></span>&nbsp;</div></td>
    <td><div align="center"><%= rs("FROM_TIME") %></div></td>
    <td><div align="center"><%= rs("TO_TIME") %></div></td>
    <td><div align="center"><%= rs("CREATOR_CODE") %></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="7"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->