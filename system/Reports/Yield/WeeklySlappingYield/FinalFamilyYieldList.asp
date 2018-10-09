<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/FinalYield/FinalYieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
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
order=" order by FF.REPORT_TIME desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if reportname<>"" then
where=where&" and FF.FINAL_FAMILYYIELD_NAME='"&reportname&"'"
end if
if fromdate<>"" then
where=where&" and FF.REPORT_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and FF.REPORT_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&reportname="&reportname&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/FinalYield/FinalFamilyYield/FinalFamilyYieldList.asp"
SQL="select FF.*,F.FACTORY_NAME,U.USER_NAME from FINAL_FAMILYYIELD_LIST FF inner join FACTORY F on FF.FACTORY_ID=F.NID left join USERS U on FF.CREATOR_CODE=U.USER_CODE where FF.NID is not null "&where&order
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
	}
}
</script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page()">
<form name="form1" method="post" action="/Reports/FinalYield/FinalFamilyYield/FinalFamilyYieldList.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-c-greenCopy"><span>Search Final Family Yield Reports </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Report Name </span> </td>
    <td height="20"><input name="reportname" type="text" id="reportname" value="<%=reportname%>"></td>
    <td>Report Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate2" value="<%=fromdate%>" size="10">
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
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#006600" bordercolordark="#FFFFFF">
  <%if admin=true then%>
  <form action="/Reports/FinalYield/FinalFamilyYield/GenerateFinalFamilyYield.asp" method="post" name="form2" target="_self" onSubmit="return form2check()">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy">Generating conditions</td>
    </tr>
  <tr>
    <td class="t-b-blue">Factory: </td>
    <td><select name="factory">
	<option value="">Select</option>
	<%=getFactory("OPTION","",factorywhereinside,"","")%>
	</select></td>
    <td class="t-b-blue">By Line</td>
    <td><input name="byline" type="checkbox" id="byline" value="1"></td>
    <td class="t-b-blue">Job Close Time</td>
    <td>From
      <input name="jfromdate" type="text" id="jfromdate" value="<%=dateadd("d",-7,date())%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.jfromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
      <input name="jfromhour" type="text" id="jfromhour" value="00:00" size="5">
      &nbsp;to
<input name="jtodate" type="text" id="jtodate" value="<%=date()%>" size="10">
<script language=JavaScript type=text/javascript>
	</script>
<script language=JavaScript type=text/javascript>function calendar4Callback(date, month, year)
	{
	document.all.jtodate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
                        </script>
&nbsp; <input name="jtohour" type="text" id="jtohour" value="00:00" size="5"></td>
  </tr>
  <tr>
    <td height="20" colspan="6"><input name="Generate" type="submit" class="t-b-Yellow" id="Generate" value="Generate Slapping Report at once" onClick="javascript:if(byline.checked){document.form2.action='/Reports/FinalYield/FinalFamilyYield/GenerateFinalFamilyYieldByLine.asp'}else{document.form2.action='/Reports/FinalYield/FinalFamilyYield/GenerateFinalFamilyYield.asp'}"></td>
    </tr>
  </form>
  <%end if%>
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="10" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="48%">Browse Final Family Yield Report </td>
        <td width="52%"><div align="right">User:
          <% =session("User") %>
        </div></td>
      </tr>
    </table></td>
  </tr>
  
    
  <tr>
    <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
	<%if admin=true then%>
    <td class="t-t-Borrow"><div align="center">Delete</div></td>
	<%end if%>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=FF.REPORT_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Time <img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=FF.REPORT_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=FF.FINAL_FAMILYYIELD_NAME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Report Name<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=FF.FINAL_FAMILYYIELD_NAME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center">Report Chart</div></td>
    <td class="t-t-Borrow"><div align="center">Period</div></td>
    <td class="t-t-Borrow"><div align="center">Chart Week </div></td>
    <td class="t-t-Borrow"><div align="center">Factory</div></td>
    <td class="t-t-Borrow"><div align="center">By Line </div></td>
    <td class="t-t-Borrow"><div align="center">Creator</div></td>
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
      </div></td>
	<%if admin=true then%>
    <td><div align="center"><%if session("code")=rs("CREATOR_CODE") then%><span class="red" style="cursor:hand" onClick="javascript:location.href='DeleteFinalFamilyYieldReport.asp?id=<%=rs("NID")%>&finalfamily_name=<%=rs("FINAL_FAMILYYIELD_NAME")%>&path=<%=path%>&query=<%=query%>'"><img src="/Images/IconDelete.gif"></span><%else%><img src="/Images/IconDelete_No.gif"><%end if%></div></td>
	<%end if%>
    <td height="20">
      <div align="center">
        <% =formatdate(rs("REPORT_TIME"),application("longdateformat"))%>
      </div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('FinalFamilyYieldReport<%if rs("BY_LINE")="1" then%>ByLine<%end if%>.asp?finalfamily_id=<%=rs("NID")%>&finalfamily_name=<%=rs("FINAL_FAMILYYIELD_NAME")%>&finalfamily_report_time=<%=rs("REPORT_TIME")%>&from_time=<%=rs("FROM_TIME")%>&to_time=<%=rs("TO_TIME")%>&factory_id=<%=rs("FACTORY_ID")%>')"><%= rs("FINAL_FAMILYYIELD_NAME") %></span>&nbsp;</div></td>
    <td><div align="center"><%if rs("CHART_WEEK")<>"" then%><span class="red" style="cursor:hand" onClick="javascript:window.open('http://<%=application("ChartServer")%>/KES1_Barcode/FinalFamilyYieldChart<%if rs("BY_LINE")="1" then%>ByLine<%end if%>.asp?factory=<%=rs("FACTORY_ID")%>&yearnumber=<%=rs("CHART_YEAR")%>&fromww=1&toww=<%=datepart("ww",rs("TO_TIME"))%>&factory_id=<%=rs("FACTORY_ID")%>')"><img src="/Images/IconReportChart.gif" width="22" height="22"></span><%else%>&nbsp;<%end if%></div></td>
    <td><div align="center"><% =formatdate(rs("FROM_TIME"),application("longdateformat"))%> to <% =formatdate(rs("TO_TIME"),application("longdateformat")) %> </div></td>
    <td><div align="center">
      <% =rs("CHART_WEEK") %>
    &nbsp;of 
    <% =rs("CHART_YEAR") %>
    -
    <% =rs("CHART_MONTH") %>
    </div></td>
    <td><div align="center"><% =rs("FACTORY_NAME") %></div></td>
    <td><div align="center"><%if rs("BY_LINE")="1" then%><img src="/Images/Yes.gif" width="16" height="16"><%end if%>&nbsp;</div></td>
    <td><div align="center"><% =rs("USER_NAME") %></div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="10"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->