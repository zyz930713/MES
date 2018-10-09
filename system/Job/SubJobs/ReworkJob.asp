<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
factory=trim(request("factory"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by REWORK_JOB_NUMBER desc,CREATE_TIME desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and REWORK_JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if line<>"" then
where=where&" and lower(LINE_NAME)='"&lcase(line)&"'"
end if
if factory<>"" then
where=where&" and P.FACTORY_ID='"&factory&"'"
end if
if fromdate<>"" then
where=where&" and CREATE_TIME>='"&fromdate&"'"
else
fromdate=dateadd("d",-7,date())
where=where&" and CREATE_TIME>='"&fromdate&"'"
end if
if todate<>"" then
where=where&" and CREATE_TIME<='"&todate&"'"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&factory="&factory&"&jobstatus="&jobstatus&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Job/ReworkJob.asp"
FactoryRight "P."
if session("code")<>"1194" then
SQL="select * from LINE_REWORK where LEADER='"&session("code")&"'"&where&order
else
SQL="select * from LINE_REWORK where 1=1"&where&order
end if
rsPR.open SQL,connPR,1,3
session("SQL")=SQL
%>
<!--#include virtual="/Components/PO_PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<!--#include virtual="/Language/Job/SubJobs/Lan_ReworkJob.asp" -->
</head>

<body onLoad="language();language_page();">
<form action="/Job/SubJobs/ReworkJob.asp" method="get" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-b-midautumn"><span id="inner_Search"></span></td>
    </tr>
    <tr>
      <td height="20"><span class="style1"><span id="inner_SearchJobNumber"></span></td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
      <td><span id="inner_SearchPartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
      <td><span id="inner_SearchLineName"></span></td>
      <td><input name="line" type="text" id="line" value="<%=line%>"></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchFactory"></span></td>
      <td height="20"><select name="factory" id="factory">
        <option value="">Factory</option>
        <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
        <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
      </select></td>
      <td><span id="inner_SearchJobStartTime"></span></td>
      <td colspan="3"><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp;<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursPRor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="48%"><span id="inner_Browse"></span></td>
            <td width="52%"><div align="right"><span style="cursPRor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Job_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
    <tr>
      <td height="20" colspan="7"><!--#include virtual="/Components/PO_PageSplit.asp" --></td>
    </tr>
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursPRor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=CREATE_TIME&ordertype=asc&pagenum=<%=firsPRtpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"><span id="inner_CreateTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursPRor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=CREATE_TIME&ordertype=desc&pagenum=<%=firsPRtpage%>&pagesize_a=<%=pagesize_s%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&sheetnumber=<%=sheetnumber%>&line=<%=line%>&factory=<%=factory%>&jobstatus=<%=jobstatus%>&currentstation=<%=currentstation%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_PrintTimes"></span></div></td>
    </tr>
<%
i=1
if not rsPR.eof then
rsPR.absolutepage=session("strpagenum")
while not rsPR.eof and i<=rsPR.pagesize 
%>
    <tr <%if rsPR("PRINT_TIMES")=0 then%>class="strongred"<%end if%>>
      <td height="20"><div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
</div></td>
      <td height="20"><div align="center" class="d_link"><strong><a href="/Job/SubJobs/ReworkJobDetail.asp?jobnumber=<%=replace(rsPR("REWORK_JOB_NUMBER"),"+","$")%>&part_number_tag=<%=rsPR("PART_NUMBER_TAG")%>&line_name=<%=rsPR("LINE_NAME")%>" target="_blank"><%=rsPR("REWORK_JOB_NUMBER")%></a></strong></div></td>
      <td><div align="center"><%=rsPR("PART_NUMBER_TAG")%></div></td>
      <td><div align="center"><%=rsPR("LINE_NAME")%></div></td>
      <td height="20"><div align="center"><% =formatdate(rsPR("CREATE_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
      <td><div align="center"><%=rsPR("QUANTITY")%>&nbsp;</div></td>
      <td><div align="center"><%=rsPR("PRINT_TIMES")%>&nbsp;</div></td>
    </tr>
<%
i=i+1
rsPR.movenext
wend
else
%>
    <tr>
      <td height="20" colspan="7"><div align="center"><span id="inner_Records"></span>No Records</div></td>
    </tr>
<%end if
rsPR.close%>
</table>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
