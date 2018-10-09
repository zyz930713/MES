<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
planer=trim(request("planer"))
progress=request("progress")
fromdate=request("fromdate")
todate=request("todate")
close_fromdate=request("close_fromdate")
close_todate=request("close_todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER "
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" and jobnumber="" and close_fromdate="" and close_todate="" then
'fromdate="2007-1-1"
fromdate=dateadd("d",-7,date())
end if
if todate="" then
'todate=dateadd("d",1,date())
end if

if jobnumber<>"" then
where=where&" and J.JOB_NUMBER like '%"&jobnumber&"%'"
end if
if partnumber<>"" then
where=where&" and lower(J.PART_NUMBER_TAG) like '%"&lcase(partnumber)&"%'"
end if
if line<>"" then
where=where&" and lower(J.LINE_NAME) like '%"&lcase(line)&"%'"
end if
if planer<>"" then
where=where&" and lower(J.CREATE_NAME) like '%"&lcase(planer)&"%'"
end if
if progress="0" then
where=where&" and J.STORE_STATUS=0"
elseif progress="1" then
where=where&" and J.STORE_STATUS=1"
end if
if fromdate<>"" then
where=where&" and to_date(J.INPUT_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.INPUT_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
if close_fromdate<>"" then
where=where&" and J.COMPLETE_DATE>=to_date('"&close_fromdate&"','yyyy-mm-dd')"
end if
if close_todate<>"" then
where=where&" and J.COMPLETE_DATE<=to_date('"&close_todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&planer="&planer&"&progress="&progress&"&fromdate="&fromdate&"&todate="&todate&"&close_fromdate="&close_fromdate&"&close_todate="&close_todate
pagename="/Job/Job/Job.asp"
FactoryRight "J."
SQL="select J.* from JOB_MASTER J left join FACTORY F on J.FACTORY_ID=F.NID where J.JOB_NUMBER is not null "&where&factorywhereoutsideand&order
'response.Write(SQL)
session("SQL")=SQL
session("SQLJob")="select J.* from JOB_MASTER J where J.JOB_NUMBER is not null "&where&factorywhereoutsideand&" order by J.JOB_NUMBER"

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
<!--#include virtual="/Language/Job/Job/Lan_Job.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="get" action="/Job/Job/Job.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">      </td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td><span id="inner_SearchLine"></span></td>
    <td><input name="line" type="text" id="line" value="<%=line%>" ></td>
    <td><span id="inner_SearchCreateTime"></span></td>
    <td><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
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
&nbsp; </td>
    </tr>
  <tr>
    <td height="20"><span id="inner_SearchPlaner"></span></td>
    <td height="20"><input name="planer" type="text" id="planer" value="<%=planer%>"></td>
    <td><span id="inner_SearchProgress"></span></td>
    <td><select name="progress" id="progress">
      <option value="">--Select Progress--</option>
      <option value="1" <%if progress="1" then%>selected<%end if%>>Finished</option>
      <option value="0" <%if progress="0" then%>selected<%end if%>>Progressing</option>
    </select></td>
    <td><span id="inner_SearchJobCloseTime"></span></td>
    <td><input name="close_fromdate" type="text" id="close_fromdate" value="<%=close_fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar3Callback(date, month, year)
	{
	document.all.close_fromdate.value=year + '-' + month + '-' + date
	}
    calendar3 = new dynCalendar('calendar3', 'calendar3Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="close_todate" type="text" id="close_todate" value="<%=close_todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar4Callback(date, month, year)
	{
	document.all.close_todate.value=year + '-' + month + '-' + date
	}
    calendar4 = new dynCalendar('calendar4', 'calendar4Callback');
    </script></td>
    <td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right">
		  <%if DBA=true then%>
          <input name="Compare" type="submit" class="t-b-Yellow" id="Compare" value="Compare" onClick="window.open('/Job/Scrap/JobScrap/JobScrap.asp','_blank')">
		  <%end if%>
          <span style="cursor:hand" title="Job Export" onClick="javascript:window.open('Job_Export.asp?fromdate=<%=fromdate%>&todate=<%=todate%>')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>
		  &nbsp;
		  <span style="cursor:hand" title="SubJob Station Export" onClick="javascript:window.open('Job_SubJob_Station_Export.asp?fromdate=<%=fromdate%>&todate=<%=todate%>')"><img src="/Images/IconReportTable.gif" width="22" height="22"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="18"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Progress"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Stock"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Planer"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StartQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_AssemblyGood"></span></div></td>
<!--  <td class="t-t-Borrow"><div align="center"><span id="inner_AssemblyYield"></span></div></td>
-->  <td class="t-t-Borrow"><div align="center"><span id="inner_LineLost"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_FinalGood"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_FinalYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_ConfirmGood"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_FinalScrap"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_FinalRemain"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.FIRST_STORE_TIME	&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_InputTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.FIRST_STORE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=recordsize%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_LastTime"></span> </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_CycleTime"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
		<td height="20"><div align="center"><a href="/Job/SubJobs/Job.asp?jobnumber=<%=rs("JOB_NUMBER")%>&fromdate=<%=fromdate%>" target="_blank"><%=rs("JOB_NUMBER")%></a></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
		<td><div align="center"><%= rs("LINE_NAME")%> </div></td>
		<td><div align="center">
		<%
		  session("aerror")=rs("JOB_NUMBER")
		  if csng(rs("START_QUANTITY"))<>0 then
		  percent=round((csng(rs("FINAL_GOOD_QUANTITY"))+csng(rs("FINAL_SCRAP_QUANTITY")))/csng(rs("START_QUANTITY")),4)
		  else
		  percent=0
		  end if
		  if isnull(rs("INPUT_TIME")) or rs("INPUT_TIME")="" then
		  progresscolor="#CC9900"
		  else
		  	if percent=1 then
		  	progresscolor="#099E09"
		  	else
		  	progresscolor="#000099"
		  	end if
		  end if%>
		  <table width="60" border="1" cellpadding="0" cellspacing="0" bordercolor="<%=progresscolor%>">
            <tr>
              <td><div align="left"><%if percent=1 then%><img src="/Images/JobProgressGreen.gif" width="60" height="8"><%else%><img src="/Images/JobProgress.gif" width="<%=percent*60%>" height="8"><%end if%></div></td>
            </tr>
          </table>
		<%=formatpercent(percent,2,-1)%></div></td>
		<td><div align="center">
	      <%if rs("STORE_STATUS")="0" then%>
	      Open
	      <%else%>
	      Close
	      <%end if%><%if DBA=true and rs("STORE_STATUS")="0" then%><span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('JobStockClose.asp?jobnumber=<%=rs("JOB_NUMBER")%>&path=<%=path%>&query=<%=query%>','_self')">C</span><%end if%>
		</div></td>
		<td><div align="center"><%= rs("CREATE_NAME")%><%if DBA=true then%><span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('JobInfoUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&path=<%=path%>&query=<%=query%>','_self')">U</span><%end if%>&nbsp;</div></td>
		<td><div align="center"><%= rs("START_QUANTITY")%></div></td>
		<td><div align="center"><%= rs("ASSEMBLY_GOOD_QUANTITY")%></div></td>
<!--		<td><div align="center"><%'= formatpercent(csng(rs("ASSEMBLY_YIELD")),2,-1)%></div></td>
-->		<td><div align="center"><a href="/Job/SubJobs/DefectCodeDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("DEFECTCODE_QUANTITY")%></a></div></td>
		<td><div align="center"><a href="/Job/Store/JobStore/JobStoreDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("FINAL_GOOD_QUANTITY")%></a></div></td>
		<%
			yield=csng(rs("ASSEMBLY_GOOD_QUANTITY"))/csng(rs("START_QUANTITY"))
		%>
		<td><div align="center"><%= formatpercent(yield,2,-1)%></div></td>
 
		<td><div align="center"><%= rs("CONFIRM_GOOD_QUANTITY")%></div></td>
        <td><div align="center"><%= rs("FINAL_SCRAP_QUANTITY")%></div></td>
		<td><div align="center"><%= csng(rs("START_QUANTITY"))-csng(rs("FINAL_GOOD_QUANTITY"))-csng(rs("FINAL_SCRAP_QUANTITY"))%></div></td>
		<td><div align="center">
		  <% =formatdate(rs("INPUT_TIME"),application("longdateformat"))%>
	    &nbsp;</div></td>
		<td><div align="center"><% =formatdate(rs("LAST_UPDATE_TIME"),application("longdateformat"))%>&nbsp;</div></td>
		<td><div align="center"><%= formatnumber(csng(rs("AVERAGE_CYCLE_TIME"))/60,2,-1)%>&nbsp;h</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="18"><div align="center"><span id="inner_Records"></span> </div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->