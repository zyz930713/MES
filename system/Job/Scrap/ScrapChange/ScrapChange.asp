<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Scrap/ScrapCheck.asp" -->
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
changecode=trim(request("changecode"))
changetype=trim(request("changetype"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" then
fromdate=date()
end if
if todate="" then
todate=date()
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
if changecode<>"" then
where=where&" and J.CHANGE_CODE like '%"&changecode&"%'"
end if
if changetype<>"" then
where=where&" and J.CHANGE_TYPE='"&changetype&"'"
end if
if fromdate<>"" then
where=where&" and to_date(J.CHANGE_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.CHANGE_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&changecode="&changecode&"&changetype="&changetype&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Scrap/ScrapChange/ScrapChange.asp"
FactoryRight "J."
SQL="select J.*,PF.NID from JOB_MASTER_SCRAP_CHANGE J left join PROFILE_FORM PF on J.PROFILE_FORM_ID=PF.NID inner join FACTORY F on J.FACTORY_ID=F.NID where J.JOB_NUMBER is not null "&where&factorywhereoutsideand&order
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
<!--#include virtual="/Language/Reports/Scrap/ScrapChange/Lan_ScrapChange.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post" action="/Job/Scrap/ScrapChange/ScrapChange.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchJobNumber"></span></td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td><span id="inner_SearchPartNumber"></span></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td><span id="inner_SearchLine"></span></td>
    <td><input name="line" type="text" id="line" value="<%=line%>"></td>
    <td><span id="inner_SearchChangeTime"></span></td>
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
    <td height="20"><span id="inner_SearchChangeCode"></span></td>
    <td height="20"><input name="changecode" type="text" id="changecode" value="<%=changecode%>"></td>
    <td><span id="inner_SearchChangeType"></span></td>
    <td><select name="changetype" id="changetype">
      <option value="">--Select Type--</option>
      <option value="1">Edit</option>
      <option value="2">Delete</option>
    </select>    </td>
    <td colspan="4"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_ChangeType"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_ChangeCode"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_ChangeReason"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_ChangeTime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_OldScrapCode"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_OldScrapQuantity"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NewScrapQuantity"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.SCRAP_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_OldScrapTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.SCRAP_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_OldNote"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_FormID"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
	  <td height="20"><div align="center">
	  <%if rs("CHANGE_TYPE")="1" then
	  change_type="ÐÞ¸Ä"
	  else
	  change_type="É¾³ý"
	  end if
	  %><%=change_type%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("CHANGE_CODE")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("CHANGE_REASON")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("CHANGE_TIME")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
	  <td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
	  <td height="20"><div align="center"><%=rs("LINE_NAME")%> </div></td>
	  <td height="20"><div align="center"><%=rs("OLD_SCRAP_CODE")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("OLD_SCRAP_QUANTITY")%></div></td>
	  <td height="20"><div align="center"><%=rs("NEW_SCRAP_QUANTITY")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("OLD_SCRAP_TIME")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("OLD_NOTE")%>&nbsp;</div></td>
	  <td height="20"><div align="center"><%=rs("NID")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="14"><div align="center"><span id="inner_NoRecords">No Records</span></div></td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->