<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/Store/StoreNoYieldCheck.asp" -->
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
code=trim(request("code"))
storetype=trim(request("storetype"))
NoYieldstatus=trim(request("NoYieldstatus"))
fromdate=request("fromdate")
todate=request("todate")
fromhour=request("fromhour")
tohour=request("tohour")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""

if fromdate="" then
fromdate=dateadd("d",-7,date())
end if
if fromhour="" then
fromhour="00:00"
end if
if todate="" then
todate=date()
end if
if tohour="" then
tohour="00:00"
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
if code<>"" then
where=where&" and J.STORE_CODE like '%"&code&"%'"
end if
if storetype<>"" then
where=where&" and J.STORE_TYPE='"&storetype&"'"
end if
if NoYieldstatus<>"" then
where=where&" and J.NO_YIELD='"&NoYieldstatus&"'"
end if
if fromdate<>"" and fromhour<>"" then
where=where&" and to_date(J.STORE_TIME)>=to_date('"&fromdate&" "&fromhour&"','yyyy-mm-dd hh24:mi:ss')"
end if
if todate<>"" and tohour<>"" then
where=where&" and to_date(J.STORE_TIME)<=to_date('"&todate&" "&tohour&"','yyyy-mm-dd hh24:mi:ss')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&code="&code&"&storetype="&storetype&"&NoYieldstatus="&NoYieldstatus&"&fromdate="&fromdate&"&todate="&todate&"&fromhour="&fromhour&"&tohour="&tohour
pagename="/Reports/Store/StoreNoYield/StoreNoYield.asp"
FactoryRight "J."
SQL="select 1,J.* from JOB_MASTER_STORE J inner join FACTORY F on J.FACTORY_ID=F.NID where J.NID is not null"&where&factorywhereoutsideand&order
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
<!--#include virtual="/Job/Store/StoreNoYield/NoYieldCheck.asp" -->
<!--#include virtual="/Language/Reports/Store/StoreNoYield/Lan_StoreNoYield.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post" action="/Job/Store/StoreNoYield/StoreNoYield.asp">
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
    <td><span id="inner_SearchStoreTime"></span></td>
<td><span id="inner_SearchFrom"></span>
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
      <input name="fromhour" type="text" id="fromhour" value="<%if fromhour<>"" then%><%=fromhour%><%else%>00:00<%end if%>" size="5">
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar4Callback');
                        </script>
&nbsp;
<input name="tohour" type="text" id="tohour" value="<%if tohour<>"" then%><%=tohour%><%else%>00:00<%end if%>" size="5"></td>
</tr>
  <tr>
    <td height="20"><span id="inner_SearchCode"></span></td>
    <td height="20"><input name="code" type="text" id="code" value="<%=code%>"></td>
    <td><span id="inner_SearchStoreType"></span></td>
<td><select name="storetype" id="storetype">
      <option value="">--Select Type--</option>
      <option value="N" <%if storetype="N" then%>selected<%end if%>>Normal</option>
      <option value="R" <%if storetype="N" then%>selected<%end if%>>Rework</option>
    </select>    </td>
<td><span id="inner_SearchNoYieldStatus"></span></td>
<td><select name="NoYieldstatus" id="NoYieldstatus">
      <option value="">--Select Status--</option>
      <option value="1" <%if NoYieldstatus="1" then%>selected<%end if%>>Checked</option>
      <option value="0" <%if NoYieldstatus="0" then%>selected<%end if%>>Unchecked</option>
    </select>    </td>
<td colspan="2"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
</tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="13" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
</tr>
<tr>
  <td height="20" colspan="13">Sort by:
    <input name="ByJob" type="radio" value="1" onClick="javascript:location.href='StoreNoYieldbyJobNumber.asp'">
Job Number
<input name="ByPart" type="radio" value="1" onClick="javascript:location.href='StoreNoYieldbyPartNumber.asp'">
Part Number </td>
</tr>
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if NoYieldchecker=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_NoYield"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Code"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_InputQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StoreQuantity"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_OntimeYield"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_StoreTime"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_StoreType"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Note"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_SubJobs"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
<%if NoYieldchecker=true then%>
	  <td><div align="center">
		  <input name="NoYieldcheck<%=i%>" type="checkbox" id="NoYieldcheck<%=i%>" value="<%=rs("NID")%>" <%if rs("NO_YIELD")="1" then%>checked<%end if%> onClick="NoYieldcheck(this,this.value)">
    </div></td>
        <%end if%>
        <td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
        <td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
        <td><div align="center"><%=rs("LINE_NAME")%> </div></td>
        <td><div align="center"><%=rs("STORE_CODE")%></div></td>
        <td><div align="center"><%=rs("INPUT_QUANTITY")%>&nbsp;</div></td>
        <td><div align="center"><%=rs("STORE_QUANTITY")%></div></td>
        <td><div align="center"><%=formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
        <td><div align="center">
          <% =formatdate(rs("STORE_TIME"),application("longdateformat"))%>
        </div></td>
        <td><div align="center"><%=rs("STORE_TYPE")%></div></td>
        <td><div align="center"><%=rs("NOTE")%>&nbsp;</div></td>
        <td><div align="center"><%=formatlongstring(rs("SUB_JOB_NUMBERS"),"<br>",50)%>&nbsp;</div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="13"><div align="center"><span id="inner_NoRecords">No Records</span></div></td>
</tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->