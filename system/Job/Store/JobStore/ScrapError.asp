<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
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
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""

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
where=where&" and FINAL_INPUT_QUANTITY<>START_QUANTITY"
elseif progress="1" then
where=where&" and FINAL_INPUT_QUANTITY=START_QUANTITY"
end if
if fromdate<>"" then
where=where&" and to_date(J.FIRST_STORE_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.FIRST_STORE_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&planer="&planer&"&progress="&progress&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Store/JobStore/ScrapError.asp"
FactoryRight "J."
SQL="select 1,J.* from JOB_MASTER J inner join FACTORY F on J.FACTORY_ID=F.NID where (J.EXCEED_ERROR='1' or START_ERROR='1') "&where&factorywhereoutsideand&order
'response.Write(SQL)
session(SQL)=SQL
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
<!--#include virtual="/Language/Reports/Store/JobStore/Lan_ScrapError.asp" -->
</head>

<body onLoad="language();language_page()">
<form name="form1" method="post" action="/Job/Store/JobStore/ScrapError.asp">
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
    <td><input name="line" type="text" id="line" value="<%=line%>" size="6"></td>
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
    <td colspan="4"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="15" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="15" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
          <% =session("User") %></td>
        <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20" colspan="15"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_JobNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"><span id="inner_PartNumber"></span><img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_Planer"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StoreStatus"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StartError"></span></div></td>
	<td class="t-t-Borrow"><div align="center"><span id="inner_ExceedError"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_StartQuantity"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_CompleteQuantity"></span></div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_ScrapQuantity"></span></div></td>
    <td class="t-b-newyear"><div align="center"><span id="inner_ERPStartQuantity"></span></div></td>
    <td class="t-b-newyear"><div align="center"><span id="inner_ERPCompleteQuantity"></span></div></td>
    <td class="t-b-newyear"><div align="center"><span id="inner_ERPScrapQuantity"></span> </div></td>
    <td class="t-t-Borrow"><div align="center"><span id="inner_ERPUpdateTime"></span></div></td>
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20"><div align="center"><%=i%></div></td>
    <td height="20"><div align="center"><a href="/Job/Store/JobStore/JobStoreDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%=rs("JOB_NUMBER")%></a></div></td>
    <td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("LINE_NAME")%> </div></td>
    <td><div align="center"><%= rs("CREATE_NAME")%>
            <%if DBA=true then%>
      <span class="red" style="cursor:hand" title="Update Quantity" onClick="window.open('JobInfoUpdate.asp?jobnumber=<%=rs("JOB_NUMBER")%>&path=<%=path%>&query=<%=query%>','_self')">U</span>
      <%end if%>
      &nbsp;</div></td>
    <td><div align="center">
	<%if rs("STORE_STATUS")="0" then
	store_status="Going|进行中"
	else
	store_status="Finished|已完成"
	end if
	arrstore_status=split(store_status,"|")
	%><% =arrstore_status(session("language"))%></div></td>
    <td><div align="center"><%if rs("START_ERROR")="1" then%><img src="/Images/Yes.gif"><%else%>&nbsp;<%end if%></div></td>
    <td><div align="center"><%if rs("EXCEED_ERROR")="1" then%><img src="/Images/Yes.gif"><%else%>&nbsp;<%end if%></div></td>
    <td><div align="center"><%= rs("START_QUANTITY")%></div></td>
    <td><div align="center"><%= rs("FINAL_GOOD_QUANTITY")%></div></td>
    <td><div align="center"><%= cint(rs("START_QUANTITY"))-cint(rs("FINAL_GOOD_QUANTITY"))-cint(rs("ERP_SCRAP_QUANTITY"))%></div></td>
    <td><div align="center"><%= rs("ERP_START_QUANTITY")%></div></td>
    <td><div align="center"><%= rs("ERP_COMPLETE_QUANTITY")%></div></td>
    <td><div align="center"><% =rs("ERP_SCRAP_QUANTITY")%></div></td>
    <td><div align="center">
      <% =formatdate(rs("ERP_UPDATE_TIME"),application("shortdateformat"))%>
      &nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="15"><div align="center">No Records </div></td>
  </tr>
  <%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->