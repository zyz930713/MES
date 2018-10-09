<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request("factory")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
line=trim(request("line"))
code=trim(request("code"))
storetype=trim(request("storetype"))
checkstatus=trim(request("checkstatus"))
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
todate=dateadd("d",date(),1)
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
if checkstatus<>"" then
where=where&" and J.RETEST_CHECK_STATUS='"&checkstatus&"'"
end if
if fromdate<>"" then
where=where&" and to_date(J.STORE_TIME)>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and to_date(J.STORE_TIME)<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&factory="&factory&"&jobnumber="&jobnumber&"&partnumber="&partnumber&"&line="&line&"&code="&code&"&storetype="&storetype&"&checkstatus="&checkstatus&"&fromdate="&fromdate&"&todate="&todate
pagename="/Store/StoreRecords.asp"
SQL="select FACTORY_NAME from FACTORY where NID='"&factory&"'"
rs.open SQL,conn,1,3
if not rs.eof then
factory_name=rs("FACTORY_NAME")
end if
rs.close
SQL="select J.* from JOB_MASTER_STORE_PRE J where J.FACTORY_ID='"&factory&"' "&where&order

'response.write sql

session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ѯ<%=factory_name%>����嵥</title>
<link href="/CSS/List.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<form name="form1" method="post" action="/Store/StoreRecords.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-c-green"><span id="inner_Search">��ѯ<%=factory_name%>����嵥</span></td>
  </tr>
  <tr>
    <td height="20"><div align="center">������</div></td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td><div align="center">�ͺ�</div></td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td><div align="center">�߱�</div></td>
    <td><input name="line" type="text" id="line" value="<%=line%>"></td>
    <td><div align="center">���ʱ��</div></td>
    <td>��
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;��
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
    <td height="20"><div align="center">�����</div></td>
    <td height="20"><input name="code" type="text" id="code" value="<%=code%>"></td>
    <td><div align="center">�������</div></td>
    <td><select name="storetype" id="storetype">
      <option value="">--ѡ������--</option>
      <option value="N">��������</option>
      <option value="R">���޹���</option>
    </select>    </td>
    <td colspan="4"><input name="factory" type="hidden" id="factory" value="<%=factory%>">
    <input name="Submit" type="submit" id="Submit" value="��ѯ"></td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-green">KE����嵥</td>
</tr>

<tr>
  <td height="20" colspan="11"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">����</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">������<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">�ͺ�<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.PART_NUMBER_TAG&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">�߱�</div></td>
  <td class="t-t-Borrow"><div align="center">�����</div></td>
  <td class="t-t-Borrow"><div align="center">��������</span></div></td>
  <td class="t-t-Borrow"><div align="center">�������</span></div></td>
  <td class="t-t-Borrow"><div align="center">��ʱ����</span></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">���ʱ��<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.STORE_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center">�������</div></td>
  <td class="t-t-Borrow"><div align="center">�ӹ���</div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
	<tr>
	  <td height="20"><div align="center"><%=i%></div></td>
	  	<%if retestchecker=true then%>
		<%end if%>
		<td height="20"><div align="center"><%= rs("JOB_NUMBER")%></div></td>
		<td height="20"><div align="center"><%= rs("PART_NUMBER_TAG")%></div></td>
		<td><div align="center"><%=rs("LINE_NAME")%> </div></td>
		<td><div align="center"><%=rs("STORE_CODE")%></div></td>
		<td><div align="center"><%=rs("INPUT_QUANTITY")%>&nbsp;</div></td>
		<td><div align="center"><%=rs("STORE_QUANTITY")%></div></td>
		<td><div align="center"><%=formatpercent(csng(rs("YIELD")),2,-1)%></div></td>
		<td><div align="center"><% =formatdate(rs("STORE_TIME"),application("longdateformat"))%></div></td>
		<td><div align="center"><%=rs("STORE_TYPE")%></div></td>
		<td><div align="center"><%=formatlongstring(rs("SUB_JOB_NUMBERS"),"<br>",50)%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="11"><div align="center"><span id="inner_NoRecords">�޼�¼</span></div></td>
  </tr>
<%end if
rs.close%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->