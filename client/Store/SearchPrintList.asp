<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=trim(request.ServerVariables("PATH_INFO"))
query=trim(request.ServerVariables("QUERY_STRING"))
query=replace(query,"&","*")
pagename="SearchPrintList.asp"
jobnumber=trim(request.QueryString("jobnumber"))
SQL="select NID from JOB_MASTER_STORE_PRE where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	while not rs.eof
		StoreNID=StoreNID&rs("NID")&","
	rs.movenext
	wend
end if
rs.close
if StoreNID<>"" then
	StoreNID=left(StoreNID,len(StoreNID)-1)
	a_StoreNID=split(StoreNID,",")
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>查询打印清单</title>
<link href="/CSS/List.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="800" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
  <td colspan="3" class="today">有工厂属性的</td>
  </tr>
<tr>
    <td class="today"><div align="center">序号</div></td>
    <td class="today"><div align="center">打印清单编号</div></td>
    <td class="today"><div align="center">包括入库工单</div></td>
  </tr>
<%
printids=""
if StoreNID<>"" then
	i=1
	for j=0 to ubound(a_StoreNID)
		SQL="select * from STORE_PRINT where PRINT_MEMBERS like '%"&a_StoreNID(j)&"%'"
		rs.open SQL,conn,1,3
		if not rs.eof then
		while not rs.eof
		if instr(printids,rs("NID"))=0 and i<>0 then
%>
	  <tr>
		<td><div align="center"><%=i%></div></td>
		<td><div align="center"><a href="/Store/PrintStoreRepeatList.asp?printid=<%=rs("NID")%>"><%=rs("NID")%></a></div></td>
		<td><%=formatlongstring(rs("PRINT_MEMBERS"),"<br>",100)%></td>
	  </tr>
<%      i=i+1
		end if
		printids=printids&rs("NID")&","
		rs.movenext
		wend
		end if
		rs.close
	next
else
%>
<tr>
    <td colspan="3"><div align="center">没有记录</div></td>
  </tr>
<%end if%>
</table>
<table width="800" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
    <td colspan="6" class="t-t-Borrow">没有工厂属性的</td>
  </tr>
  <tr>
    <td class="t-t-Borrow"><div align="center">序号</div></td>
    <td class="t-t-Borrow"><div align="center">修改工厂</div></td>
    <td class="t-t-Borrow"><div align="center">工单号</div></td>
    <td class="t-t-Borrow"><div align="center">型号</div></td>
    <td class="t-t-Borrow"><div align="center">包括入库工单</div></td>
    <td class="t-t-Borrow"><div align="center">入库时间</div></td>
  </tr>
  <%
SQL="select * from JOB_MASTER_STORE_PRE where FACTORY_ID not like 'FA%' and STORE_TIME>=to_date('2008-12-1','yyyy-mm-dd hh24:mi:ss')"
rs.open SQL,conn,1,3
if not rs.eof then
i=1
while not rs.eof
%>
  <tr>
    <td><div align="center"><%=i%></div></td>
    <td><select name="factory" id="factory" onChange="javascript:if (this.selectedIndex!=0){location.href='ChangeJobFactory.asp?path=<%=path%>&query=<%=query%>&id=<%=rs("NID")%>&factory='+this.options[this.selectedIndex].value}">
	<option value="">Select</option>
	<%=getFactory("OPTION",null)%>
    </select></td>
    <td><div align="center"><%=rs("JOB_NUMBER")%></div></td>
    <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
    <td><%=formatlongstring(rs("SUB_JOB_NUMBERS"),"<br>",100)%>&nbsp;</td>
    <td><%=rs("STORE_TIME")%></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td colspan="6"><div align="center">没有记录</div></td>
  </tr>
  <%end if
  rs.close%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->