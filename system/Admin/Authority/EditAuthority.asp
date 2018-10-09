<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Authority/AuthorityCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
code=ucase(request.QueryString("code"))
force=request.QueryString("force")
forcefactory=request.QueryString("forcefactory")
set rs1=server.CreateObject("adodb.recordset")
SQL1="select AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID from OPERATORS where CODE='"&code&"'"
rs1.open SQL1,conn,1,3
AUTHORIZED_STATIONS_ID=rs1("AUTHORIZED_STATIONS_ID")
AUTHORIZED_PARTS_ID=rs1("AUTHORIZED_PARTS_ID")
rs1.close
set rs1=nothing
if AUTHORIZED_STATIONS_ID<>"" then
	anid=split(AUTHORIZED_STATIONS_ID,",")
else
	dim anid(0)
end if
if AUTHORIZED_PARTS_ID<>"" then
	apid=split(AUTHORIZED_PARTS_ID,",")
else
	dim apid(0)
end if
if force="1" then
	'SQL="select S.*,F.FACTORY_NAME from STATION S inner join FACTORY F on S.FACTORY_ID=F.NID where F.NID='"&forcefactory&"' order by F.FACTORY_NAME,S.STATION_NAME"
	SQL="select S.*,F.FACTORY_NAME from Station_NEW S inner join FACTORY F on S.FACTORY_ID=F.NID where F.NID='"&forcefactory&"' AND IS_DELETE='0' "
	SQL=SQL+" order by F.FACTORY_NAME,S.STATION_NAME"
else
	FactoryRight "S."
	SQL="select S.*,F.FACTORY_NAME from Station_NEW S inner join FACTORY F on S.FACTORY_ID=F.NID "&factorywhereoutside&" AND IS_DELETE='0' "
	SQL=SQL+" order by F.FACTORY_NAME,S.STATION_NAME"
end if
'response.Write(SQL)
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Admin/Authority/Lan_EditAuthority.asp" -->
</head>

<body onLoad="language()">
<form action="/Admin/Authority/EditAuthority1.asp" method="post" name="checkform" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_Browse"></span> <%=code%> <span id="inner_Browse2"></span></td>
</tr>
<tr>
  <td height="20" colspan="4" class="t-c-greenCopy"><span id="inner_User"></span>:
    <% =session("User") %></td>
  </tr>
<tr>
  <td width="49" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td width="48" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Select"></span></div></td>
  <td width="71" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Fatotory"></span></div></td>
  <td width="768" height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_StaionName"></span></div></td>
  </tr>
<%
if not rs.eof then
i=1
while not rs.eof%>
<tr>
  <td><div align="center"><%=i%></div></td>
  <td><div align="center"><span class="red">
    <input name="nfactory<%=i%>" type="hidden" id="nfactory<%=i%>" value="<%=rs("FACTORY_ID")%>">
    <input name="n_nid<%=i%>" type="hidden" id="n_nid<%=i%>" value="<%=rs("NID")%>">
    <input name="nid<%=i%>" type="checkbox" id="nid<%=i%>" value="1"
	<%
	for l=0 to ubound(anid)
	if rs("NID")=anid(l) then%>checked
	<%
	end if
	next%>>
  </span></div></td>
  <td><div align="center"><%=rs("FACTORY_NAME")%></div></td>
  <td height="20"><div align="left"><%=rs("STATION_NAME")%>&nbsp;(<%=rs("STATION_CHINESE_NAME")%>)</div></td>
</tr>
<%rs.movenext
i=i+1
wend
else%>
  <tr>
    <td height="20" colspan="4">&nbsp;</td>
  </tr>
<%end if
rs.close%>  
</table>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  
  <tr>
    <td width="49" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO2"></span></div></td>
    <td width="48" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Select2"></span></div></td>
    <td width="71" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Fatotory2"></span></div></td>
    <td width="768" height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_PartName"></span></div></td>
  </tr>
<%
if force="1" then
	SQL="select 1,P.*,F.FACTORY_NAME from PART P inner join FACTORY F on P.FACTORY_ID=F.NID where P.ROUTINE_TYPE=0 and F.NID='"&forcefactory&"' "
	SQL=SQL+" AND mother_routing_id IN (SELECT NID FROM ROUTING WHERE IS_DELETE='0') order by F.FACTORY_NAME,P.PART_NUMBER"
else
	FactoryRight "P."
	SQL="select 1,P.*,F.FACTORY_NAME from PART P inner join FACTORY F on P.FACTORY_ID=F.NID where ROUTINE_TYPE=0"
	SQL=SQL+" AND mother_routing_id IN (SELECT NID FROM ROUTING WHERE IS_DELETE='0')" &factorywhereoutsideand&" order by F.FACTORY_NAME,P.PART_NUMBER"
end if

rs.open SQL,conn,1,3
if not rs.eof then
j=1
while not rs.eof%>
  <tr>
    <td><div align="center"><%=j%></div></td>
    <td><div align="center"><span class="red">
      <input name="pfactory<%=j%>" type="hidden" id="pfactory<%=j%>" value="<%=rs("FACTORY_ID")%>">
      <input name="p_pid<%=j%>" type="hidden" id="p_pid<%=j%>" value="<%=rs("NID")%>">
      <input name="pid<%=j%>" type="checkbox" id="nid" value="1"
	<%
	for l=0 to ubound(apid)
	if rs("NID")=apid(l) then%>checked
	<%
	end if
	next%>>
    </span></div></td>
    <td><div align="center"><%=rs("FACTORY_NAME")%></div></td>
    <td height="20"><div align="left"><strong><%=rs("PART_NUMBER")%></strong>&nbsp;(<%=rs("PART_RULE")%>)&nbsp;</div></td>
  </tr>
  <%rs.movenext
j=j+1
wend
else%>
  <tr>
    <td height="20" colspan="4">&nbsp;</td>
  </tr>
  <%end if
rs.close%>
  <tr>
    <td height="20" colspan="4"><div align="center">
      <input name="icount" type="hidden" id="icount" value="<%=i-1%>">
      <input name="jcount" type="hidden" id="jcount" value="<%=j-1%>">
      <input name="code" type="hidden" id="code" value="<%=code%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <select name="factory" id="factory">
        <option value=""></option>
        <option value="all" <%if station="all" then%>selected<%end if%>>All</option>
        <option value="null" <%if station="null" then%>selected<%end if%>>None</option>
        <%FactoryRight ""%>
        <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
      </select>
      <select name="type" id="type">
	  	<option value=""></option>
        <option value="all">All</option>
        <option value="station">Station</option>
        <option value="part">Part</option>
      </select>
      <input name="CheckAll" type="button" id="CheckAll" value="Check All" onClick="if  (document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='' ||document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='all' ||document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='null'){checkalla()}else{checkallc(document.checkform.factory.options[document.checkform.factory.selectedIndex].value,document.checkform.type.options[document.checkform.type.selectedIndex].value)}">
&nbsp;
<input name="UnCheckAll" type="button" id="UnCheckAll" value="Uncheck All" onClick="if  (document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='' ||document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='all' ||document.checkform.factory.options[document.checkform.factory.selectedIndex].value=='null'){uncheckalla()}else{uncheckallc(document.checkform.factory.options[document.checkform.factory.selectedIndex].value,document.checkform.type.options[document.checkform.type.selectedIndex].value)}">
&nbsp;
<input name="Update" type="submit" id="Update" value="Update">
      &nbsp;
      <input name="Reset" type="reset" id="Reset" value="Reset">
      &nbsp;</div></td>
  </tr>
</table>
</form>
<!--#include virtual="/Functions/CheckControl.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->