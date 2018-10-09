<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Line/LineCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetPart.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
linename=request.QueryString("linename")
thisstatus=request.QueryString("status")
factory=request.QueryString("factory")
if linename<>"" then
where=where&" and lower(L.LINE_NAME) like '%"&lcase(linename)&"%'"
end if
if thisstatus="" or thisstatus="all" then
where=where&""
else
where=where&" and L.STATUS="&thisstatus
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and L.FACTORY_ID is null"
else
where=where&" and L.FACTORY_ID='"&factory&"'"
end if

pagename="/Admin/Line/Line.asp"
pagepara="&linename="&linename&"&factory="&factory
FactoryRight "L."
SQL="select L.*,F.FACTORY_NAME,S.SECTION_NAME,U.USER_CHINESE_NAME as LEADER_NAME,US.USER_NAME as SUPERVISOR_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID left join SECTION S on L.SECTION_ID=S.NID left join USERS U on L.LEADER=U.USER_CODE left join USERS US on L.SUPERVISOR=US.USER_CODE where L.LINE_NAME is not null "&where&factorywhereoutsideand&" order by L.LINE_NAME"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Admin/Line/Lan_Line.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Admin/Line/Line.asp" method="get" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr align="center">
    <td width="60" height="20"><span id="inner_SearchLine"></span></td>
    <td width="100"><input name="linename" type="text" id="linename" value="<%=linename%>"></td>
    <td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="17" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="17" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Line/AddLine.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_Add"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="17"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>  
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_LineName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_FactoryCode"></span></div></td>
    <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CodeDate690"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CodeDate"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CodeLineName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CodeName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_690CodeName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_VersionNumber"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Factory"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Section"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_GroupLeader"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Supervisor"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_MachineLabels"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_Product"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditLine.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
<!--    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('DeleteLine.asp?id=<%'=rs("NID")%>&line_name=<%'=rs("LINE_NAME")%>&path=<%'=path%>&query=<%'=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>-->
	<td><div align="center"><span class="red">
	  <%if rs("STATUS")="1" then%>
	  <span style="cursor:hand" onClick="javascript:location.href='Disableline.asp?id=<%=rs("NID")%>&linename=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Line"><img src="/Images/Enabled.gif"></span>
	  <%else%>
	  <span style="cursor:hand" onClick="javascript:location.href='EnableLine.asp?id=<%=rs("NID")%>&linename=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Line"><img src="/Images/Disabled.gif"></span>
	  <%end if%>
    </span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("LINE_NAME") %></div></td>
    <td height="20"><div align="center"><%= rs("FACTORY_CODE") %>&nbsp;</div></td>
     <td height="20"><div align="center"><%= rs("CODE_Date690") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("CODE_Date") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("CODE_LINENAME") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("CODE_NAME") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("CODE_NAME2") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("VERSION_NUMBER")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%></div></td>
    <td><div align="center"><%= rs("SECTION_NAME")%></div></td>
    <td><div align="center"><%= rs("LEADER_NAME")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("SUPERVISOR_NAME")%>&nbsp;</div></td>
    <td><div align="center">&nbsp;
  </td>
    <td height="20"><div align="left"><%= rs("PRODUCT")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="17"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->