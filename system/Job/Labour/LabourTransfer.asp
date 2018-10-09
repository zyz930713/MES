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
linename=request("linename")
factory=request.QueryString("factory")
if linename<>"" then
where=where&" and lower(L.LINE_NAME) like '%"&lcase(linename)&"%'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and L.FACTORY_ID is null"
else
where=where&" and L.FACTORY_ID='"&factory&"'"
end if

pagename="/Job/Labour/LabourTransfer.asp"
pagepara="&linename="&linename&"&factory="&factory
FactoryRight "L."
if session("code")<>"1194" and session("code")<>"1488" then
SQL="select 1,LL.*,U.USER_CHINESE_NAME from LINE_LABOUR_TRANSFER LL inner join LINE L on LL.LINE_NAME=L.LINE_NAME left join USERS U on L.LEADER=U.USER_CODE where L.LEADER='"&session("code")&"' "&where&factorywhereoutsideand&" order by L.LINE_NAME"
else
SQL="select 1,LL.*,U.USER_CHINESE_NAME from LINE_LABOUR_TRANSFER LL inner join LINE L on LL.LINE_NAME=L.LINE_NAME left join USERS U on L.LEADER=U.USER_CODE where 1=1 "&where&factorywhereoutsideand&" order by L.LINE_NAME"
end if
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Job/Labour/Lan_LabourTransfer.asp" -->
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language();language_page()">
<form action="/Job/Shift/Shift.asp" method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="11%" height="20"><span id="inner_SearchLine"></span></td>
    <td width="18%"><input name="linename" type="text" id="linename" value="<%=linename%>"></td>
    <td width="71%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Browse"></span></td>
</tr>
<tr>
  <td height="20" colspan="9" class="t-c-greenCopy">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%"><span id="inner_User"></span>: 
          <% =session("User") %></td>
        <td width="50%"><div align="right"><a href="/Job/Labour/AddLabourTransfer.asp">Add a new Transfer Eemtry</a></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="9"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" nowrap class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <td height="20" colspan="2" nowrap class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_YearIndex"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_Weekindex"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_TransferType"></span></div></td>
  <td width="241" class="t-t-Borrow"><div align="center"><span id="inner_TransferHour"></span></div></td>
  <td width="241" height="20" class="t-t-Borrow"><div align="center"><span id="inner_InputCode"></span></div></td>
  <td width="235" class="t-t-Borrow"><div align="center"><span id="inner_UpdateTime"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(session("strpagenum")-1)*recordsize+i%>
  </div></td>
    <td height="20"><div align="center" class="red"><span style="cursor:hand" onClick="javascript:window.open('EditLabourTransfer.asp?id=<%=rs("NID")%>&linename=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td width="241" height="20" class="red"><div align="center"><span style="cursor:hand" onClick="javascript:if(confirm('Are you sure to delete this Action?')){window.open('DeleteLabourTransfer.asp?id=<%=rs("NID")%>&linename=<%=rs("LINE_NAME")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete"></span></div></td>
    <td><div align="center"><%= rs("YEAR_INDEX") %></div></td>
    <td><div align="center"><%= rs("WEEK_INDEX") %></div></td>
    <td><div align="center"><%= rs("TRANSFER_TYPE") %></div></td>
    <td><div align="center"><%= rs("TRANSFER_HOUR") %></div></td>
    <td height="20"><div align="center"><%= rs("INPUT_CODE") %></span></div></td>
    <td><div align="center"><%= rs("UPDATE_TIME")%></div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="9">No Records&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->