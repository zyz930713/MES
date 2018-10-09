<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Model/ModelCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<%
modelname=request("modelname")
modelstatus=request("status")
factory=request("factory")
where=""
if modelname<>"" then
where=where&" and P.ITEM_NAME like '%"&modelname&"%'"
end if
if modelstatus<>"" and modelstatus<>"all" then
where=where&" and P.ITEM_STATUS='"&modelstatus&"'"
end if
if factory="" or factory="all" then
where=where&""
elseif factory="null" then
where=where&" and P.FACTORY_ID is null"
else
where=where&" and P.FACTORY_ID='"&factory&"'"
end if

pagepara="&modelname="&modelname&"&status="&modelstatus&"&factory="&factory
SQL="select P.*,F.FACTORY_NAME from PRODUCT_MODEL P inner join FACTORY F on P.FACTORY_ID=F.NID where P.ITEM_NAME is not null"&where&" order by P.ITEM_NAME"
'response.Write(SQL)
'response.End()
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Admin/Model/Model.asp" method="get" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td  height="20"><span id="inner_SearchPartNumber"></span></td>
    <td ><input name="modelname" type="text" id="modelname" value="<%=modelname%>"></td>
    <td ><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursPRor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="17" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="17" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="17"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Description"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_BoxSize"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SmallPACK"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerPN"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerDefine"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerLabel"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerDesc"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerPegapn"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerConfig"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_YNLittleLable"></span></div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="factory" id="factory" onChange="location.href='<%=pagename%>?modelname=<%=modelname%>&factory='+this.options[this.selectedIndex].value+'&modelstatus=<%=modelstatus%>'">
      <option value="all">Factory</option>
      <%FactoryRight ""%>
      <%= getFactory("OPTION",factory,factorywhereinside,"","") %>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center">
    <select name="status" id="status" onChange="location.href='<%=pagename%>?modelname=<%=modelname%>&factory=<%=factory%>&status='+this.options[this.selectedIndex].value">
      <option>Status</option>
      <option value="all">All</option>
      <option value="ACTIVE" <%if modelstatus="ACTIVE" then%>selected<%end if%>>ACTIVE</option>
      <option value="INACTIVE" <%if modelstatus="INACTIVE" then%>selected<%end if%>>INACTIVE</option>
      <option value="OBSOLETE" <%if modelstatus="OBSOLETE" then%>selected<%end if%>>OBSOLETE</option>
      <option value="LTB" <%if modelstatus="LTB" then%>selected<%end if%>>LTB</option>
      <option value="PROTO TYPE" <%if modelstatus="PROTO TYPE" then%>selected<%end if%>>PROTO TYPE</option>
	  <option value="GREEN" <%if modelstatus="GREEN" then%>selected<%end if%>>GREEN</option>
	  <option value="ENTERING">ENTERING</option>
    </select>
  </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime"></span> </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LeadTime2"></span></div></td>
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
  <td><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditModel.asp?id=<%=rs("ITEM_ID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <%end if%>
    <td height="20"><div align="center"><%= rs("ITEM_NAME") %></div></td>
	<td><div align="center"><%= rs("DESCRIPTION")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("BOX_SIZE")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("SMALL_PACK")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("CUSTOMER_PN")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("CUSTOMER_DEFINE")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("CUSTOMER_LABEL")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("CUSTOMER_DESC")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("CUSTOMER_PEGAPN")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("CUSTOMER_CONFIG")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("YESNOLITTLELABLE")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("FACTORY_NAME")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("ITEM_STATUS")%></div></td>
    <td><div align="center"><%= rs("LEAD_TIME")%></div></td>
    <td><div align="center"><%= rs("LEAD_TIME2")%></div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="17" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->