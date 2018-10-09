<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
	admin=true
	subSeriesName=request("txtSubSeriesName")
	SQL="SELECT * FROM GENERAL_SETTING WHERE IsExpired='0' AND MODELNAME is null "
	if subSeriesName <>"" then
		SQL=SQL+" AND SUBSERIESNAME LIKE '"&subSeriesName&"%' "
	end if
	rs.open SQL,conn,1,3
	
	if request.QueryString("action")="1" then
		id=request.QueryString("id")
		'SQL="Update GENERAL_SETTING set IsExpired='1',LASTUPDATECODE='"&session("code")&"',LSTUPDATEDATETIME='"&now()&"' "
		SQL="delete GENERAL_SETTING "
		SQL=SQL+" where subseriesname='"&id&"'"
		set rsD=server.createobject("adodb.recordset")
		rsD.open SQL,conn,1,3
		response.write "<script>window.alert('Delete one setting successfully!');location.href='GeneralSetting.asp';</script>"
	end if
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>
<body onLoad="language_page();language(<%=session("language")%>);">
<form  method="post" name="form1" action="GeneralSetting.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="5" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_SubSeries"></span></td>
    <td><input name="txtSubSeriesName" type="text" id="txtSubSeriesName" value="<%=subSeriesName%>"></td>    
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()" ></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%><a href="/Admin/GeneralSetting/AddGeneralSetting.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a><%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="12"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td class="t-t-Borrow"><div align="center"><span id="td_SubSeries"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_CurrentAQL"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_MaxAQL"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_MinAQL"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
</tr>
<%
	i=1
	if rs.recordcount>0 then
	while not rs.eof and i<=rs.pagesize 		
%>
<tr>
  <td height="20" ><div align="center"><% =(cint(session("strpagenum"))-1)*recordsize+i%></div></td>
  <%if admin=true then%>
  <td height="20" ><div align="center">
	<span class="red" style="cursor:hand" onClick="javascript:window.open('AddGeneralSetting.asp?action=2&id=<%=rs("GNID")%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('GeneralSetting.asp?action=1&id=<%=rs("subseriesname")%>','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>	
  <%end if%>
  <td><div align="center"><%=rs("SUBSERIESNAME")%> &nbsp;</div></td>
  <td><div align="center"><%=rs("CurrentAQL")%>&nbsp;</div></td>
  <td><div align="center"><%=rs("MAXAQL")%>&nbsp;</div></td>
  <td><div align="center"><%=rs("MINAQL")%>&nbsp;</div></td>
   <td height="20" ><div align="center">
	<span class="red" style="cursor:hand" onClick="javascript:window.showModalDialog('EditSubSeries.asp?id=<%=rs("SUBSERIESNAME")%>','_blank','dialogHeight:600px;dialogWidth:1024px')"><img src="/Images/IconEdit.gif" alt="Click to edit sub series"></span></div></td>
</tr>
<%i=i+1
	rs.movenext
	wend
	else
%>

  <tr>
    <td height="20" colspan="12"><div align="center"><span id="inner_Records"></span>&nbsp;</div></td>
  </tr> 
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->