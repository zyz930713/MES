<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Packing/PACKINGCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->


<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
TEST_NAME=request("txt_TEST_NAME")
pagename="/Admin/PACKING/PackingTest.asp"
where=""
if TEST_NAME<>"" then
where=where&" and TEST_NAME like '%"&TEST_NAME&"%'"
end if
SQL="select * from part_test_name  where PART_NUMBER is not null"&where&" order by Test_type,PART_NUMBER,TEST_NAME"

rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form method="post" name="form1" target="_self">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
    <td height="20" colspan="7" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
<tr align="center">
    <td height="20" width="80"><span id="inner_TestName"></span></td>
    <td width="100"><input name="txt_TEST_NAME" type="text" id="txt_TEST_NAME" value="<%=TEST_NAME%>"></td> 
	<td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
</tr>	 
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">  
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="7" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/PACKING/AddPacking.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="7"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow" ><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_TestName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_AlarmMsg"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_testType"></span></div></td>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20" ><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
    <td height="20"><div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditPackingTest.asp?PART_ID=<%=rs("PART_ID")%>&path=<%=path%>&query=<%=query%>&Action=Edit','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('EditPackingTest1.asp?PART_ID=<%=rs("PART_ID")%>&TEST_NAME=<%=rs("TEST_NAME")%>&path=<%=path%>&query=<%=query%>&Action=Del','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
	<%end if%>
    <td height="20"><div align="center"><%= rs("PART_NUMBER") %></div></td>
	<td height="20"><div align="center"><%= rs("TEST_NAME") %></div></td>
    <td height="20"><div align="left"><%= rs("ALARM_MSG")%>&nbsp;</div></td>
    <td height="20"><div align="left"><%= rs("TEST_TYPE")%>&nbsp;</div></td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="6" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if

rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->