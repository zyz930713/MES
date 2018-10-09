<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")

line=request("line")
computer=request("computer")

where="where 1=1 "

if line<>"" then
where=where&" and line_name like '%"&line&"%'"
end if
if computer<>"" then
where=where&" and computer_name like '%"&computer&"%'"
end if

pagepara="&line="&line&"&computer="&computer
pagename="/System/ComputerSet/ComputerSet.asp" 
SQL="select computer_name,printer_name,line_name,PRODUCT from computer_printer_mapping "&where
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form name="form1" method="post" >
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr><td><table ><tr>
    <td height="20"><span id="inner_Computer"></span></td>
    <td height="20"><input name="computer" type="computer" id="code" value="<%=computer%>"></td>
    <td><span id="inner_SearchLineName"></span></td>
    <td><input name="line" type="text" id="line" value="<%=line%>"></td>    
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr></table></td></tr>
</table>
</form>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
          <% =session("User") %>        </td>
        <td width="50%"><div align="right"><a href="/System/ComputerSet/AddComputerSet.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a></div></td>
      </tr>
    </table></td>
  </tr>
  <tr class="silver-t-t">
    <td height="20" colspan="16"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
  <tr class="t-t-Borrow">
    <td height="20"><div align="center"><span id="inner_NO"></span></div></td>
    <td height="20" colspan="2"><div align="center"><span id="inner_Action"></span></div></td>
    <td height="20"><div align="center"><span id="td_Computer"></span></div></td>
	<td height="20"><div align="center"><span id="td_Printer"></span></div></td>  
    <td height="20"><div align="center"><span id="td_LineName"></span></div></td> 
    <td height="20"><div align="center"><span id="inner_SubSeries"></span></div></td>
       
  </tr>
  <%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize 
%>
  <tr>
    <td height="20">
      <div align="center">
        <% =(session("strpagenum")-1)*recordsize+i%>
    </div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:window.open('/System/ComputerSet/EditComputerSet.asp?computer=<%=rs("computer_name")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif"></span></div></td>
    <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this reord?\n您确定删除此记录吗？')){window.open('/System/ComputerSet/DeleteComputerSet.asp?computer=<%=rs("computer_name")%>&path=<%=path%>&query=<%=query%>','main')}"><img src="/Images/IconDelete.gif"></span></div></td>
    <td height="20"><div align="center"><%=rs("computer_name")%></div></td>    
    <td height="20"><div align="center"><%=rs("printer_name")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("line_name")%>&nbsp;</div></td>
    <td><div align="center"><%=rs("PRODUCT")%>&nbsp;</div></td>
  </tr>
  <%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="16"><div align="center"><span id="inner_Records"></span>&nbsp;</div></td>
  </tr>
  <%end if%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->