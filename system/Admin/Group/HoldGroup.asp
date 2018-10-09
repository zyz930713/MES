<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Group/GroupCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
SQL="select NID,GROUP_NAME,GROUP_TYPE,GROUP_MEMBERS,STATUS,GROUP_CHINESE_NAME,FACTORY_ID from SYSTEM_GROUP where GROUP_TYPE='Hold' order by NID"
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>

<body onLoad="language_page();language(<%=session("language")%>);">

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><span id="inner_BrowseData"></span>
  </td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:<% =session("User") %></td>      
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="11"><!--#include virtual="/Components/PageSplit.asp" --></td>
</tr>
<tr>
  <td height="20" class="t-t-Borrow" align="center"><span id="inner_NO"></span></td>  
  <td height="20" colspan="2" class="t-t-Borrow" align="center"><span id="inner_Action"></span></td>    
  <td height="20" class="t-t-Borrow" align="center"><span id="td_Name"></span></td>
  <td class="t-t-Borrow" align="center"><span id="td_CHName"></span></td> 
  <td class="t-t-Borrow" align="center"><span id="td_GroupType"></span></td>
  <td class="t-t-Borrow" align="center"><span id="td_Member"></span></td>  
</tr>
<%
if rs.eof then
%>
  <tr>
    <td height="20" colspan="11"><div align="center"><span id="inner_Records"></span>&nbsp;</div></td>
  </tr>
<%
else
i=1
while not rs.eof
%>
  <tr <%if rs("STATUS")="0" then%>class="t-b-blue"<%end if%>>
    <td height="20" align="center"><%=i%></td>	
    <td height="20" align="center">&nbsp;
        <%if rs("STATUS")="1" then%>
            <span style="cursor:hand" onClick="javascript:window.open('EditGroup.asp?id=<%=rs("NID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"  align="absmiddle"></span>
        <%end if%>
    </td><td height="20" align="center">
        <%if rs("STATUS")="1" then%>
            <span style="cursor:hand" onClick="javascript:location.href='DisableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to disable Group"><img src="/Images/Enabled.gif"></span>
        <%else%>
            <span style="cursor:hand" onClick="javascript:location.href='EnableGroup.asp?id=<%=rs("NID")%>&groupname=<%=rs("GROUP_NAME")%>&path=<%=path%>&query=<%=query%>'" title="Click to enable Group"><img src="/Images/Disabled.gif"></span>
        <%end if%>
    </td>        
    <td height="20" align="center"><%=rs("GROUP_NAME")%></td>
    <td align="center"><%=rs("GROUP_CHINESE_NAME")%></td>    
    <td align="center"><%=rs("GROUP_TYPE")%></td>
    <td align="left" >
    <%if rs("GROUP_MEMBERS") <> "" then
        response.Write(formatlongstring(replace(rs("GROUP_MEMBERS"),",",",&nbsp;"),"<br>",140))
      else
        response.Write("&nbsp;")
      end if    
    %></td>
  </tr>
<%
i=i+1
rs.movenext
wend
end if
%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
</body>
</html>