<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
factory=request("factory")

FactoryRight "SG."
SQL="select 1,SG.*,F.FACTORY_NAME,S.SECTION_NAME from SERIES_GROUP SG inner join FACTORY F on SG.FACTORY_ID=F.NID left join SECTION S on SG.SECTION_ID=S.NID where 1=1 order by SG.SERIES_GROUP_NAME"
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>家族清单</title>
<link href="/CSS/Print.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="540"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="3"><div align="center">家族清单</div></td>
</tr>

<tr>
  <td height="20" colspan="3">&nbsp;</td>
</tr>
<tr class="today">
  <td height="20" class="t-t-DarkBlue"><div align="center">序列</div></td>
  <td height="20" class="t-t-DarkBlue"><div align="center">家族名称 </div></td>
  <td class="t-t-DarkBlue"><div align="center">包括型号</div></td>
  </tr>
<%
i=1
if not rs.eof then
while not rs.eof
%>
<tr>
  <td height="20"><div align="center">
    <% =i%>
  </div></td>
    <td height="20"><div align="center"><%= rs("SERIES_GROUP_NAME") %></div></td>
    <td><%if rs("INCLUDED_SYSTEM_ITEMS")<>"" then%>
      <%=rs("INCLUDED_SYSTEM_ITEMS")%>
      <%end if%>&nbsp;</td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="3"><div align="center">无数据&nbsp;</div></td>
  </tr>
<%end if
rs.close%>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->