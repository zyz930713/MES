<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/OracleLine/OracleLineCheck.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
organization_id=request("organization_id")
if organization_id="" then
organization_id=24
end if
pagename="/Admin/OracleLine/OracleLine.asp"
pagepara="&organization_id="&organization_id
SQL="select L.LINE_ID,L.LINE_CODE,L.ORGANIZATION_ID,S.SubQuantity from tbl_Wip_Lines L left join tbl_Wip_Line_Sub S on L.Line_Id=S.Line_Id where L.ORGANIZATION_ID="&organization_id&" order by L.LINE_CODE"
rsPR.open SQL,connPr,1,3
%>
<!--#include virtual="/Components/PageSelect_PR.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy">Browse Line List in Oracle </td>
</tr>
<tr>
  <td height="20" colspan="5" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User:
          <% =session("User") %></td>
        <td width="50%"><div align="right">
          <%if admin=true then%>
          <a href="/Admin/Line/AddLine.asp?path=<%=path%>&query=<%=query%>" target="main" class="white">Add a New Line </a>
          <%end if%>
        </div></td>
      </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="5"><!--#include virtual="/Components/PageSplit_PR.asp" --></td>
</tr>
<tr>
  <td height="20" colspan="5" class="t-t-Borrow"><input name="or" type="radio" value="24" onClick="javascript:location.href='<%=pagename%>?organization_id='+this.value">
    C10
    <input name="or" type="radio" value="25" onClick="javascript:location.href='<%=pagename%>?organization_id='+this.value">
    C11</td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">No</div></td>
  <%if admin=true then%>
  <td height="20" class="t-t-Borrow"><div align="center">Action</div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center">Line Code </div></td>
  <td class="t-t-Borrow"><div align="center">Orgnization</div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Sublot Quantity </div></td>
</tr>
<%
i=1
if not rsPR.eof then
rsPR.absolutepage=cint(session("strpagenum"))
while not rsPR.eof and i<=rsPR.pagesize 
%>
<tr>
  <td height="20"><div align="center">
      <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
  <td height="20"><div align="center"><span class="red"style="cursPRor:hand" onClick="javascript:window.open('EditOracleLine.asp?line_id=<%=rsPR("LINE_ID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div></td>
  <%end if%>
  <td height="20"><div align="center"><%= rsPR("LINE_CODE") %></div></td>
  <td><div align="center"><%= rsPR("ORGANIZATION_ID")%></div></td>
  <td height="20"><div align="center"><%= rsPR("SubQuantity")%>&nbsp;</div></td>
</tr>
<%
i=i+1
rsPR.movenext
wend
else
%>
<tr>
  <td height="20" colspan="5">No Records&nbsp;</td>
</tr>
<%end if
rsPR.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>