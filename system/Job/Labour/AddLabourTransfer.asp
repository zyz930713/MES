<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
SQL="select LINE_NAME from LINE where LEADER='3125'"
'SQL="select LINE_NAME from LINE where LEADER='"&session("code")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
hostline=rs("LINE_NAME")
end if
rs.close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Action/FormCheck.js" type="text/javascript"></script>
</head>

<body>
<form action="/Admin/Action/AddAction1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Add a New Labour Transfer </td>
</tr>
<tr>
  <td height="20">Host Line </td>
  <td height="20" class="red"><%= hostline %>&nbsp;</td>
</tr>
<tr>
  <td height="20">Year</td>
  <td height="20" class="red"><input name="year" type="text" id="year" value="<%year(date())%>"></td>
</tr>
<tr>
  <td width="113" height="20"><div align="left">Week</div></td>
    <td width="641" height="20" class="red">
      <div align="left"><%=weekindex%>
        <input name="week" type="text" id="week" value="<%weekday(date())%>">
      </div></td>
    </tr>
<tr>
  <td height="20">Transfer Hour </td>
  <td height="20"><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#0066CC" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="2" class="t-b-blueReal"><div align="center">Transfer out </div></td>
      <td height="20" colspan="2" class="t-b-Green"><div align="center">Transfer in </div></td>
      </tr>
    <tr>
      <td height="20" class="t-b-blueReal"><div align="center">Line</div></td>
      <td height="20" class="t-b-blueReal"><div align="center">Hour</div></td>
      <td height="20" class="t-b-Green"><div align="center">Line</div></td>
      <td height="20" class="t-b-Green"><div align="center">Hour</div></td>
    </tr>
	<%SQL="select * from LINE where FACTORY_ID='"&session("factory")&"' and LINE_NAME<>'"&hostline&"'"
	rs.open SQL,conn,1,3
	i=1
	while not rs.eof
	%>
    <tr>
      <td height="20"><div align="center"><%= rs("LINE_NAME") %>
        <input name="out_id<%=i%>" type="hidden" id="out_id<%=i%>">
      </div></td>
      <td height="20"><div align="center">
        <input name="out_hour<%=i%>" type="text" id="out_hour<%=i%>">
      Hours</div></td>
      <td height="20"><div align="center"><%= rs("LINE_NAME") %>
        <input name="in_id<%=i%>" type="hidden" id="in_id<%=i%>">
      </div></td>
      <td height="20"><div align="center">
        <input name="in_hour<%=i%>" type="text" id="in_hour<%=i%>">
       Hours </div></td>
    </tr>
    <%
	rs.movenext
	i=i+1
	wend
	rs.close%>
  </table></td>
</tr>
  
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="lcount" type="hidden" id="lcount" value="<%=i%>">
	  <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Save">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/Functions/TableControl.asp" -->
