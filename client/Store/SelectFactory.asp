<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Select Factory 选择工厂</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css" />
</head>

<body bgcolor="#339966">
<form id="form1" name="form1" method="post" action="/Store/Store1.asp">
  <table width="400" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" class="t-t-DarkBlue" align="center" >Select Factory 选择工厂</td>
    </tr>
	<%SQL="select NID||'-'||FACTORY_NAME AS FACTORY,FACTORY_NAME from FACTORY order by FACTORY_NAME"
	rs.open SQL,conn,1,3
	if not rs.eof then
	while not rs.eof%>
    <tr>
      <td height="20"><input name="factory" type="radio" value="<%=rs("FACTORY")%>" onclick="form1.Next.disabled=false;" />
		<%=rs("FACTORY_NAME")%>
	  </td>
    </tr>
	<%
	rs.movenext
	wend
	end if
	rs.close%>
    <tr>
      <td height="20"><div align="center">
        <input type="submit" name="Next" value="Next 下一步" disabled />
      </div></td>
    </tr>
  </table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->