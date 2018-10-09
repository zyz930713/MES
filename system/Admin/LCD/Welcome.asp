<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
thisday=request.QueryString("thisday")
SQL="select * from DAILY_WELCOME where to_char(WELCOME_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
if not rs.eof then
file_name=rs("FILE_NAME")
end if
rs.close
%>
<html>
<head>
<title>Submit job you selected</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function formcheck()
{
	with(document.form1)
	{
		if (attachment.value=="")
		{
		alert("File Name can not be blank!");
		return false;
		}
	}
}
</script>
</head>

<body>
<form action="/Admin/LCD/Welcome1.asp" method="post" enctype="multipart/form-data" name="form1" onSubmit="return formcheck()">
  <table width="600" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
    
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy">Upload Welcome Powerpoint File </td>
    </tr>
    <tr>
      <td width="124" height="20">Date</td>
      <td width="470" height="20"><%=thisday%></td>
    </tr>
    <tr>
      <td height="20">Existed File Name </td>
      <td height="20"><%=file_name%>&nbsp;</td>
    </tr>
    <tr>
      <td height="20">File</td>
      <td height="20"><input name="attachment" type="file" id="attachment" size="50">
        Format is PPS; Size &lt;=5MB</td>
    </tr>
    <tr>
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="20" colspan="2"><div align="center"> 
          <input name="thisday" type="hidden" id="thisday" value="<%=thisday%>">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input name="Upload" type="submit" id="Upload" value="Upload">
          &nbsp; 
          <input type="reset" name="Submit2" value="Reset">
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->