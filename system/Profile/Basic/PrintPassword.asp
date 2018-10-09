<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
	usercode=session("code")
	
	if(request.querystring("action")="1") then
		oldpassword=request("OldPassword")
		NewPassword=request("NewPassword")
		sql="select * from LABEL_PRINT_LINE_LEADER where USER_CODE='"+usercode+"'"
		set rsUserCode=server.createobject("adodb.recordset")
		rsUserCode.open sql,conn,1,3
		if(rsUserCode.recordcount=0) then
			response.write "<script>window.alert('不是授权用户!')</script>"
		else
			if(rsUserCode("PASSWORD")<>oldpassword)then
				response.write "<script>window.alert('旧密码不对!')</script>"
			else
				sql="update LABEL_PRINT_LINE_LEADER set PASSWORD='"+NewPassword+"' where USER_CODE='"+usercode+"'"
				set rsChangePassword=server.createobject("adodb.recordset")
				rsChangePassword.open sql,conn,1,3
				response.write "<script>window.alert('成功修改密码!')</script>"
			end if 
		end if
	end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Action/FormCheck.js" type="text/javascript"></script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<script>
	function ChangePassword()
	{
		if(form1.OldPassword.value=="")
		{
			window.alert("请输入旧密码!");
			return;
		}
		
		if(form1.NewPassword.value!=form1.NewPasswordConfirm.value)
		{
			window.alert("两次输入的新密码不一样!");
			return;
		}
		
		form1.action="PrintPassword.asp?action=1";
		form1.submit();
		
	}
</script>
</head>

<body>
<form  method="post" name="form1" target="_self">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Change Label Print Password</td>
</tr>
<tr>
  <td width="113" height="20"><div align="left">User Code<span class="red">*</span></div></td>
    <td width="641" height="20" class="red">
      <div align="left">
        <input name="usercode" type="text" id="usercode" size="50" value="<%=usercode%>" readonly=true>
      </div></td>
    </tr>
<tr>
  <td width="113" height="20"><div align="left">Old Password<span class="red">*</span></div></td>
    <td width="641" height="20" class="red">
      <div align="left">
        <input name="OldPassword" type="password" id="OldPassword" size="50">
      </div></td>
    </tr>
  <tr>
    <td height="20">New Password<span class="red">*</span></td>
    <td height="20"><input name="NewPassword" type="password" id="NewPassword" size="50"></td>
  </tr>
  <tr>
    <td height="20">New Password Confirm<span class="red">*</span></td>
    <td height="20"><input name="NewPasswordConfirm" type="password" id="NewPasswordConfirm" size="50"></td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input type="button" name="Save" value="Save" onclick="ChangePassword()">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/Functions/TableControl.asp" -->
