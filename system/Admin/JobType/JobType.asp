<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<%

action=request.QueryString("Action")
JOB_TYPE=request.QueryString("JOB_TYPE")
if action="edit" then
	if(JOB_TYPE<>"" )then
		set rsJOB_TYPE=server.createobject("adodb.recordset")
		SQL="SELECT * FROM JOB_TYPE_SETTING WHERE JOB_TYPE='"+JOB_TYPE+"'"
		rsJOB_TYPE.open SQL,conn,1,3
		if rsJOB_TYPE.recordcount>0 then
			JOB_TYPE=rsJOB_TYPE("JOB_TYPE")
			JOB_TYPE_DESC=rsJOB_TYPE("JOB_TYPE_DESC")
		end if 
	end if 
end if 

if(action="2") then
	JOB_TYPE=request("txtJOB_TYPE")
	JOB_TYPE_DESC=request("txtJOB_TYPE_DESC")
	action=request("txtaction")
	set rsJOB_TYPE=server.createobject("adodb.recordset")
	SQL="SELECT * FROM JOB_TYPE_SETTING WHERE JOB_TYPE='"+JOB_TYPE+"'"
	rsJOB_TYPE.open SQL,conn,1,3
	if(rsJOB_TYPE.recordcount>0 and action<>"edit") then
		response.write("<script>alert('The Job Type"+JOB_TYPE+" is Defined already!');</script>")
	end if
	
	if(rsJOB_TYPE.recordcount>0 and action="edit") then
		rsJOB_TYPE("JOB_TYPE")=JOB_TYPE
		rsJOB_TYPE("JOB_TYPE_DESC")=JOB_TYPE_DESC
		rsJOB_TYPE.update
		response.write("<script>alert('Save Successfully!');location.href='JobTypeList.asp'</script>")
	end if 
	
	if(rsJOB_TYPE.recordcount=0 ) then
		rsJOB_TYPE.addnew
		rsJOB_TYPE("JOB_TYPE")=JOB_TYPE
		rsJOB_TYPE("JOB_TYPE_DESC")=JOB_TYPE_DESC
		rsJOB_TYPE.update
		response.write("<script>alert('Save Successfully!');location.href='JobTypeList.asp'</script>")
	end if 
end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script>
	function SaveData()
	{
		if(document.getElementById("txtJOB_TYPE").value=="")
		{
			window.alert("Please input Job Type!");
			return;
		}
		 
		if(document.getElementById("txtJOB_TYPE_DESC").value=="")
		{
			window.alert("Please input Job Type Description!");
			return;
		}
		form1.action="JobType.asp?action=2";
		form1.submit();
	}
</script>

</head>

<body>
<form method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
	<%if action<>"edit" then%>
      <td height="20" colspan="10" class="t-c-greenCopy">Add a New Job Type</td>
	  <%else%>
	  <td height="20" colspan="10" class="t-c-greenCopy">Edit a New Job Type</td>
	 <%end if %>
    </tr>
  
    <tr>
      <td height="20" width="15%">Job Type <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
          <input name="txtJOB_TYPE" type="text" id="txtJOB_TYPE" size="30" value="<%=JOB_TYPE%>" <%if action="edit" then response.write "readonly" end if %>>
      </div></td>
    </tr>
    <tr>
      <td height="20">Job Type Description  <span class="red">*</span></td>
      <td height="20" colspan="9"><input name="txtJOB_TYPE_DESC" type="text" id="txtJOB_TYPE_DESC" size="30" value="<%=JOB_TYPE_DESC%>"></td>
    </tr>
    <tr>
      <td height="20" colspan="10"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
		    <input name="txtaction" type="hidden" id="txtaction" size="30" value="<%=action%>" >
        <input type="reset" name="Submit7" value="Return" width="75px" onclick="location.href='JobTypeList.asp'">
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->