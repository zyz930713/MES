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
PRIORITY_LEVEL=request.QueryString("PRIORITY_LEVEL")
if action="edit" then
	if(PRIORITY_LEVEL<>"" )then
		set rsPRIORITYLEVEL=server.createobject("adodb.recordset")
		SQL="SELECT * FROM JOB_PRIORITY_SETTING WHERE PRIORITY_LEVEL='"+PRIORITY_LEVEL+"'"
		rsPRIORITYLEVEL.open SQL,conn,1,3
		if rsPRIORITYLEVEL.recordcount>0 then
			PRIORITY_LEVEL=rsPRIORITYLEVEL("PRIORITY_LEVEL")
			PRIORITY_DEC=rsPRIORITYLEVEL("PRIORITY_DEC")
		end if 
	end if 
end if 

if(action="2") then
	PRIORITY_LEVEL=request("txtPRIORITY_LEVEL")
	PRIORITY_DEC=request("txtPRIORITY_DEC")
	action=request("txtaction")
	set rsPRIORITYLEVEL=server.createobject("adodb.recordset")
	SQL="SELECT * FROM JOB_PRIORITY_SETTING WHERE PRIORITY_LEVEL='"+PRIORITY_LEVEL+"'"
	rsPRIORITYLEVEL.open SQL,conn,1,3
	if(rsPRIORITYLEVEL.recordcount>0 and action<>"edit") then
		response.write("<script>alert('The Priority Level"+PRIORITY_LEVEL+" is Defined already!');</script>")
	end if
	
	if(rsPRIORITYLEVEL.recordcount>0 and action="edit") then
		rsPRIORITYLEVEL("PRIORITY_LEVEL")=PRIORITY_LEVEL
		rsPRIORITYLEVEL("PRIORITY_DEC")=PRIORITY_DEC
		rsPRIORITYLEVEL.update
		response.write("<script>alert('Save Successfully!');location.href='JobPriorityList.asp'</script>")
	end if 
	
	if(rsPRIORITYLEVEL.recordcount=0 ) then
		rsPRIORITYLEVEL.addnew
		rsPRIORITYLEVEL("PRIORITY_LEVEL")=PRIORITY_LEVEL
		rsPRIORITYLEVEL("PRIORITY_DEC")=PRIORITY_DEC
		rsPRIORITYLEVEL.update
		response.write("<script>alert('Save Successfully!');location.href='JobPriorityList.asp'</script>")
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
		if(document.getElementById("txtPRIORITY_LEVEL").value=="")
		{
			window.alert("Please input Priority Level!");
			return;
		}
		if(isNaN(document.getElementById("txtPRIORITY_LEVEL").value))
		{
			window.alert("Priority Level should be numeric!");
			return;
		}
		if(document.getElementById("txtPRIORITY_DEC").value=="")
		{
			window.alert("Please input Priority Description!");
			return;
		}
		form1.action="JobPriority.asp?action=2";
		form1.submit();
	}
</script>

</head>

<body>
<form method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
	<%if action<>"edit" then%>
      <td height="20" colspan="10" class="t-c-greenCopy">Add a New Priority Level</td>
	  <%else%>
	  <td height="20" colspan="10" class="t-c-greenCopy">Edit a New Priority Level</td>
	 <%end if %>
    </tr>
  
    <tr>
      <td height="20" width="15%">Priority Level <span class="red">*</span> </td>
      <td height="20" class="red" colspan="9">
        <div align="left">
          <input name="txtPRIORITY_LEVEL" type="text" id="txtPRIORITY_LEVEL" size="30" value="<%=PRIORITY_LEVEL%>" <%if action="edit" then response.write "readonly" end if %>>
      </div></td>
    </tr>
    <tr>
      <td height="20">Priority Description  <span class="red">*</span></td>
      <td height="20" colspan="9"><input name="txtPRIORITY_DEC" type="text" id="txtPRIORITY_DEC" size="30" value="<%=PRIORITY_DEC%>"></td>
    </tr>
    <tr>
      <td height="20" colspan="10"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
		    <input name="txtaction" type="hidden" id="txtaction" size="30" value="<%=action%>" >
        <input type="reset" name="Submit7" value="Return" width="75px" onclick="location.href='JobPriorityList.asp'">
</div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->