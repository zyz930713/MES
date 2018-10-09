
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
	if (request.querystring("Action")="") then
		batchno=request.querystring("batchno")
		jobno=request.querystring("jobno")
	
		SQL="SELECT * FROM Change_Model_History WHERE BatchNo='"+batchno+"' AND IsCurrentVersion='1'"
		set rsModelName=server.createobject("adodb.recordset")
		rsModelName.open SQL,conn,1,3
		OldModel=""
		NewModel=""
		if(rsModelName.recordcount>0)then
			OldModel=rsModelName("PrevoiusModel")
			NewModel=rsModelName("CurrentModel")
		else
			SQL="SELECT * FROM JOB_MASTER WHERE JOB_NUMBER='"+jobno+"'"
			set rsModelName2=server.createobject("adodb.recordset")
			rsModelName2.open SQL,conn,1,3
			if rsModelName2.recordcount>0 then
				OldModel=rsModelName2("PART_NUMBER_TAG")
				NewModel=rsModelName2("PART_NUMBER_TAG")
			end if 
		end if
	else
		batchno=request("txtbatchno")
		oldmodel=request("txtCurrentModel")
		newmodel=request("txtChangeToModel")
		
		SQL="UPDATE Change_Model_History SET IsCurrentVersion='0' WHERE BatchNo='"+batchno+"'"
		set rsUpdate1=server.createobject("adodb.recordset")
		rsUpdate1.open SQL,conn,1,3

		SQL="INSERT INTO Change_Model_History (BatchNo,PrevoiusModel,CurrentModel,ChangePerson,ChangeDateTime,IsCurrentVersion)"
		SQL=SQL+" values('"+batchno+"','"+OldModel+"','"+newmodel+"','"+Session("code")+"','"+cstr(now())+"','1')"
		set rsUpdate2=server.createobject("adodb.recordset")
		rsUpdate2.open SQL,conn,1,3
		response.write("<script>window.alert('Change Model successfully!');location.href='ChangeModel.asp'</script>")
		
	end if		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
<Script>
	function SaveData()
	{
		if (document.getElementById("txtChangeToModel").value=="")
		{
			window.alert("Please input new model!");
			return;
		}
		form1.action="ChangeModel_Detail.asp?Action=1";
		form1.submit();
	}
</script>
</head>

<body>

<form  method="post" name="form1" target="_self">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Change Model</td>
</tr>
<tr>
  <td width="116" height="20">BatchNo</td>
  <td width="638" height="20" class="red">
    <div align="left">
	<%=batchno%>
      <input name="txtbatchno" type="hidden" id="txtbatchno" value="<%=batchno%>" size="50">
  </div></td>
</tr>
<tr>
  <td height="20">Current Model</td>
  <td height="20"><%=NewModel%> <input name="txtCurrentModel" type="hidden" id="txtCurrentModel" value="<%=NewModel%>" size="50"></td>
</tr>
<tr>
  <td height="20">Change To Model</td>
  <td height="20"><input name="txtChangeToModel" type="text" id="txtChangeToModel" value="<%=ToModel%>" size="50"></td>
</tr>
 
  <tr>
    <td height="20" colspan="2">
     
      <input type="button" name="btnSave" value="Save" onclick="SaveData()"> &nbsp;   <input type="button" name="btnBack" value="Back" onclick="javascript:history.back()">
	</td>
    </tr>
</table>
</form>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
