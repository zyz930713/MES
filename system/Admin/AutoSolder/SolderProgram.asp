<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
TERMINAL_PART=request.QueryString("TERMINAL_PART")
Program_Name=""

if TERMINAL_PART<>"" then
	SQLS="Select * from SOLDER_PROGRAM WHERE upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"' order by TERMINAL_PART"
	rs.open SQLS,conn,1,3
	if(rs.recordcount>0) then
		Program_Name=rs("PROGRAM_NUMBER").value
	end if 
	rs.close
end if 

if(request.QueryString("action")="1") then
	TERMINAL_PART=request.Form("txtTerminalPart")
	Program_Name=request.Form("txtProgramName")

	INCLUDEMODEL=request.Form("txtIncludeModel")
	set rs0=server.createobject("adodb.recordset")
	SQL="DELETE FROM SOLDER_PROGRAM WHERE  upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"'"
	rs0.open SQL,conn,1,3

	set rs1=server.createobject("adodb.recordset")
	SQL="INSERT INTO SOLDER_PROGRAM VALUES('"&ucase(TERMINAL_PART)&"','"&ucase(Program_Name)&"')"
	rs1.open SQL,conn,1,3

	set rs2=server.createobject("adodb.recordset")
	SQL="DELETE FROM SOLDER_PART WHERE  upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"'"
	rs2.open SQL,conn,1,3

	arrModel=split(INCLUDEMODEL,",")
	for i=0 to ubound(arrModel)
		if arrModel(i)<>"" then
			set rs3=server.createobject("adodb.recordset")
			SQL="INSERT INTO SOLDER_PART VALUES('"&ucase(arrModel(i))&"','"&ucase(TERMINAL_PART)&"')"
			rs3.open SQL,conn,1,3
		end if 
	next 
	response.write "<script>window.alert('Successfully Save');location.href='solderlist.asp'</script>"
end if 

function getIncludeModel(TERMINAL_PART)
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select * from SOLDER_PART WHERE upper(TERMINAL_PART)='"& ucase(TERMINAL_PART) &"' order by FG_PART"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		output=output& rsS("FG_PART").value &","
		rsS.movenext
	wend
end if
getIncludeModel=output
rsS.close
set rsS=nothing
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script>
	function fRemoveModel()
	{
		if(document.getElementById("txtRemoveModel").value!="")
		{
			var removeModelStr;
			removeModelStr=document.getElementById("txtRemoveModel").value;
			document.getElementById("txtIncludeModel").value=document.getElementById("txtIncludeModel").value.replace(","+removeModelStr,"").replace(removeModelStr+",","").replace(+removeModelStr,"")
			window.alert(removeModelStr+" is removed");
			document.getElementById("txtRemoveModel").value="";
		}	
		else
		{
			window.alert("Please key in the model you want to remove!");
		}
		
	}
	
	function RemoveaLLModelS()
	{
		if(window.confirm("Do you confirm remove all models you selected?"))
		{
			 
			document.getElementById("txtIncludeModel").value="";
		}	
		 
		
	}
	
	function SaveData()
	{
		if(document.getElementById("txtTerminalPart").value=="")
		{
			window.alert("Please key in Terminal Part!");
			return;
		}	
	 
	 	if(document.getElementById("txtProgramName").value=="")
		{
			window.alert("Please key in Program Name");
			return;
		}	
			if(document.getElementById("txtIncludeModel").value=="")
		{
			window.alert("Please select Model");
			return;
		}	

		form1.action="SolderProgram.asp?action=1";
		form1.submit();
	}
</script>
</head>
<body>
<form name="form1" id="form1" method="post" action="SolderProgram.asp">
	<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr>
	  <td height="20" colspan="2" class="t-c-greenCopy">Solder Program</td>
	</tr>
	<tr>
	  <td width="127" height="20">Terminal Part<span class="red">*</span> </td>
	  <td width="627" height="20" class="red">
			<input name="txtTerminalPart" id="txtTerminalPart" type="text"  value="<%=TERMINAL_PART%>">
		</tr>
	<tr>
	  <td height="20">Program Name<span class="red">*</span></td>
	  <td height="20">
			<input name="txtProgramName" type="text" id="txtProgramName" value="<%=Program_Name%>" >
	  </td>
	</tr>
	<Tr>
		<Td>&nbsp;</Td>
		<td> <input name="SelectModel" type="Button" id="SelectModel" value="Select Model" onclick="window.showModalDialog('../Subseries_New/ModelSelect.asp',window,'dialogHeight:600px;dialogWidth:500px');">
			<input name="removeModel" type="Button" id="removeModel" value="Remove Model" onclick="fRemoveModel()"><input name="txtRemoveModel" type="text" id="txtRemoveModel" value="">
			<input name="removeModels" type="Button" id="removeModels" value="Remove All Models" onclick="RemoveaLLModelS()">
		</td>
	</Tr>
	<tr>
		<td height="20">Included Model</td>
		<td><textarea name="txtIncludeModel" cols="100%" rows="10"  id="txtIncludeModel"  readonly="true"><%=getIncludeModel(TERMINAL_PART)%></textarea></td>
	</tr>
	  <tr>
		<td height="20" colspan="2"><div align="center">
		  <input type="button" name="button" value="Save" onclick="SaveData()">
		  <input type="reset" name="Submit7" value="Reset">
	</div></td>
		</tr>
	</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->