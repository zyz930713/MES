<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
 
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%

LaborRate=request("txtLaborRate")
OH=request("txtOH")
action=request.QueryString("Action")
if(action="") then
	SQL="select * FROM LABORRATE_OH"
	set rsLABORRATE_OH=Server.CreateObject("adodb.recordset")
	rsLABORRATE_OH.open SQL,conn,1,3
	
	if(rsLABORRATE_OH.recordcount>0) then
		LaborRate=rsLABORRATE_OH("LaborRate")
		OH=rsLABORRATE_OH("OH")
	end if 
end if 
if(action="2") then
	SQL="DELETE FROM LABORRATE_OH"
	set rsLABORRATE_OH1=Server.CreateObject("adodb.recordset")
	rsLABORRATE_OH1.open SQL,conn,1,3
	
	SQL="INSERT INTO LABORRATE_OH(LABORRATE,OH) VALUES('"+LaborRate+"','"+OH+"')"
	set rsLABORRATE_OH2=Server.CreateObject("adodb.recordset")
	rsLABORRATE_OH2.open SQL,conn,1,3
	
	
	response.write "<script>alert('Insert Successfully!')</script>"
end if 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	 
	function SaveData()
	{
		if(document.getElementById("txtLaborRate")=="")
		{
			alert("please input Labor Rate!");
			return;
		}
		if(document.getElementById("txtOH")=="")
		{
			alert("please input OH!");
			return;
		}
		form1.action="LABORRATE_OH.asp?Action=2";
		form1.submit();
	}
</script>
</head>

<body onLoad="language();language_page();language_jobnote()">
<form method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" class="t-b-midautumn"  colspan="2">Labor Rate and OH Setting</td>
    </tr>
    <tr>
      <td height="20" style="width:10%">Labor Rate</td>
	   <td height="20"><input type="text" id="txtLaborRate" name="txtLaborRate" value="<%=LaborRate%>"></td>
    </tr>
	 <tr>
      <td height="20"  style="width:10%">OH</td>
	   <td height="20" ><input type="text" id="txtOH" name="txtOH" value="<%=OH%>"></td>
  
    </tr>
	 
	  <tr>
      <td height="20" colspan="2">
	  	<input type="button" id="btnSave" name="btnSave" value="Save" onClick="SaveData()">
	  </td>
    </tr>
</table>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_TICKET_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
