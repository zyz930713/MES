<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Machine/MachineCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from MACHINE where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Machine/FormCheck.js" type="text/javascript"></script>
</head>

<body>

<form action="/Admin/Machine/EditMachine1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Machine  </td>
</tr>
<tr>
  <td width="118" height="20"><div align="left">Machine Number <span class="red">*</span> </div></td>
  <td width="636" height="20" class="red">
    <div align="left">
      <input name="machinenumber" type="text" id="machinenumber" value="<%=rs("MACHINE_NUMBER")%>">
  </div></td>
</tr>
<tr>
  <td height="20"><div align="left">Machine Name <span class="red">*</span> </div></td>
  <td height="20"><input name="machinename" type="text" id="machinename" value="<%=rs("MACHINE_NAME")%>"></td>
</tr>
<tr>
  <td height="20">Belonged Factory <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value="">-- Select Factory --</option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION",rs("FACTORY_ID"),factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20">Is Locked </td>
  <td height="20"><input name="locked" type="checkbox" id="locked" value="1" <%if rs("LOCKED")="1" then%>checked<%end if%>></td>
</tr>
<tr>
  <td height="20">Stations to be used</td>
  <td height="20"><table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" class="t-t-Borrow"><div align="center">Available Stations </div></td>
      <td height="20"><div align="center"></div></td>
      <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
      <td><div align="center"></div></td>
    </tr>
    <tr>
      <td rowspan="7"><select name="fromitem" size="10" id="fromitem">
        <%
		FactoryRight "S."
		if rs("STATIONS_USED")<>"" then
			where=" where S.NID not in ('"&replace(rs("STATIONS_USED"),",","','")&"') and S.NID<>'"&id&"' "&factorywhereoutsideand
			else
			where=" where S.NID<>'"&id&"' "&factorywhereoutsideand 
			end if%>
        <%= getStation(true,"OPTION","",where,"","","") %>
      </select></td>
      <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount()"></div></td>
      <td rowspan="7"><select name="toitem" size="10" multiple id="toitem">
        <%if rs("STATIONS_USED")<>"" then
			where=" where S.NID in ('"&replace(rs("STATIONS_USED"),",","','")&"')"
			%>
        <%= getStation(true,"OPTION","",where,"",rs("STATIONS_USED"),"") %>
        <%end if%>
      </select></td>
      <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
    </tr>
    <tr>
      <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem);selectedcount()"></div></td>
      <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
    </tr>
    <tr>
      <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem);selectedcount()"></div></td>
      <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
    </tr>
    <tr>
      <td><div align="center"></div></td>
      <td><div align="center"></div></td>
    </tr>
    <tr>
      <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem);selectedcount()"></div></td>
      <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
    </tr>
  </table></td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="stationscount" type="hidden" id="stationscount">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
