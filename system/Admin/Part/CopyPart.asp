<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
randomize 
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
SQL="select * from PART where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Part/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="selectedcount();maxintervaltip()">
<form action="/Admin/Part/CopyPart1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Copy a part</td>
</tr>
<tr>
  <td width="175" height="20">Part Number <span class="red">*</span> </td>
  <td width="579" height="20" class="red">
      <div align="left">
        <input name="partnumber" type="text" id="partnumber" value="<%=rs("PART_NUMBER")%>-<%=cint(10000*Rnd)%>">
      </div></td>
    </tr>
  <tr>
    <td height="20">Part Rule <span class="red">*</span></td>
    <td height="20"><input name="part_rule" type="text" id="part_rule" value="<%=rs("PART_RULE")%>">
* means any single letter. </td>
  </tr>
  <tr>
    <td height="20">Belonged Factory <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
        <option value="">-- Select Section --</option>
        <%= getFactory("OPTION",rs("FACTORY_ID"),factoryiteminside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20">Section <span class="red">*</span></td>
    <td height="20"><select name="section" id="section">
        <option value="">-- Select Section --</option>
        <%= getSection("OPTION",rs("SECTION_ID"),"","","") %>
      </select>    </td>
  </tr>
  <tr>
    <td height="20">Initial Quantity <span class="red">*</span></td>
    <td height="20"><input name="initial_quantity" type="text" id="initial_quantity" value="<%=rs("INITIAL_QUANTITY")%>"></td>
  </tr>
  
  <tr>
    <td height="20">Include Stations  <span class="red">*</span></td>
    <td height="20"><div align="center">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"> <div align="center">Available Stations </div></td>
          <td height="20"><div align="center"></div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td rowspan="7"><select name="fromitem" size="10" id="fromitem">
			<%if rs("STATIONS_INDEX")<>"" then
			where=" where NID not in ('"&replace(rs("STATIONS_INDEX"),",","','")&"')"
			else
			where=""
			end if%>
			<%= getStation(true,"OPTION","",where,"","","") %>
			</select></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount()"></div></td>
          <td rowspan="7"><select name="toitem" size="10" multiple id="toitem">
		  	<%if rs("STATIONS_INDEX")<>"" then
			where=" where NID in ('"&replace(rs("STATIONS_INDEX"),",","','")&"')"
			%>
			<%= getStation(true,"OPTION","",where,"",rs("STATIONS_INDEX"),"") %>
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
      </table>
    </div></td>
    </tr>
  <tr>
    <td height="20">Routine Type of Stations <span class="red">*</span></td>
    <td height="20"><input name="stationroutine" type="radio" value="0" checked>
      Fixed
      <input name="stationroutine" type="radio" value="1">
      Repeated</td>
  </tr>
  <tr>
    <td height="20">Max Interval between Stations </td>
    <td height="20"><input name="maxinterval" type="text" id="maxinterval" onChange="maxintervaltip()" value="<%=rs("MAX_INTERVAL")%>" size="60">
    minutes. <span id="maxintervalinsert"></span><br>
    Use &quot;,&quot; to split interval time, for example, &quot;10,15,12,8,10,8&quot;. 0 means unlimited. </td>
  </tr>
  <tr>
    <td height="20">Copy Options</td>
    <td height="20"><input name="allstations" type="checkbox" id="allstations" value="checkbox">
      All Stations 
      <input name="allactions" type="checkbox" id="allactions" value="checkbox">
      All Actions </td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="id" type="hidden" id="id" value="<%=id%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="stationscount" type="hidden" id="stationscount">
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
<!--#include virtual="/WOCF/BOCF_Close.asp" -->