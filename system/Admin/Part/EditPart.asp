<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
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

<body onLoad="selectedcount2();maxintervaltip()">
<div id="insert_frame" style="position:absolute;visibility: hidden;">
<iframe id="insert_iframe" src=""></iframe></div>
<form action="/Admin/Part/EditPart1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a   Part</td>
</tr>
<tr>
  <td width="152" height="20">Part Number <span class="red">*</span> </td>
    <td width="602" height="20">
      <div align="left">
        <input name="partnumber" type="text" id="partnumber" value="<%=rs("PART_NUMBER")%>" size="50">
      </div></td>
    </tr>
<tr>
  <td height="20">Job Rule</td>
  <td height="20"><input name="job_rule" type="text" id="job_rule" value="<%=rs("JOB_RULE")%>" size="50">
    * means any single letter</td>
</tr>
<tr>
  <td height="20">Part Rule <span class="red">*</span></td>
  <td height="20"><input name="part_rule" type="text" id="part_rule" value="<%=rs("PART_RULE")%>" size="50">
    * means any single letter, use &quot;,&quot; to separate different rule. </td>
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
    <td height="20">Section <span class="red">*</span></td>
    <td height="20"><select name="section" id="section">
        <option value="">-- Select Section --</option>
		<%FactoryRight "S."%>
        <%= getSection("OPTION",rs("SECTION_ID"),factorywhereoutside,"","") %>
      </select>    </td>
  </tr>
  <tr>
    <td height="20">Initial Quantity  <span class="red">*</span></td>
    <td height="20"><input name="initial_quantity" type="text" id="initial_quantity" value="<%=rs("INITIAL_QUANTITY")%>"></td>
  </tr>
  
  <tr>
    <td height="20">Applied Stations <span class="red">*</span><br>
      (if selected lines are null, that means any line can produce this Part.)</td>
    <td height="20"><div align="center">
        <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td height="20" class="t-t-Borrow"><div align="center">Available Lines </div></td>
            <td height="20"><div align="center">&nbsp;</div></td>
            <td height="20" class="t-t-Borrow"><div align="center">Selected Lines <span id="selectedinsert"></span></div></td>
            <td><div align="center">&nbsp;</div></td>
          </tr>
          <tr>
            <td rowspan="7"><select name="fromitem1" size="10" multiple id="fromitem1">
            <%
			FactoryRight "L."
			if rs("LINES_INDEX")<>"" then
			where=" where L.NID not in ('"&replace(rs("LINES_INDEX"),",","','")&"')"&factorywhereoutsideand
			else
			where=factorywhereoutsideand
			end if%>
                <%= getLine("OPTION","",where,"","") %>
            </select></td>
            <td><div align="center"><img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem1,document.form1.toitem1);selectedcount1()"></div></td>
            <td rowspan="7"><select name="toitem1" size="10" multiple id="toitem1">
            <%
			FactoryRight "L."
			if rs("LINES_INDEX")<>"" then
			where=" where L.NID in ('"&replace(rs("LINES_INDEX"),",","','")&"')"
			%>
                <%= getLine("OPTION","",where,"","") %>
                <%end if%>
            </select></td>
            <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem1,document.form1.fromitem1);selectedcount1()"></div></td>
            <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem1,document.form1.toitem1);selectedcount1()"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem1)"> </div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
            <td><div align="center"></div></td>
          </tr>
          <tr>
            <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem1,document.form1.fromitem1);selectedcount1()"></div></td>
            <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem1)"> </div></td>
          </tr>
        </table>
    </div></td>
  </tr>
  
  <tr>
    <td height="20">Include Stations  <span class="red">*</span></td>
    <td height="20"><div align="center">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"> <div align="center">Available Stations </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td rowspan="7"><select name="fromitem2" size="10" multiple id="fromitem2">
			<%
			FactoryRight "S."
			if rs("STATIONS_INDEX")<>"" then
			where=" where S.NID not in ('"&replace(rs("STATIONS_INDEX"),",","','")&"')"&factorywhereoutsideand
			else
			where=factorywhereoutsideand
			end if
			session("filterwhere")=where%>
			<%= getStation(true,"OPTION","",where," order by S.FACTORY_ID,S.STATION_NAME","","") %>
			</select></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem2,document.form1.toitem2);selectedcount2()"></div></td>
          <td rowspan="7"><select name="toitem2" size="10" multiple id="toitem2">
		  	<%
			FactoryRight "S."
			if rs("STATIONS_INDEX")<>"" then
			where=" where S.NID in ('"&replace(rs("STATIONS_INDEX"),",","','")&"')"
			%>
			<%= getStation(true,"OPTION","",where,"",rs("STATIONS_INDEX"),"") %>
			<%end if%>
			</select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem2,document.form1.fromitem2);selectedcount2()"></div></td>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem2,document.form1.toitem2);selectedcount2()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem2,document.form1.fromitem2);selectedcount2()"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem2)"> </div></td>
        </tr>
      </table>
    </div></td>
    </tr>
  <tr>
    <td height="20">Stations to be filtered </td>
    <td height="20"><input name="key" type="text" id="key">
      <label for="Submit"></label>
      <input type="button" name="Filter" value="Filter" id="Submit" onClick="javascript:document.all.insert_iframe.src='/Admin/Part/FilterStations.asp?key='+document.form1.key.value"></td>
  </tr>
  <tr>
    <td height="20">Routine Type of Stations   <span class="red">*</span></td>
    <td height="20"><input name="stationroutine" type="radio" value="0" <%if rs("STATIONS_ROUTINE")="0" then%>checked<%end if%>> 
      Fixed 
      <input name="stationroutine" type="radio" value="1" <%if rs("STATIONS_ROUTINE")="1" then%>checked<%end if%>>
      Repeated</td>
  </tr>
  <tr>
    <td height="20">Max Interval between Stations </td>
    <td height="20"><input name="maxinterval" type="text" id="maxinterval" onChange="maxintervaltip()" value="<%=rs("MAX_INTERVAL")%>" size="60">
    minutes. <span id="maxintervalinsert"></span><br>
    Use &quot;,&quot; to split interval time, for example, &quot;10,15,12,8,10,8&quot;. 0 means unlimited. </td>
  </tr>
  <tr>
    <td height="20">Target Yield</td>
    <td height="20"><input name="yield" type="text" id="yield" value="<%=rs("TARGET_YIELD")%>">
      %</td>
  </tr>
  <tr>
    <td height="20">Meet Priority </td>
    <td height="20"><select name="priority" id="priority">
	<%for i=0 to 10%>
	<option value="<%=i%>" <%if cint(rs("MEET_PRIORITY"))=i then%>selected<%end if%>><%=i%></option>
	<%next%>
    </select>    </td>
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
