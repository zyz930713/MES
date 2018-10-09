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
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
'GUID
strGUID="GU"&NID_SEQ("GUID")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Part/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language="javascript">
	function setAction()
	{	
		var text=document.getElementById("toitem2").value;		 
		var GUID=document.getElementById("txtGuid").value;
		
		window.open("SetActionDefectCode.asp?StationId="+text+"&GUID="+GUID,"","'height=150, width=900, top=150, left=250, toolbar=no, menubar=no, scrollbars=yes, resizable=no,location=no, status=no");
	}
</script>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="language(<%=session("language")%>);">
<div id="insert_frame" style="position:absolute;visibility: hidden;"><iframe id="insert_iframe" src=""></iframe></div>
<form action="/Admin/Part/AddRouting_New1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
</tr>
    <tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>
<tr>
  <td width="177" height="20"><span id="td_RoutingNumber"></span> <span class="red">*</span> </td>
  <td width="577" height="20" class="red">
      <div align="left">
        <input name="partnumber" type="text" id="partnumber" size="50">
      </div></td>
    </tr>
	<!--
  <tr>
    <td height="20">Job Rule </td>
    <td height="20"><input name="job_rule" type="text" id="job_rule" size="50">
      * means any single letter</td>
  </tr>
  -->
  <tr>
    <td height="20"><span id="td_RoutingRule"></span>   <span class="red">*</span></td>
    <td height="20"><input name="part_rule" type="text" id="part_rule" size="50"> 
    * means any single letter, use"," to separate different rule. </td>
  </tr>
  <tr>
  	<td height="20"><span id="td_RoutingType"></span>  <span class="red">*</span></td>
	<td height="20">
		<select name="partType" id="factory">
		  <option value=""></option>
		  <option value="0">Normal</option>
		  <option value="1">Rework</option>
		  <option value="4">Retest</option>
		  <!--
		  <option value="5">Slapping</option>
		  -->
		</select>
	</td>
  </tr>
  <tr>
    <td height="20"><span id="td_Factory"></span>  <span class="red">*</span></td>
    <td height="20"><select name="factory" id="factory">
      <option value=""></option>
	  <%FactoryRight ""%>
      <%= getFactory("OPTION","",factorywhereinside,"","") %>
    </select></td>
  </tr>
  <tr>
    <td height="20"><span id="td_Section"></span> <span class="red">*</span></td>
    <td height="20"><select name="section" id="section">
	<option value=""></option>
	<%FactoryRight "S."%>
	<%= getSection("OPTION","",factorywhereoutside,"","") %>
    </select>    </td>
  </tr>
  <tr>
    <td height="20"><span id="td_IniQty"></span> <span class="red">*</span></td>
    <td height="20"><input name="initial_quantity" type="text" id="initial_quantity" ></td>
  </tr>
  <tr>
    <td height="20"><p><span id="td_LineName"></span> <span class="red">*</span></p>      </td>
    <td height="20">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
      <tr>
        <td height="20" class="t-t-Borrow"><div align="center">Available Lines </div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center">Selected Lines <span id="selectedinsert"></span></div></td>
        <td><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="7"><select name="fromitem1" size="10" multiple id="fromitem1">
            <%FactoryRight "L."%>
            <%= getLine("OPTION","",factorywhereoutsideand,"","") %>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem1,document.form1.toitem1)"></div></td>
        <td rowspan="7"><select name="toitem1" size="10" multiple id="toitem1">
                        </select></td>
        <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem1)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem1,document.form1.fromitem1)"></div></td>
        <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem1)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem1,document.form1.toitem1)"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem1)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem1,document.form1.fromitem1)"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem1)"> </div></td>
      </tr>
    </table>
	if selected lines are null, that means any line can produce this Part.
	</td>
  </tr>
  
  <tr>
    <td height="20"><span id="inner_Stations"></span> <span class="red">*</span></td>
    <td height="20"><div align="center">
      <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
        <tr>
          <td height="20" class="t-t-Borrow"> <div align="center">Available Stations </div></td>
          <td height="20"><div align="center">&nbsp;</div></td>
          <td height="20" class="t-t-Borrow"><div align="center">Selected Stations <span id="selectedinsert"></span></div></td>
          <td><div align="center">&nbsp;</div></td>
        </tr>
        <tr>
          <td rowspan="7"><span id="insert_stations"><select name="fromitem2" size="10" multiple id="fromitem2">
		    <%FactoryRight "S."
			session("filterwhere")=factorywhereoutsideand%>
			<%= getStation_New(true,"OPTION","",factorywhereoutsideand," order by S.FACTORY_ID,S.STATION_NUMBER","","") %>
			</select></span></td>
          <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem2,document.form1.toitem2)"></div></td>
          <td rowspan="7"><select name="toitem2" size="10" multiple id="toitem2" ondblclick="javascript:setAction();">
          </select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem2,document.form1.fromitem2)"></div></td>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem2,document.form1.toitem2)"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem2)"> </div></td>
        </tr>
        <tr>
          <td><div align="center"></div></td>
          <td><div align="center"></div></td>
        </tr>
        <tr>
          <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem2,document.form1.fromitem2)"></div></td>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem2)"> </div></td>
        </tr>
      </table>
    </div></td>
    </tr>
 
  <tr>
    <td height="20"><span id="td_StationSeq"></span> <span class="red">*</span></td>
    <td height="20"><input name="stationroutine" type="radio" value="0" checked>
      Fixed
      <input name="stationroutine" type="radio" value="1">
      Repeated</td>
  </tr>
  <tr>
    <td height="20"><span id="td_MaxInterval"></span></td>
    <td height="20"><input name="maxinterval" type="text" id="maxinterval" size="60" onChange="maxintervaltip()">
    minutes. <span id="maxintervalinsert"></span> <br>
    Use &quot;,&quot; to split interval time, for example, &quot;10,15,12,8,10,8&quot;. 0 means unlimited.</td>
  </tr>
  <tr>
    <td height="20"><span id="td_TargetYield"></span></td>
    <td height="20"><input name="yield" type="text" id="yield">
      %</td>
  </tr>
  <tr>
    <td height="20"><span id="td_Priority"></span></td>
    <td height="20"><select name="priority" id="priority">
        <%for i=0 to 10%>
        <option value="<%=i%>"><%=i%></option>
        <%next%>
      </select>    </td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="stationscount" type="hidden" id="stationscount">
	  <input type="hidden" name="txtGuid" value="<%=strGUID%>"  />
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK" >
&nbsp;
<input type="reset" name="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->