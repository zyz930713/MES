<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/FAMILY/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="deselectedcount()">
<form action="/Admin/FAMILY/AddSeriesGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Add a New Series Group </td>
</tr>
<tr>
  <td width="161" height="20">Series Group Name <span class="red">*</span> </td>
  <td width="593" height="20">
      <div align="left">
        <input name="seriesgroupname" type="text" id="seriesgroupname" size="50">
      </div></td>
    </tr>
<tr>
  <td height="20">Belonged Factory <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value="">-- Select Factory --</option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION","",factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20">Section <span class="red">*</span></td>
  <td height="20"><select name="section" id="section">
    <option value="">-- Select Section --</option>
    <%FactoryRight "S."%>
    <%= getSection("OPTION","",factorywhereoutside,"","") %>
  </select>  </td>
</tr>
<tr>
  <td height="20">Included Series </td>
  <td height="20"><div align="center">
    <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
	  <tr>
        <td height="20" class="t-t-Borrow"><div align="center">Available Sereis <span id="deselectedinsert"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center">Selected Series <span id="selectedinsert"></span></div></td>
        <td><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="7"><span id="insert_stations">
          <select name="fromitem" size="10" multiple id="fromitem">
		  <%=getSeries("OPTION","",""," order by SERIES_NAME","")%>
          </select>
        </span></td>
        <td><div align="center"><img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount();deselectedcount()"></div></td>
        <td rowspan="7"><select name="toitem" size="10" multiple id="toitem">
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem);selectedcount();deselectedcount()"></div></td>
        <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"> <img src="/Images/Button_Add_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.fromitem,document.form1.toitem);selectedcount();deselectedcount()"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"><img src="/Images/Button_Remove_All.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_all(document.form1.toitem,document.form1.fromitem);selectedcount();deselectedcount()"></div></td>
        <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.toitem)"> </div></td>
      </tr>
    </table>
  </div></td>
</tr>

<tr>
  <td height="20">Excepted from Overall </td>
  <td height="20"><input name="overall_exception" type="checkbox" id="overall_exception" value="1"></td>
</tr>
<tr>
  <td height="20">Lead Time </td>
  <td height="20"><input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" size="4">
      <select name="lead_time_unit" id="lead_time_unit">
        <option value="">-- Select--</option>
        <option value="MM">Minutes</option>
        <option value="HH">Hours</option>
        <option value="DD">Days</option>
    </select></td>
</tr>
<tr>
  <td height="20">WIP Time </td>
  <td height="20"><input name="wip_time" type="text" id="wip_time" onChange="numbercheck(this)" size="4">
    <select name="wip_time_unit" id="wip_time_unit">
      <option value="">-- Select--</option>
      <option value="MM">Minutes</option>
      <option value="HH">Hours</option>
      <option value="DD">Days</option>
    </select></td>
</tr>
<tr>
  <td height="20">Target of First-pased Yield <span class="red">*</span></td>
  <td height="20"><input name="first_yield" type="text" id="first_yield" value="0">
    %</td>
</tr>
<tr>
  <td height="20">Target of Internal Yiled  <span class="red">*</span></td>
  <td height="20"><input name="internal_yield" type="text" id="internal_yield" value="0">
    %</td>
</tr>
<tr>
  <td height="20">Target of Retest Yield <span class="red">*</span></td>
  <td height="20"><input name="inspect_yield" type="text" id="inspect_yield" value="0">
    %</td>
</tr>
<tr>
  <td height="20">Target of Final Yield <span class="red">*</span></td>
  <td height="20"><input name="final_yield" type="text" id="final_yield" value="0">
    %</td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Save">
&nbsp;
<input name="Reset" type="reset" id="Reset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->