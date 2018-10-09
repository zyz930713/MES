<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/SeriesGroup/SeriesGroupCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from SERIES_GROUP where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/SeriesGroup/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="deselectedcount()">
<form action="/Admin/SeriesGroup/EditSeriesGroup1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Series Group   </td>
</tr>
<tr>
  <td width="153" height="20"><div align="left">Series Group Name <span class="red">*</span> </div></td>
    <td width="601" height="20" class="red">
      <div align="left">
        <input name="seriesgroupname" type="text" id="seriesgroupname" value="<%=rs("SERIES_GROUP_NAME")%>">
      </div></td>
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
  </select>  </td>
</tr>
<tr>
  <td height="20">Include Series </td>
  <td height="20"><div align="center">
    <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
		  <tr>
        <td height="20" class="t-t-Borrow"><div align="center">Available Series <span id="deselectedinsert"></span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center">Selected Sereis <span id="selectedinsert"></span></div></td>
        <td><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="7">
          <select name="fromitem" size="10" multiple id="fromitem">
		  <%=getSeries("OPTION","","where S.NID not in ('"&replace(rs("INCLUDED_SERIES"),",","','")&"')"," order by SERIES_NAME","")%>
          </select></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount();deselectedcount()"></div></td>
        <td rowspan="7"><select name="toitem" size="10" multiple id="toitem">
		<%=getSeries("OPTION",""," where S.NID in ('"&replace(rs("INCLUDED_SERIES"),",","','")&"')"," order by S.SERIES_NAME","")%>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.toitem)"> </div></td>
      </tr>
      <tr>
        <td><div align="center"></div></td>
        <td><div align="center"></div></td>
      </tr>
      <tr>
        <td><div align="center"><img src="/Images/Button_Remove.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.toitem,document.form1.fromitem);selectedcount();deselectedcount()"></div></td>
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
  <td height="20">Included Parts</td>
  <td height="20"><%=highlightsamestring(rs("INCLUDED_SYSTEM_ITEMS"),",")%>&nbsp;</td>
</tr>
<tr>
  <td height="20">Excepted from Overall </td>
  <td height="20"><input name="overall_exception" type="checkbox" id="overall_exception" value="1" <%if rs("OVERALL_EXCEPTION")="1" then%>checked<%end if%>></td>
</tr>
<tr>
  <td height="20">Lead Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("LEAD_TIME")),newtime)%>
    <input name="lead_time" type="text" id="lead_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="lead_time_unit" id="lead_time_unit">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
<tr>
  <td height="20">WIP Time </td>
  <td height="20"><%unit=unitconvert(csng(rs("WIP_TIME")),newtime)%>
    <input name="wip_time" type="text" id="wip_time" onChange="numbercheck(this)" value="<%=newtime%>" size="4">
    <select name="wip_time_unit" id="wip_time_unit">
      <option value="">-- Select--</option>
      <option value="MM" <%if unit="MM" then%>selected<%end if%>>Minutes</option>
      <option value="HH" <%if unit="HH" then%>selected<%end if%>>Hours</option>
      <option value="DD" <%if unit="DD" then%>selected<%end if%>>Days</option>
    </select></td>
</tr>
<tr>
  <td height="20">Target of First-pased Yield<span class="red">*</span></td>
  <td height="20"><input name="first_yield" type="text" id="first_yield" value="<%=rs("TARGET_FIRSTYIELD")%>">
    %</td>
</tr>
<tr>
  <td height="20">Target of Internal Yield  <span class="red">*</span></td>
  <td height="20"><input name="internal_yield" type="text" id="internal_yield" value="<%=csng(rs("TARGET_INTERNALYIELD"))*100%>">
%</td>
</tr>
<tr>
  <td height="20">Target of Retest Yield <span class="red">*</span></td>
  <td height="20"><input name="inspect_yield" type="text" id="inspect_yield" value="<%=csng(rs("TARGET_INSPECTYIELD"))*100%>">
    %</td>
</tr>
<tr>
  <td height="20">Target of Final Yield <span class="red">*</span></td>
  <td height="20"><input name="final_yield" type="text" id="final_yield" value="<%=rs("TARGET_YIELD")%>">
    %</td>
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
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="Submit" value="Update">
&nbsp;
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
