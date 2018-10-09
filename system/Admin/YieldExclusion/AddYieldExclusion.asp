<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/YieldExclusion/YieldExclusionCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
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
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/YieldExclusion/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="refreshmodel('fromitem','document.form1.models_filter_notin.value','');">
<div id="filterDiv" style="visibility: hidden; position: absolute"><iframe id="filterFrame"></iframe></div>
<form action="/Admin/YieldExclusion/AddYieldExclusion1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Add a New  Yield Exclusion </td>
</tr>
<tr>
  <td width="150" height="20">Yield Exclusion Name<span class="red">*</span> </td>
  <td width="604" height="20" class="red">
      <div align="left">
        <input name="yield_exclusion_name" type="text" id="yield_exclusion_name">
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
  <td height="20">Included Models</td>
  <td height="20"><div align="center">
    <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
      <%ModelFactoryRight
	  excluded_models=""
	  SQL="select INCLUDED_SYSTEM_ITEMS from YIELD_EXCLUSION"
	  rs.open SQL,conn,1,3
	  while not rs.eof
		  if not isnull(rs("INCLUDED_SYSTEM_ITEMS")) and rs("INCLUDED_SYSTEM_ITEMS")<>"" then
		  excluded_models=excluded_models&rs("INCLUDED_SYSTEM_ITEMS")&","
		  end if
	  rs.movenext
	  wend
	  rs.close
	  if excluded_models<>"" then
	  excluded_models=replace(excluded_models,",","','")
	  end if%>
      <%
	  fromModel=" and segment1 like '%-_____-%' and segment1 not in ('"&excluded_models&"')"&modelfactoryand
	  orderModel=" order by segment1"
	  session("fromModel")=fromModel
	  session("orderModel")=orderModel
	  strmodel= getModel("OPTION",""," top 100 ",fromModel,orderModel,"",idcount) %>
	  <tr>
        <td height="20" class="t-t-Borrow"><div align="center">Available Models <span id="deselectedinsert">(<%=idcount%>)</span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center">Selected Models <span id="selectedinsert"></span></div></td>
        <td><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="7"><span id="insert_stations">
          <select name="fromitem" size="10" multiple id="fromitem">
		  <%=strmodel%>
          </select>
        </span></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount();deselectedcount()"></div></td>
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
  <td height="20" colspan="2">Filter: not in 
    <input name="models_filter_notin" type="text" id="models_filter_notin" onChange="refreshmodel('fromitem',this.value,'')" value="<% =session("yieldexclusion_filter")%>" size="130"></td>
  </tr>
<tr>
  <td height="20" colspan="2">Filter: in 
    <input name="models_filter_in" type="text" id="models_filter_in" onChange="refreshmodel('fromitem','',this.value)" size="130"></td>
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
<input type="reset" name="Submit7" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->