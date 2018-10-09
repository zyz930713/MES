<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
server.ScriptTimeout=500%>
<!--#include virtual="/Admin/FinanceSeries/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetModel_N.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
id=request.QueryString("id")
set rs1=server.CreateObject("adodb.recordset")
SQL="select * from FINANCE_SERIES where NID='"&id&"'"
rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/FinanceSeries/FormCheck.js" type="text/javascript"></script>
</head>

<body onLoad="selectedcount();if(document.form1.models_filter_notin.value!='') {refreshmodel('fromitem','document.form1.models_filter_notin.value','',document.form1.model_filter_forcedin);}">
<div id="filterDiv" style="visibility: hidden; position: absolute"><iframe id="filterFrame"></iframe></div>
<form action="/Admin/FinanceSeries/EditSeries1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="760"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy">Edit a Finance Series   </td>
</tr>
<tr>
  <td width="155" height="20"><div align="left">Series Name <span class="red">*</span> </div></td>
    <td width="599" height="20" class="red">
      <div align="left">
        <input name="seriesname" type="text" id="seriesname" value="<%=rs("SERIES_NAME")%>">
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
<!--<tr>
  <td height="20">Include Models</td>
  <td height="20"><div align="center">
    <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
		<%'FactoryRight "PM."
		'session("fromModel")=" and (PM.SERIES_ID<>'"&rs("NID")&"' or PM.SERIES_ID is null) and F.NID='"&rs("FACTORY_ID")&"' and lower(PM.ITEM_NAME) like '"&ucase(rs("SERIES_NAME"))&"%'"&factorywhereoutsideand
		'session("orderModel")=" order by PM.ITEM_NAME"%>
		<%'model_string=getModel("OPTION",null,session("fromModel"),session("orderModel"),"",idcount)%>
		<tr>
        <td height="20" class="t-t-Borrow"><div align="center">Available Models  <span id="deselectedinsert">(<%=idcount%>)</span></div></td>
        <td height="20"><div align="center">&nbsp;</div></td>
        <td height="20" class="t-t-Borrow"><div align="center">Selected Models <span id="selectedinsert"></span></div></td>
        <td><div align="center">&nbsp;</div></td>
      </tr>
      <tr>
        <td rowspan="7"><select name="fromitem" size="10" multiple id="fromitem">
		<%'=model_string%>
        </select></td>
        <td><div align="center"> <img src="/Images/Button_Add.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_move(document.form1.fromitem,document.form1.toitem);selectedcount();deselectedcount()"></div></td>
        <td rowspan="7"><select name="toitem" size="10" multiple id="toitem">
          <%'= getModel("OPTION",null," and PM.SERIES_ID='"&rs("NID")&"'"," order by PM.ITEM_NAME",null,"") %>
          </select>        </td>
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
</tr>-->
<!--<tr>
  <td height="20" colspan="2">Filter: not in
    <input name="models_filter_notin" type="text" id="models_filter_notin" onChange="refreshmodel('fromitem',this.value,'',document.form1.model_filter_forcedin)" value="<% =session("series_filter")%>" size="130"></td>
  </tr>
<tr>
  <td height="20" colspan="2">Filter: 
    <input name="model_filter_forcedin" type="checkbox" id="model_filter_forcedin" value="1" onClick="refreshmodel('fromitem','',this.value,document.form1.model_filter_forcedin)">
forced in
    <input name="models_filter_in" type="text" id="models_filter_in" onChange="refreshmodel('fromitem','',this.value,document.form1.model_filter_forcedin)" size="110"></td>
  </tr>-->
<tr>
  <td height="20">Target Yield <span class="red">*</span></td>
  <td height="20"><input name="yield" type="text" id="yield" value="<%=rs("TARGET_YIELD")%>">
    %</td>
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
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->