<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->
<!--#include virtual="/Functions/GetLine.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
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
<script language="JavaScript" src="/Admin/SubSeries_New/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
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
</script>
</head>

<body onLoad="language(<%=session("language")%>);">
<div id="filterDiv" style="visibility: hidden; position: absolute"><iframe id="filterFrame"></iframe></div>
<form action="/Admin/SubSeries_New/AddSeries1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span></td>
</tr>
<tr> 
	<td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:<% =session("User") %></td>
</tr>
<tr>
  <td width="20%" height="20"><span id="td_SubSeries"></span> <span class="red">*</span> </td>
  <td  height="20" class="red">
      <div align="left">
        <input name="seriesname" type="text" id="seriesname">
      </div></td>
    </tr>
<tr>
  <td height="20"><span id="td_Factory"></span> <span class="red">*</span></td>
  <td height="20"><select name="factory" id="factory">
    <option value="">-- Select Factory --</option>
	<%FactoryRight ""%>
    <%= getFactory("OPTION","",factorywhereinside,"","") %>
  </select></td>
</tr>
<tr>
  <td height="20"><span id="td_Section"></span>  <span class="red">*</span></td>
  <td height="20"><select name="section" id="section">
    <option value=""></option>
    <%FactoryRight "S."%>
    <%= getSection("OPTION","",factorywhereoutside,"","") %>
  </select>  </td>
</tr>
<tr>
  <td height="20"><span id="inner_Line"></span> <span class="red">*</span></td>
  <td height="20"><select name="line" id="line">
  <option value=""></option>
    <%FactoryRight "L."%>
    <%= getLine("OPTION","",factorywhereoutsideand," order by LINE_NAME","") %>
  </select></td>
</tr>

 
<tr>
  <td height="20"><span id="td_Series"></span> <span class="red">*</span></td>
  <td height="20"><select name="Series" id="Series">
    <option value=""></option>
	<%FactoryRight "S."%>
    <%= getSeries("OPTION","",factorywhereoutside,"","") %>
  </select>  </td>
</tr>

<tr>
	<td height="20"><span id="td_Model"></span></td>
	<td><textarea name="txtIncludeModel" cols="100%" rows="10"  id="txtIncludeModel" readonly="true"></textarea></td>
</tr>
<Tr>
	<Td>&nbsp;</Td>
	<td> <input name="btnSelModel" type="Button" id="btnSelModel" value="Select Model" onclick="window.showModalDialog('ModelSelect.asp?models='+document.all.txtIncludeModel.value,window,'dialogHeight:500px;dialogWidth:600px');">
		<input name="txtRemoveModel" type="text" id="txtRemoveModel" value="">
		<input name="btnRemvModel" type="Button" id="btnRemvModel" value="Remove Model" onclick="fRemoveModel()">
		<input name="btnRemvAllModel" type="Button" id="btnRemvAllModel" value="Remove All Models" onclick="RemoveaLLModelS()">
	</td>
</Tr>
<tr>
  <td height="20"><span id="td_FirstPassYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="firstyield" type="text" id="firstyield" value="0">
%</td>
</tr>
<tr>
  <td height="20"><span id="td_InternalYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="internal_yield" type="text" id="internal_yield" value="0">
    %</td>
</tr>
<tr>
  <td height="20"><span id="td_TargetYield"></span> <span class="red">*</span></td>
  <td height="20"><input name="yield" type="text" id="yield" value="0">
    %</td>
</tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
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