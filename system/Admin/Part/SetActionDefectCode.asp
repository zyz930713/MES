<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
StationId=request.QueryString("StationId")
strGUID=request.QueryString("GUID")
RoutingId=request("rid")
task=request.QueryString("task")
word=request.QueryString("word")
ActionCount=1
DefectCount=1
CategoryCount=1
if(word<>"") then
	response.write ("<script>alert('"+word+"')</script>")
end if 
%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetDefectCode.asp" -->
<!--#include virtual="/Functions/GetAction.asp" -->
<!--#include virtual="/Functions/GetCategory.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Barcode System - Scan</title>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css" />
<script language="javascript">
var intRowIndex = 0;
function reload(tbIndex1,tbIndex2)
{
	if (parseInt(tbIndex1)>5)
	{
		document.form1.deleteAction.disabled="";
	}
	else
	{
		document.form1.deleteAction.disabled="disabled";
	}
	
	if (parseInt(tbIndex2)>4)
	{
		document.form1.deleteDefectCode.disabled="";
	}
	else
	{
		document.form1.deleteDefectCode.disabled="disabled";
	}
}
function insertActionRow(tbIndex)
{ 
	 
	if (parseInt(tbIndex)>=5)
	{
		document.form1.deleteAction.disabled="";
	}
	var ActionCount=parseInt(document.getElementById("txtActionCount").value);
	ActionCount=ActionCount+1;
	document.getElementById("txtActionCount").value=ActionCount
	var objRow = document.getElementById("ActionTable").insertRow(tbIndex);
	var objCel = objRow.insertCell(0);
	objCel.innerHTML = "  <select id='Action"+ActionCount+"'  style='width:350px'><option>-- Select Action --</option><%=getAction_New(true,"OPTION","",""," order by Action_CODE ","","")%></select>";
	objRow.attachEvent("onclick",getIndex);
}

function insertDefectRow(tbIndex)
{
	if (parseInt(tbIndex)>=4)
	{
		document.form1.deleteDefectCode.disabled="";
	}
	var DefectCount=parseInt(document.getElementById("txtDefectCount").value);
	 DefectCount=DefectCount+1;
	 document.getElementById("txtDefectCount").value=DefectCount
	 var objRow = DefectTable.insertRow(tbIndex);
	 var objCel = objRow.insertCell(0);
objCel.innerHTML = "<select  id='Defect"+DefectCount+"'  style='width:350px'><option value=''>-- Select Defect Code --</option><%=getDefectCode_New("OPTION","",""," order by Defect_Code ","") %></select>"; 
	 
	 
	  var objCel = objRow.insertCell(1);
	 objCel.innerHTML = "  <select id='TansactionType"+DefectCount+"'  style='width:100px'><!--<option value='0' selected='selected'>None</option><option value='1'>Rework</option>--><option value='2'>Scrap</option><!--<option value='3'>Readjust</option><option value='4'>Change Model</option><option value='5'>Slapping</option>--></select>";
 	 objRow.attachEvent("onclick",getIndex);
}

function deleteActionRow(tbIndex)
{ 
	
	if (parseInt(tbIndex)<=5)
	{
		document.form1.deleteAction.disabled="disabled";
	}
	var ActionCount=parseInt(document.getElementById("txtActionCount").value);
	ActionCount=ActionCount-1;
	document.getElementById("txtActionCount").value=ActionCount
 	document.getElementById("ActionTable").deleteRow(tbIndex);
}
function deleteDefectRow(tbIndex)
{

	if (parseInt(tbIndex)<=4)
	{
		document.form1.deleteDefectCode.disabled="disabled";
	}	
	var DefectCount=parseInt(document.getElementById("txtDefectCount").value);
	DefectCount=DefectCount-1;
	document.getElementById("txtDefectCount").value=DefectCount
 	DefectTable.deleteRow(tbIndex);
	//alert(DefectCount)
}


//add category
function insertCategoryRow(tbIndex)
{
	if (parseInt(tbIndex)>=4)
	{
		document.form1.deleteCategory.disabled="";
	}
	var CategoryCount=parseInt(document.getElementById("txtCategoryCount").value);
	 CategoryCount=CategoryCount+1;
	 document.getElementById("txtCategoryCount").value=CategoryCount
	 var objRow = CategoryTable.insertRow(tbIndex);
	 var objCel = objRow.insertCell(0);
objCel.innerHTML = "<select  id='CATEGORY"+CategoryCount+"'  style='width:350px'><option value=''>-- Select Category --</option><%= getCategory("OPTION","",""," order by CATEGORY_NAME ")%></select>"; 

 	 objRow.attachEvent("onclick",getIndex);
}


function deleteCategoryRow(tbIndex)
{

	if (parseInt(tbIndex)<=4)
	{
		document.form1.deleteCategory.disabled="disabled";
	}	
	var CategoryCount=parseInt(document.getElementById("txtCategoryCount").value);
	CategoryCount=CategoryCount-1;
	document.getElementById("txtCategoryCount").value=CategoryCount
 	CategoryTable.deleteRow(tbIndex);
	//alert(DefectCount)
}



function getIndex()
{
	 intRowIndex = event.srcElement.parentElement.rowIndex;
	 pos.innerText = intRowIndex;
}
function getSubmit()
{

	var ActionCount=parseInt(document.getElementById("txtActionCount").value);
	var DefectCount=parseInt(document.getElementById("txtDefectCount").value);
	var CategoryCount=parseInt(document.getElementById("txtCategoryCount").value);
	
	var actionString="";
	var defectString="";
	var categoryString="";
	
	var typeString="";
	for(var i=1;i<=ActionCount;i++)
	{
		var q=1;
		var Aname="Action"+i;
		var obj=document.getElementById(Aname)
		var index=obj.selectedIndex;
		var selectTextA=obj.options[index].text;
		if (obj.value!="")
		{
			for(q=i+1;q<=ActionCount;q++)
			{
				var Bname="Action"+q;
				var objB=document.getElementById(Bname)
				var indexB=objB.selectedIndex;
				var selectTextB=objB.options[indexB].text;
				if(selectTextA==selectTextB)
				{
					alert("Dont't select the same Action");
					return false;
					  
				}
				
			} 
			actionString=actionString+document.getElementById(Aname).value+",";	
		}
		
	}
	actionString=actionString.substring(0,actionString.length-1);
	 
	for(var j=1;j<=DefectCount;j++)
	{
		var Dname="Defect"+j;
		var objD=document.getElementById(Dname);
		var indexD=objD.selectedIndex;
		var selectTextD=objD.options[indexD].text;
		if(document.getElementById(Dname).value!="") 
		{
			for(p=j+1;p<=DefectCount;p++)
			{
				var Pname="Defect"+p;
				var objP=document.getElementById(Pname)
				var indexP=objP.selectedIndex;
				var selectTextP=objP.options[indexP].text;
				if(selectTextD==selectTextP)
				{
					alert("Dont't select the same Defect Code");
					return false;
					  
				}			
			} 

			defectString=defectString+document.getElementById(Dname).value+",";
			var Tname="TansactionType"+j;
			typeString=typeString+document.getElementById(Tname).value+",";
		}
	}
	defectString=defectString.substring(0,defectString.length-1);
	typeString=typeString.substring(0,typeString.length-1);
	
	
	if(isNaN(document.form1.txtCount.value)){ 
	   alert('Material Count 必须是数字！') 
	   document.form1.txtCount.focus(); 
	   return false;	   
	}
	
	
	for(var j=1;j<=CategoryCount;j++)
	{
		var Dname="CATEGORY"+j;
		var objD=document.getElementById(Dname);
		var indexD=objD.selectedIndex;
		var selectTextD=objD.options[indexD].text;
		if(document.getElementById(Dname).value!="") 
		{
			for(p=j+1;p<=CategoryCount;p++)
			{
				var Pname="CATEGORY"+p;
				var objP=document.getElementById(Pname)
				var indexP=objP.selectedIndex;
				var selectTextP=objP.options[indexP].text;
				if(selectTextD==selectTextP)
				{
					alert("Dont't select the same Category Name");
					return false;
					  
				}			
			} 
			categoryString=categoryString+document.getElementById(Dname).value+",";
		}
	}
	categoryString=categoryString.substring(0,categoryString.length-1);
	
	
	document.form1.action="/Admin/Part/SetActionDefectCode1.asp?actionString="+actionString+"&defectString="+defectString+"&typeString="+typeString+"&categoryString="+categoryString;
	document.form1.submit();
	return true;
}

function checkCount()
{
	if(isNaN(document.form1.txtCount.value)){ 
	   alert('Material Count 必须是数字！') 
	   document.form1.txtCount.focus(); 	   
	}
}
</script>
</head>

<body onLoad="reload(ActionTable.rows.length,DefectTable.rows.length-1);">
<form action="" method="post" name="form1" target="_self" onSubmit="return getSubmit();">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" id="ActionTable">
<tr>
	<td colspan="2" align="center" class="t-c-GrayLight"  height="20px" ><strong> Set Action and Defect Code </strong></td>
</tr>
<tr>
	<td class="t-c-greenCopy" height="20px" colspan="2">Set Action </td>
</tr>

<tr>
	<td colspan="2">
		<input type="button" value="Add Row" id="addAction"   onclick="insertActionRow(document.getElementById('ActionTable').rows.length)"/>
		<input name="button" type="button"  disabled="disabled" id="deleteAction"   onclick="deleteActionRow(document.getElementById('ActionTable').rows.length-1)" value="Delete Row" /> 	 
		&nbsp;</td>
</tr>
<tr>
	<td colspan="2"><Strong>Action Name</Strong></td>
</tr>
<% 
	'get current action 
SQL1="SELECT * FROM ROUTING_ACTION_DETAIL_TEMP WHERE GUID='"&strGUID&"' AND STATION_ID='"&StationId&"' ORDER BY TO_NUMBER(ACTION_SEQENCE)"
 

set rs1=server.createobject("adodb.recordset")
rs1.open SQL1,conn,1,3
if rs1.recordcount>0 then
	ActionCount=rs1.recordcount
	for i=1 to ActionCount
	 
	selId="Action"&i
%>
<tr>
	<td colspan="2">
	<select id="<%=selId %>"  style="width:350px">
		<option value="">-- Select Action --</option>
		<%= getAction_New(true,"OPTION",rs1("ACTION_ID"),""," order by Action_CODE","","") %>
	</select>	  </td>
</tr>
<%
	rs1.movenext	 
	next
else
%>
<tr>
	<td colspan="2">
		<select id="Action1"  style="width:350px">
			<option value="">-- Select Action --</option>
			<%= getAction_New(true,"OPTION","",""," order by Action_CODE","","") %>
		</select>	</td>
</tr>
<%
end if
'response.end()
%>
</table>
<br>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" id="DefectTable">
<tr>
	<td class="t-c-greenCopy" height="20px" colspan="2">Set Defect Code </td>
</tr>
<tr>
	<td colspan="2">
		<input type="button" value="Add Row" id="addDefectCode" onclick="insertDefectRow(DefectTable.rows.length)"/> &nbsp;	 
		<input type="button" value="Delete Row" id="deleteDefectCode"   disabled="disabled"  onclick="deleteDefectRow(DefectTable.rows.length-1)"  />					</td>
</tr>
<tr>
	<td algin="center"><Strong>DefectCode</Strong></td>
	<td algin="center" width="50%"><Strong>Transaction Type</Strong></td>
</tr>
<%
SQL2="SELECT * FROM ROUTING_DEFECT_DETAIL_TEMP WHERE GUID='"&strGUID&"' AND STATION_ID='"&StationId&"' ORDER BY DEFECTCODE_SEQENCE"
set rs2=server.createobject("adodb.recordset")
rs2.open SQL2,conn,1,3

if rs2.recordcount>0 then
	DefectCount=rs2.recordcount
	for i=1 to rs2.recordcount
		SelId="Defect"&i
		SelType="TansactionType"&i
%>
	<tr>
	<td  width="50%" >
		<select  id="<%=SelId%>"  style="width:350px">
			<option value="">-- Select Defect Code --</option>
			<%= getDefectCode_New("OPTION",rs2("DEFECTCODE_ID"),""," order by Defect_Code ","") %>
		</select>	
	</td>
	<td width="50%">
		<select id="<%=SelType%>"  style="width:100px">
			<!--
			<option value="0" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="0" then%> selected <% end if%>>None</option>
			
			<option value="1" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="1" then%> selected <% end if%> >Rework</option>
			-->
			<option value="2" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="2" then%> selected <% end if%>>Scrap</option>
			<!--
			<option value="3" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="3" then%> selected <% end if%>>Readjust</option>
			<option value="4" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="4" then%> selected <% end if%>>Change Model</option>
			<option value="5" <%if rs2("DEFECTCODE_TRANSACTION_TYPE")="5" then%> selected <% end if%>>Slapping</option>
			-->
		</select>	
	</td>
<%
	rs2.movenext
	'i=i+1
	next
else
%>
	<tr>
	<td  width="50%">
	<select  id="Defect1"  style="width:350px">
			<option value="">-- Select Defect Code --</option>
			<%= getDefectCode_New("OPTION","",""," order by Defect_Code ","") %>
	</select>	
	</td>
	<td width="50%">
		<select id="TansactionType1"  style="width:100px">
		<!--
			<option value="0" >None</option>
			
			<option value="1" >Rework</option>
			-->
			<option value="2" >Scrap</option>
			<!--
			<option value="3" >Readjust</option>
			<option value="4" >Change Model</option>
			<option value="5" >Slapping</option>
			-->
		</select>	
	</td>
	<%
end if
%>
</tr>
</table>
<br/>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >

<tr>
	<td class="t-c-greenCopy" height="20px" colspan="2">Set Material Count </td>
</tr>
<tr>
<%
	SQL="Select * from MATERIAL_COUNT where ROUTING_ID='"&RoutingId&"' and STATION_ID='"&StationId&"'"
	set rsM=server.createobject("adodb.recordset")
	rsM.open SQL,conn,1,3
	if rsM.recordcount>0 then
		MeterialCount=rsM("MATERIAL_COUNT")
	else
		MeterialCount=0
	end if
%>
	<td ><Strong>Material Count</Strong></td>
	<td><input name="txtCount" id="txtCount" type="text" value="<%=MeterialCount%>" onchange="checkCount()"/></td>
</tr>

</table>

<br>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" id="CategoryTable">
<tr>
	<td class="t-c-greenCopy" height="20px" colspan="2">Set Station Category </td>
</tr>
<tr>
	<td colspan="2">
		<input type="button" value="Add Row" id="addCategory" onclick="insertCategoryRow(CategoryTable.rows.length)"/> &nbsp;	 
		<input type="button" value="Delete Row" id="deleteCategory"   disabled="disabled"  onclick="deleteCategoryRow(CategoryTable.rows.length-1)"  />					</td>
</tr>
<tr>
	<td algin="center"><Strong>Category</Strong></td>
</tr>
<%
SQL2="SELECT * FROM STATION_MATERIAL_BOM WHERE ROUTING_ID='"&RoutingId&"' AND STATION_ID='"&StationId&"' ORDER BY SEQUENCE_ID"
set rs2=server.createobject("adodb.recordset")
rs2.open SQL2,conn,1,3

if rs2.recordcount>0 then
	CategoryCount=rs2.recordcount
	for i=1 to rs2.recordcount
		SelId="CATEGORY"&i
%>
	<tr>
	<td  width="50%" >
		<select  id="<%=SelId%>"  style="width:350px">
			<option value="">-- Select Category --</option>
			<%= getCategory("OPTION",rs2("CATEGORY_ID"),""," order by CATEGORY_NAME ") %>
		</select>	
	</td>
<%
	rs2.movenext
	'i=i+1
	next
else
%>
	<tr>
	<td  width="50%">
	<select  id="CATEGORY1"  style="width:350px">
			<option value="">-- Select Category --</option>
			<%= getCategory("OPTION","",""," order by CATEGORY_NAME ") %>
	</select>	
	</td>
	<%
end if
%>
</tr>
</table>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >
<tr>
	<td align="center" colspan="2">
		 
		<input type="hidden" name="StationId" value="<%=StationId%>"  />
		<input type="hidden" name="RoutingId" value="<%=RoutingId%>"  />
		<input type="hidden" name="txtGuid" value="<%=strGUID%>"  />
		<input type="hidden" name="txtActionCount" value="<%=ActionCount%>" />
		<input type="hidden" name="txtDefectCount" value="<%=DefectCount%>" />
		
		<input type="hidden" name="txtCategoryCount" value="<%=CategoryCount%>" />
		
		<input type="submit" value="Save" />&nbsp;
		 
	</td>
</tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->