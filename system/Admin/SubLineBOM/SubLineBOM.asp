<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
action=request.QueryString("Action")
isEdit=request.QueryString("Action1")
rows=request("txtRows")

PART_NUMBER=request("txtPART_NUMBER")


if(PART_NUMBER="") then
	PART_NUMBER=request.QueryString("PART_NUMBER")
end if 

if(isEdit="edit" and action="") then
	PART_NUMBER=request.QueryString("PART_NUMBER")
	set rsRows=server.createobject("adodb.recordset")
	SQL="SELECT count(distinct SUB_PART_NUMBER) FROM ITEM_BOM WHERE PART_NUMBER='"+PART_NUMBER+"'"
 
	rsRows.open SQL,conn,1,3
 
	if(not isnull(rsRows(0))) then
		if (cint(cstr(rsRows(0)))>1) then
			rows=cint(cstr(rsRows(0)))
		else
			rows=1
		end if 
	else
		rows=1
	end if 
end if 
 
if(rows="") then
	rows=1
end if 

redim Arry(rows,4)

if(isEdit="edit" and action="") then
	PART_NUMBER=request.QueryString("PART_NUMBER")
	set rsBOM=server.createobject("adodb.recordset")
	SQL="SELECT * FROM ITEM_BOM WHERE PART_NUMBER='"+PART_NUMBER+"' ORDER BY SEQUENCE"
	rownumber=0
	rsBOM.open SQL,conn,1,3
	while not rsBOM.eof
		Arry(rownumber,0)=rsBOM("SUB_PART_NUMBER")
		Arry(rownumber,1)=rsBOM("QTY")
		Arry(rownumber,2)=rsBOM("DESCRIPTION")
		Arry(rownumber,3)=rsBOM("SEQUENCE")
		rownumber=rownumber+1
		rsBOM.movenext
	wend
end if

'Add New Row
if(action="3") then 
		for rownumber=0 to ubound(Arry)
			Arry(rownumber,0)=request("txtMaterialPartNumber"&cstr(rownumber))
			Arry(rownumber,1)=request("txtQty"&cstr(rownumber))
			Arry(rownumber,2)=request("txtDesc"&cstr(rownumber))
			Arry(rownumber,3)=cstr(rownumber)
		next 
end if 


'Save Data
if(action="2") then
	SaveTempData(rows)
end if 


function SaveTempData(rows)
	isExisting=""
	response.write isEdit
	if(isEdit="edit") then
		set rsBOMDelete=server.createobject("adodb.recordset")
		SQL="DELETE ITEM_BOM WHERE PART_NUMBER='"+PART_NUMBER+"'"
		rsBOMDelete.OPEN SQL,conn,1,3
	else
		set rsBOMExisting=server.createobject("adodb.recordset")
		SQL="SELECT * FROM ITEM_BOM WHERE PART_NUMBER='"+PART_NUMBER+"'"
		rsBOMExisting.OPEN SQL,conn,1,3
		if(rsBOMExisting.recordcount>0) then
			isExisting="Part Number:"+PART_NUMBER+" BOM exists!"
		end if 
	end if 
	
	if(isExisting="") then
		for i=0 to rows-1 
				subpartnumber=request("txtMaterialPartNumber"&i)
				qty=request("txtQty"&i)
				Desc=request("txtDesc"&i)
				
				Arry(i,0)=subpartnumber
				Arry(i,1)=qty
				Arry(i,2)=Desc
				Arry(i,3)=cstr(i)
			
				set rsBOM=server.createobject("adodb.recordset")
				SQL="INSERT INTO ITEM_BOM"
				SQL=SQL+"(PART_NUMBER,SUB_PART_NUMBER,QTY,DESCRIPTION,SEQUENCE)"
				SQL=SQL+"VALUES ('"+PART_NUMBER+"','"+cstr(subpartnumber)+"','"+cstr(qty)+"','"+cstr(Desc)+"','"+CSTR(i)+"')"
				 
				rsBOM.open SQL,conn,1,3
		next 
		response.write("<script>alert('Save Successfully!');location.href='SubLineBOMList.asp'</script>")
	else
		response.write "<script>alert('"+isExisting+"');</script>"
	end if 
	SaveTempData=""
end function 
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Admin/Station/FormCheck.js" type="text/javascript"></script>
<script>
	
	function SaveData()
	{
		if(document.getElementById("txtPART_NUMBER").value=="")
		{
			alert("Please input numeric for Part Number!");
			return;
		}
		var isEdit=document.getElementById("txtisEdit").value
		form1.action="SubLineBOM.asp?action=2&Action1="+isEdit;
		form1.submit();
	}
	
	function CreateRows()
	{
		if(isNaN(document.getElementById("txtRows").value))
		{
			alert("Please input numeric for Row Number!");
			return;
		}
		if(parseFloat(document.getElementById("txtRows").value)<=0)
		{
			alert("Row Number should be larger that 0!");
			return;
		}
		var isEdit=document.getElementById("txtisEdit").value
		
		form1.action="SubLineBOM.asp?action=3&Action1="+isEdit;
		form1.submit();
	}
</script>

</head>

<body>

<form method="post" name="form1" target="_self">
  <table width="100%"  border="2" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
	<%if isEdit<>"edit" then%>
      <td height="20" colspan="6" class="t-c-greenCopy">Add a New BOM</td>
	  <%else%>
	  <td height="20" colspan="6" class="t-c-greenCopy">Edit a BOM</td>
	 <%end if %>
    </tr>
 	 <tr>
	   <td height="20">Rows</td>
	   <td height="20">
	   		<input name="txtRows" type="text" id="txtRows" size="30" value="<%=rows%>">
	   </td>
	   <td height="20" colspan="4"><input type="button" name="btnCreateRows" value="Create Rows" width="75px" onclick="CreateRows()">
	   	<input name="txtRowsHidden" type="hidden" id="txtRowsHidden" size="30" value="<%=rows%>">
	   </td>
    </tr>
	
    <tr>
      <td height="20" width="15%">Part Number <span class="red">*</span> </td>
      <td height="20" class="red" colspan="5">
        <div align="left">
          <input name="txtPART_NUMBER" type="text" id="txtPART_NUMBER" size="30" value="<%=PART_NUMBER%>" <%if isEdit="edit" then response.write "readonly" end if %>>
      </div></td>
    </tr>
    
	
	
	 <%for mm=0 to  ubound(Arry)-1%>
	
   <tr>
    <td height="20" width="10%">Material Part Number</td>
     <td height="20" width="10%">
	 	   <input name="txtMaterialPartNumber<%=mm%>" type="text" id="txtMaterialPartNumber<%=mm%>" size="30" value="<%=Arry(mm,0)%>" />
	</td>
    <td width="10%">Qty</td>
     <td height="20" width="10%"> 
    	<input name="txtQty<%=mm%>" type="text" id="txtQty<%=mm%>" size="30" value="<%=Arry(mm,1)%>" />
	 </td>
      <td width="10%">Description</td>
     <td height="20" width="60%"> 
    	<input name="txtDesc<%=mm%>" type="text" id="txtDesc<%=mm%>" size="100" value="<%=Arry(mm,2)%>" />
	 </td>
	</tr>
		<%
			next %>
  	 

    <tr>
      <td height="20" colspan="6"><div align="center">
          <input type="button" name="button" value="Save" width="75px" onclick="SaveData()">
		  <input name="txtisEdit" type="hidden" id="txtisEdit" size="30" value="<%=isEdit%>" />
	        <input type="reset" name="Submit7" value="Return" width="75px" onclick="location.href='SubLineBOMList.asp'">
		 </td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->