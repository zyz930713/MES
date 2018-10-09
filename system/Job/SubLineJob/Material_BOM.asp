<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Trans.asp" -->

<%	
	dim arry 
	Material_Part_Number=request.QueryString("Material_Part_Number")
	JobNumber=request.QueryString("JobNumber")
	qty=request.QueryString("qty")
	if(request.QueryString("Action")="") then
		if(jobnumber="0")then
			sqlstr="select PART_NUMBER,SUB_PART_NUMBER,QTY,DESCRIPTION from ITEM_BOM where PART_NUMBER='"+Material_Part_Number+"' order by SEQUENCE"
			rs.open sqlstr,conn,1,3
			arry= rs.GetRows
		else
			sqlstr="select PART_NUMBER,ITEM_ANME,REQUIRE_QUANTITY,DESCRIPTION from job_bom where job_number='"+JobNumber+"' order by ITEM_ANME"
			rs.open sqlstr,connTicket,1,3
			arry= rs.GetRows
		end if 
	end if 

	if(request.QueryString("Action")="2") then
		RowNumber=request.QueryString("RowNumber")
		set rsTBL_LOT_BOM_ITEM_Delete=server.createobject("adodb.recordset")
		SQLStr= "delete TBL_LOT_BOM_ITEM where INVENTORY_ITEM_ID='"+jobnumber+"' and ORGANIZATION_ID='24' and WIP_ENTITY_ID='"+jobnumber+"'"
		
		rsTBL_LOT_BOM_ITEM_Delete.open SQLStr,connTicket,1,3
			
		for i=0 to RowNumber			
			SubPartNumber=request("txtMaterialPartNumber"&i)
			Qty=request("txtQty"&i)
			Dec=request("txtDec"&i)
			set rsTBL_LOT_BOM_ITEM=server.createobject("adodb.recordset")
					SQL="INSERT INTO TBL_LOT_BOM_ITEM (INVENTORY_ITEM_ID, ORGANIZATION_ID, INVENTORY_ITEM_NAME, DESCRIPTION, WIP_ENTITY_ID, REQUIRE_QUANTITY, LAST_UPDATE_DATE,operation_seq_num)"
					SQL=SQL+"VALUES('"+JobNumber+"','24','"+SubPartNumber+"','"+Dec+"','"+jobnumber+"','"+cstr(Qty)+"',sysdate,"+cstr(i)+")"
			rsTBL_LOT_BOM_ITEM.open SQL,connTicket,1,3		
		next 
		set rs2=server.createobject("adodb.recordset")
		sqlstr="select PART_NUMBER,ITEM_ANME,REQUIRE_QUANTITY,DESCRIPTION from job_bom where job_number='"+JobNumber+"' order by ITEM_ANME"
		rs2.open sqlstr,connTicket,1,3
		arry= rs2.GetRows
		
		response.write "<script>alert('修改成功!')</script>"
	end if 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<base target="_self">
<script>
	function SaveData()
	{
		var count=document.getElementById("txtcount").value;
		for(var i=0; i<=count-1;i++)
		{
		
			if (document.getElementById("txtMaterialPartNumber"+i).value=="")
			{
				alert("请填写"+(i+1)+"行的Sub Part Number!");
				return;
			}
			
			if (isNaN(document.getElementById("txtQty"+i).value))
			{
				alert("请选择"+(i+1)+"行的Job Qty应该为数字!");
				return;
			}

		}
		JobNumber=document.getElementById("txtjobnumber").value
		Material_Part_Number=document.getElementById("txtMaterial_Part_Number").value

		count=count-1
		form1.action="Material_BOM.asp?Action=2&RowNumber="+count+"&JobNumber="+JobNumber+"&Material_Part_Number="+Material_Part_Number;
		form1.submit();
	}
</script>
<body>
<form  name="form1" id="form1" method="post">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
  	 <td height="20" class="t-t-Borrow">Part Number</td>
	  <td height="20"class="t-t-Borrow">Sub Part Number</td>
	 <td height="20"class="t-t-Borrow">Qty</td>
	  <td height="20" class="t-t-Borrow">Description</td>
  </tr>
  <%if(jobnumber="0")then%>
		  <%For row = 0 To UBound(arry,2)%>
		  <tr>
			 <td height="20"><%=arry(0, row)%></td>
			  <td height="20"><%=arry(1, row)%></td>
			 <td height="20"><%=arry(2,row)%></td>
			  <td height="20"><%=arry(3, row)%></td>
		  </tr>		 
		 <%Next%>
	<%else%>
		<%For row = 0 To UBound(arry,2)%>
		 <tr>
		 <td height="20"><%=arry(0, row)%></td>
		  <td height="20"><input type="text" id="txtMaterialPartNumber<%=row%>" name="txtMaterialPartNumber<%=row%>" value="<%=arry(1, row)%>"></td>
		 <td height="20"><input type="text" id="txtQty<%=row%>" name="txtQty<%=row%>" value="<%=arry(2, row)%>"></td>
		  <td height="20"><input type="text" id="txtDec<%=row%>" name="txtDec<%=row%>" value="<%=arry(3, row)%>"></td>
		  </tr>	
		<%Next%>
		<tr>
  			 <td height="20" colspan="4" align="center"><input type="button" id="btnSave" name="btnSave" value="Update Job BOM" onclick="SaveData()" class="t-b-Yellow">
			 <input type="hidden" id="txtcount" name="txtcount" value="<%=row%>"> 
			  <input type="hidden" id="txtjobnumber" name="txtjobnumber" value="<%=JobNumber%>"> 
			   <input type="hidden" id="txtMaterial_Part_Number" name="txtMaterial_Part_Number" value="<%=Material_Part_Number%>"> 
			 </td>
 		 </tr>
	<%end if %>

</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->