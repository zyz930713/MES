<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
'SubJobNumber=session("JOB_NUMBER")
'job=split(SubJobNumber,"-")
'JobNumber=job(0)

action=request("action")




if action="1" then
	
	
	
end if





%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Record Page</title>
<script language="javascript">
function addDefectQty(){
	if(event.keyCode==13){//Key:Enter
		
		toSave();		
	}
}
function toSave()
{
	
	var Box_ID=document.getElementById('Box_ID');
	var PART_NUMBER=document.getElementById('txt12NC');
	var Job=document.getElementById('Job');
	var Qty=document.getElementById('Qty');
	if(Box_ID.value.length!=16){
			alert("Box ID Error\扫描的签标位数不对，请确认！");
			document.all.Box_ID.select();
			return false;
	}
	else
	{
	
	var Box_ID=Box_ID.value;
	var PART_NUMBER=PART_NUMBER.value;
	var Job=Job.value;
	var Qty=Qty.value;
	

		var deftInfo = window.showModalDialog("Material_HV_Record2.asp?Box_ID="+Box_ID+"&PART_NUMBER="+PART_NUMBER+"&Job="+Job+"&Qty="+Qty,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");

		if(deftInfo!=""){	
			alert(deftInfo);
			document.all.Box_ID.value="";
			
			document.all.Box_ID.select();
			return false;
		}else{
			document.all.labelId.value="";
			document.all.LotNo.value="";
			document.all.labelId.select();
		}
		
		
		
		
	//document.form1.action="Material_HV_Record.asp?action=1";
	//document.form1.submit();
	}
}
function closePage()
{
	window.close();
}


</script>
 
<link href="../CSS/GeneralKES1.css" rel="stylesheet" type="text/css" />
<base target="_self">
</head>

<body bgcolor="#99CCFF" onload="javascript:document.all.Box_ID.focus();">
<form id="form1" name="form1" method="post" action="">
<table width="50%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" id="ActionTable">

<tr>
	<td class="t-c-greenCopy" height="20" colspan="3" align="center">Box ID  WarehouseRec</td>
</tr>
<tr>
	<td width="49%" align="center">BOX ID</td>
	<td width="37%" align="center"><input name="Box_ID" type="text" id="Box_ID" onkeydown="if(this.value&&event.keyCode==13){event.keyCode=9;}"  value="" size="30"/></td>
	<td rowspan="4" align="center">&nbsp;</td>
  </tr>
<tr>
  <td align="center">12NC</td>
  <td width="37%" align="center"><input name="txt12NC" type="text" id="txt12NC" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}"  value="" size="30"/></td>
</tr>
<tr>
  <td align="center">JOB</td>
  <td width="37%" align="center"><input name="Job" type="text" id="Job" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}"  value="" size="30"/></td>
</tr>
		<tr>
			<td align="center">Qty</td>
			<td align="center"><input name="Qty" type="text" id="Qty"   onkeydown="addDefectQty()" value="" size="30"/></td>
        </tr>

<tr>
	<td align="center" colspan="3"><input name="save" id="save" type="button" value="Save 保存"  onclick="toSave();"/>
	  &nbsp;
	 <input name="close" id="close" type="button" value="Close 关闭" onclick="closePage();" />
	
	
	 	 </td>
</tr>

</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->