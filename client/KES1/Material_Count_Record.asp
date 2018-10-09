<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
'SubJobNumber=session("JOB_NUMBER")
'job=split(SubJobNumber,"-")
'JobNumber=job(0)
RoutingId=getNewRoutineId(request("part_number_id"))
StationId=getNewStationId(request("currentStationId"))
action=request("action")

if action<>"1" then
	if RoutingId="" or StationId="" then	 
		response.Write("<script>window.alert('没有相应的Rutine设定!');window.close();</script>")
	end if
end if



if action="1" then
	counts=cint(request.Form("count"))
	RoutingId=request("part_number_id")
	StationId=request("currentStationId")
	set rsD=server.createobject("adodb.recordset")
	SQL="Delete  from MATERIAL_COUNT_RECORD where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"' and  STATION_ID='"&StationId&"' and ROUTING_ID='"&RoutingId&"' "
	rsD.open SQL,conn,1,3
	
	
	for i=1 to counts
		lableValue=request.Form("label"&i)
		AmountValue=request.Form("amount"&i)
		set rsC=server.createobject("adodb.recordset")
		if lableValue<>"" and AmountValue<>"" then
			'check Label Id
		if len(lableValue) >11  then
			  SQL="select * from EMR_PACK_DETAIL where box_id='"&lableValue&"'"
		else
			SQL="Select * from MR_Dispatch where job_number='"&session("JOB_NUMBER")&"' and (Sheet_Number like '%-" &session("SHEET_NUMBER")& "-%' or Sheet_Number like '" &session("SHEET_NUMBER")&"-%' or Sheet_Number like '%-"&session("SHEET_NUMBER")& "' or Sheet_Number = '" &session("SHEET_NUMBER")& "' OR SHEET_NUMBER IS NULL) and LABEL_NO='"&lableValue&"' and IS_Delete='0'"
		end if
			'response.Write(sql)
			rsC.open SQL,conn,1,3
			if rsC.recordcount=0 then
				CheckValue=0
				ErrorLabel=lableValue
				exit for
			else		
				CheckValue=1
			end if
		else
			CheckValue=2
		end if
	next
	
	if CheckValue=1 then
		for i=1 to counts
			lableValue=request.Form("label"&i)
			AmountValue=request.Form("amount"&i)
			if lableValue<>"" and AmountValue<>"" then
			
				set rsM=server.createobject("adodb.recordset")
				
				if len(lableValue) >11  then
				SQL="Select * from MATERIAL_COUNT_RECORD where LABELID='"&lableValue&"'"
				else
				SQL="Select * from MATERIAL_COUNT_RECORD where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"' and LABELID='"&lableValue&"' and STATION_ID='"&StationId&"'"
				end if
				rsM.open SQL,conn,1,3
				if rsM.recordcount=0 then 
					rsM.addnew
					rsM("JOB_NUMBER")=session("JOB_NUMBER")
					rsM("SUBJOBNUMBER")=session("SHEET_NUMBER")
					rsM("STATION_ID")=StationId
					rsM("ROUTING_ID")=RoutingId
					rsM("LABELID")=lableValue
					rsM("AMOUNT")=AmountValue
					rsM("SCAN_DATETIME")=now
					rsM.update
					rsM.close
					actionSuccess=1
				else
				job_number=rsM("job_number")
				sheet_number=rsM("SUBJOBNUMBER")
				job=job_number&"-"&sheet_number
				ErrorLabel=lableValue
					saveToTemp()
					actionSuccess=0
					response.Write("<script>window.alert('请不要输入重复的 "&ErrorLabel&"此标签!\r此标签号已经在"&job&"使用过！');</script>")
					
					exit for
				end if
			else
				saveToTemp()
				response.Write("<script>window.alert('All the items cannot be blank!\r请将数据填写完整！');</script>")
			end if
			
		next
			if actionSuccess=1 then				
				clearTemp()
				response.Write("<script>window.alert('Save Successfully!\r保存成功!'); dialogArguments.document.getElementById('Next').disabled=false; window.close();</script>")
			end if
		
	elseif CheckValue=2 then
		clearTemp()
		saveToTemp()
		response.Write("<script>window.alert('All the items cannot be blank!\r请将数据填写完整！');</script>")
	else
		clearTemp()
		saveToTemp()
		response.Write("<script>window.alert('"&ErrorLabel&" is not exist!\r"&ErrorLabel&"不存在！');</script>")
		 
		 
	end if
	

end if

function clearTemp()
	set rsD=server.createobject("adodb.recordset")		
	SQL="Delete from MATERIAL_COUNT_RECORD_Temp where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"'"
	rsD.open SQL,conn,1,3
end function

function saveToTemp()
	for i=1 to counts
			lableValue=request.Form("label"&i)
			AmountValue=request.Form("amount"&i)
			if lableValue<>"" and AmountValue<>"" then
				set rsT=server.createobject("adodb.recordset")			 
				SQL="Select * from MATERIAL_COUNT_RECORD_Temp where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"' and LABELID='"&lableValue&"'"
				rsT.open SQL,conn,1,3
				'if rsT.recordcount=0 then 
					rsT.addnew
				'end if
				rsT("JOB_NUMBER")=session("JOB_NUMBER")
				rsT("SUBJOBNUMBER")=session("SHEET_NUMBER")
				rsT("STATION_ID")=StationId
				rsT("ROUTING_ID")=RoutingId
				rsT("LABELID")=lableValue
				rsT("AMOUNT")=AmountValue
				rsT("SEQ")=i			
				rsT.update
				rsT.close	
			end if		 
		next
end function

function getNewRoutineId(Part_Id)
	SQL="select NID from ROUTING where nid=(select MOTHER_ROUTING_ID from part where nid='"&Part_Id&"')"	 
	set rsPart=server.createobject("adodb.recordset")
	rsPart.open SQL,conn,1,3
	if rsPart.recordcount>0 then
		getNewRoutineId=rsPart("NID")
	end if
	rsPart.close
end function

function getNewStationId(Station_Id)
	SQL="select NID from STATION_NEW where nid=(select MOTHER_STATION_ID from station where nid='"&Station_Id&"')"
	set rsSt=server.createobject("adodb.recordset")
	rsSt.open SQL,conn,1,3
	if rsSt.recordcount>0 then
		getNewStationId=rsSt("NID")
	end if
	rsSt.close
end function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Record Page</title>
<script language="javascript">
function toSave()
{
	document.form1.action="Material_Count_Record.asp?action=1";
	document.form1.submit();
}
function closePage()
{
	window.close();
}
function checkNum(txtQtyNumber)
{
	if(isNaN(txtQtyNumber.value)){ 
	   alert('Material Qty 必须是数字！') ;
	   txtQtyNumber.focus(); 
	  
	}
}

</script>
 
<link href="../CSS/GeneralKES1.css" rel="stylesheet" type="text/css" />
<base target="_self">
</head>

<body bgcolor="#99CCFF">
<form id="form1" name="form1" method="post" action="">
<table width="50%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" id="ActionTable">

<tr>
	<td class="t-c-greenCopy" height="20px" colspan="2" align="center">Material Qty Record</td>
</tr>
<tr>
	<td align="center">Label ID</td>
	<td align="center">Qty</td>
</tr>
<%
SQLA="Select * from MATERIAL_COUNT_RECORD where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"' and STATION_ID='"&StationId&"' order by LABELID"

SQLR="Select * from MATERIAL_COUNT_RECORD_TEMP where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SUBJOBNUMBER='"&session("SHEET_NUMBER")&"' and STATION_ID='"&StationId&"' order by SEQ desc"

set rsF=server.createobject("adodb.recordset")
set rsR=server.createobject("adodb.recordset")

rsF.open SQLA,conn,1,3
rsR.open SQLR,conn,1,3

SQL="Select * from MATERIAL_COUNT where ROUTING_ID='"&RoutingId&"' and STATION_ID='"&StationId&"'"		 
rs.open SQL,conn,1,3
if rs.recordcount>0 then
	counts=cint(rs("MATERIAL_COUNT"))
	for i=1 to counts
		labelId="label"&i
		amountId="amount"&i
		if not rsF.eof then
	%>
		<tr>
			<td align="center"><input name="<%=labelId%>" id="<%=labelId%>" type="text"  value="<%=rsF("LABELID")%>"/></td>
			<td align="center"><input name="<%=amountId%>" id="<%=amountId%>" type="text" value="<%=rsF("AMOUNT")%>" onblur="checkNum(this)"/></td>
		</tr>
	<%
		rsF.movenext
		elseif not rsR.eof then
	%>
	<tr>
		<td align="center"><input name="<%=labelId%>" id="<%=labelId%>" type="text"  value="<%=rsR("LABELID")%>"/></td>
		<td align="center"><input name="<%=amountId%>" id="<%=amountId%>" type="text" value="<%=rsR("AMOUNT")%>" onblur="checkNum(this)"/></td>
	</tr>
	<%	
		rsR.movenext
		else
	%>
		<tr>
			
			<td align="center"><input name="<%=labelId%>" id="<%=labelId%>" type="text" /></td>
			<td align="center"><input name="<%=amountId%>" id="<%=amountId%>" type="text" onblur="checkNum(this)"/></td>
		</tr>
	<%
		end if
	next
else
	response.Write("<script>window.alert('没有相关设定，请通知Leo Li!');window.close();</script>")
end if
rsF.close()
rsR.close()

%>
<tr>
	<td align="center" colspan="2"><input name="save" id="save" type="button" value="Save 保存"  onclick="toSave();"/>&nbsp;
	 <input name="close" id="close" type="button" value="Close 关闭" onclick="closePage();" />
	 <input type="hidden" value="<%=counts%>" id="count" name="count" />
	 <input type="hidden" value="<%=RoutingId%>" id="part_number_id" name="part_number_id" />
	 <input type="hidden" value="<%=StationId%>" id="currentStationId" name="currentStationId" />
	
	 	 </td>
</tr>

</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->