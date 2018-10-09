<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/ReTestAndQA.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
ComputerName=request("computername")
 response.Write(ComputerName)

OQCName=GetOQCName(ComputerName)
    response.Write(OQCName)

if OQCName<>"OQC" or isnull(OQCName) then

	'response.Redirect("/OQC/OQCinfo.asp")
	
	word="<span align='center' style='color:red;'>此电脑不是OQC专用电脑，不能做OQC！</span>"
	response.Redirect("OQCinfo.asp?word="&word)
end if 


function GetOQCName(ComputerName)
		set rsPrinterName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(ComputerName)+"'"
	 
		rsPrinterName.open SQL,conn,1,3
	 	if(rsPrinterName.recordcount=0) then
			GetOQCName=""
			exit function
		end if 
		
		if (rsPrinterName(1)="") then
			GetOQCName=""
			exit function
		end if 
		
		GetOQCName=rsPrinterName(1)
		rsPrinterName.close
		set rsPrinterName=nothing
	end function

op=request("operatorcode")
action=request("Action")
BatchNo=request("txtBatchNo")
Model=request("txtModelName")
StartQty=request("txtStartQty")
AQL=request("txtAQL")
SampleQty=request("txtSampleQty")
isSave="0"
errMsg=""

'get main job number
if BatchNo<>"" then
	arrJob=split(BatchNo,"-")
	if ubound(arrJob) >2 then
	if arrJob(1)="E" or arrJob(1)="R" then
		jobNumber=arrJob(0)&"-"&arrJob(1)			
	else
		jobNumber=arrJob(0)
	end if
	end if	
end if

if action="1" then	
	SQL="SELECT * FROM Job_OQC WHERE BatchNoList = '"&BatchNo&"'"
	rs.open SQL,conn,1,3
	if rs.recordcount>0 then
		if rs("OQCEndTime") <> "" then
			errMsg="This batch no has done OQC Check.此批号已经做过OQC检测。"
		else
			OQCNID = rs("OQCNID") 
			Model = rs("ModelName")
			StartQty = rs("StartQty")
			AQL = rs("AQL")
			SampleQty = rs("SampleQty")
			isSave = "1"
		end if
	end if
	rs.close		
	
	if errMsg="" then
		SQL="select actualqty,subjoblist,transactiontype from label_print_history where batchno='"&BatchNo&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then			
			if rs("transactiontype")<>"-1" then
				errMsg="This batch no is not good product, does not neet to do OQC Check. \n该批号为不良品，不需要做OQC检测。"
			else
				StartQty = rs("actualqty")
				sheetNumber= rs("subjoblist")
				'subJob=jobNumber&"-"&string(3-len(sheetNumber),"0")&sheetNumber
				sheetNumberS=split(sheetNumber,"-")
			end if
		else
			errMsg="This batch no is not exits. \n该批号不存在。"
		end if
		rs.close
		if errMsg="" then
			SQL="select part_number_id,part_number_tag,finished_stations_id from job where job_number = '"&jobNumber&"' and sheet_number = '"&sheetNumberS(0)&"'"	
			rs.open SQL,conn,1,3
			if not rs.eof then
				Model = rs("part_number_tag")
				session("DEFECT_STATIONS")=rs("finished_stations_id")
				session("PART_NUMBER_ID")=rs("part_number_id")
			end if
			rs.close
			if isSave="0" then
				AQL = GetCurrentAQL(Model)
				SampleQty = QA_SampleQty(StartQty,AQL)
			end if
			
		end if
	end if
'Save Data	
elseif action="2" then
	OQCID="QC"&NID_SEQ("OQC_SEQ")	
	SQL="SELECT * FROM Job_OQC WHERE 1=2"
	rs.open SQL,conn,1,3
	rs.addnew
		rs("OQCNID")=OQCID
		rs("BatchNoList")=BatchNo
		rs("ModelName")=Model
		rs("StartQty")=StartQty
		rs("AQL")=AQL
		rs("SampleQty")=SampleQty
		rs("OQCStartTime")=now()
		rs("StartOperator")=op
	rs.update
	rs.close
	response.write "<script>window.alert('Saving data successfully!成功保存数据!');location.href='OQC.asp?operatorcode="&op&"&ComputerName="&ComputerName&"';</script>"

'OQC judge
elseif action="3" then
	successMsg="Saving data successfully!成功保存数据!"
	href="OQC.asp?operatorcode="&op&"&ComputerName="&ComputerName
	
	
	OQCID=request("OQCNID")
	transTime = cstr(now())
	SQL="update Job_OQC set OQCEndTime='"+transTime+"',AQL='"+AQL+"',SampleQty='"+SampleQty+"' WHERE OQCNID='"+OQCID+"'"
	rs.open SQL,conn,1,3
	
	SQL="SELECT * FROM JOB_OQC_DETAIL WHERE 1=2"
	rs.open SQL,conn,1,3		
	result=request("oqcResult")
	if result = "Pass" then
		TotalFailUnit=0
		IsPassOQC="PASS"
		rs.addnew
		rs("OQCDETAILNID")="QD"&NID_SEQ("OQC_DETAIL_SEQ")
		rs("OQCNID")=OQCID
		rs("BATCHNO")=BatchNo
		rs("SHOWSEQUENCE")="0"
		rs("PRODUCTIONRETESTREJECTQTY")=0
		rs("Qty")=StartQty
		rs("LineLossQty")=0
		rs("OQCReject")=0
		rs("IsPassOQC")=1
		rs("OQCEndTime")=transTime
		rs.update
	else
		IsPassOQC="FAIL"
		deftCount=cint(request("defect_count"))
		for i=0 to deftCount-1
			rs.addnew
			rs("OQCDETAILNID")="QD"&NID_SEQ("OQC_DETAIL_SEQ")
			rs("OQCNID")=OQCID
			rs("BATCHNO")=BatchNo
			rs("SHOWSEQUENCE")=i
			rs("PRODUCTIONRETESTREJECTQTY")=0
			rs("Qty")=StartQty
			rs("LineLossQty")=0
			rs("OQCReject")=request("defect_quantity"&i)
			rs("DEFECTDESC")=request("defect_code"&i)&":"&request("defect_name"&i)
			rs("REJECTTRANSACTIONTYPE")=request("station_id"&i)
			rs("IsPassOQC")=0
			rs("OQCEndTime")=transTime
			rs.update
		next		
	end if	
	rs.close
		
	NewAQL=SetCurrentAQL(Model,IsPassOQC,AQL,BatchNo,TotalFailUnit)
	
	'create 500 job number and print
	if IsPassOQC="FAIL" then
		sheetNo=500
		SQL="select * from rework_printing where job_number = '"&jobNumber&"' order by seq desc "
		rs.open SQL,conn,1,3
		if not rs.eof then
			sheetNo=cint(rs("seq"))+1
		end if
		rs.addnew
			rs("job_number")=jobNumber
			rs("qty")=StartQty
			rs("batchno")=BatchNo
			rs("station_name")=request("station_id0")
			rs("seq")=sheetNo
			rs("model")=Model
		rs.update
		rs.close
		
		'update 2D code info
		SQL="select subjoblist from label_print_history where batchno='"&BatchNo&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
			subJobList=rs("subjoblist")
		end if
		SQL="update job_2d_code set sheet_number='"&sheetNo&"',lm_user='"&op&"',lm_time=sysdate where defect_code_id is null and job_number='"&jobNumber&"' and sheet_number in ('"&replace(subJobList,"-","','")&"')"
		conn.execute(SQL)
		
		successMsg="Saving data successfully, please print retest label! 成功保存数据, 请打印retest工单。"
		href=href&"&reject_batchno="&BatchNo
	end if
	
	response.write "<script>window.alert('"&successMsg&"');location.href='"&href&"';</script>"
end if
	
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>OQC Check OQC检测</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script type="text/javascript">
	function submitForm(action,result)
	{				
		if(form1.txtBatchNo.value.indexOf(",")>-1){
			alert("You can only input a BatchNo at one time. \n一次只能输入一个批号。");
			form1.txtBatchNo.select();
			return false;
		}
		if(action != "1"){
			if(!form1.operatorcode.value){
				alert("Op Code cannot be blank. \n工号不能为空。")
				form1.operatorcode.focus();
				return false;
			}else if(!form1.txtBatchNo.value){
				alert("Batch No cannot be blank. \n批号不能为空。")
				form1.txtBatchNo.focus();
				return false;			
			}
			if(!form1.txtStartQty.value){
				alert("Total Qty cannot be blank. \n总数量不能为空。")
				return false;
			}
			if(!form1.txtSampleQty.value){
				alert("Sample Qty cannot be blank. \n抽样数不能为空。")
				return false;
			}
		}
		//when OQC judge,check defect
		if(action == "3" && !submitDefect(result)){
			return false;
		}
		
		form1.Action.value=action;
		form1.oqcResult.value=result;
		form1.submit();
	}
	function changeAQL(){
		var selAql = document.all.selAQL.value.split("-");
		document.all.txtAQL.value=selAql[0];
		document.all.txtSampleQty.value=selAql[1];
	}
	
	<%if errMsg <> "" then%>
		alert("<%=errMsg%>");
		window.location.href="OQC.asp?operatorcode=<%=op%>&txtBatchNo=<%=BatchNo%>";
	<%end if%>
</script>
</head>
<body bgcolor="#339966" onload="form1.txtBatchNo.focus();">
<form method="post" name="form1" action="OQC.asp" >
<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
	<tr class="t-t-Borrow" >
		<td colspan="8" class="t-t-DarkBlue">Scan OQC Check 扫描OQC</td>
    </tr>
	
	<tr>
		<td nowrap>Op Code 工号 <span class="red">*</span> </td>
		<td nowrap colspan="7">
			<input name="operatorcode" type="operatorcode" id="code" value="<%=op%>" size="20"  autocomplete="off">
	  		<br><span id="errorinsertcode" class="strongred"></span>
		</td>
	</tr>
	<tr>
		<td nowrap >BatchNo 批号 <span class="red">*</span></td>
		<td colspan="7">
			<input  type="text" name="txtBatchNo" id="txtBatchNo" value="<%=BatchNo%>" size="40" onchange="if(this.value){submitForm('1')}"
				onKeyDown="if(this.value&&event.keyCode==13){event.keyCode=9;}" >
		</td>
	</tr>
	<tr>
		<td nowrap >Reject BatchNo 拒绝批号</td>
		<td colspan="7">
			<input  type="text" name="rejBatch" id="rejBatch" value="<%=request("reject_batchno")%>" size="40"
				onKeyDown="if(this.value&&event.keyCode==13){event.keyCode=9;}" >
		</td>
	</tr>	
	<tr>
		<td nowrap >Part No 料号</td>
		<td ><input  type="text" name="txtModelName" id="txtModelName" value="<%=Model%>" readonly /></td>		
		<td nowrap>AQL</td>
		<td ><select id="selAQL" name="selAQL" onchange="changeAQL()" style="width:70px">
			<%=GetSelectAQL(StartQty,AQL)%>
			</select>
			<input  type="hidden" name="txtAQL" id="txtAQL" size="4" value="<%=AQL%>"readonly/>
		</td>
		<td nowrap>Total Qty 总数量</td>
		<td >
			<input  type="text" name="txtStartQty" size="7" id="txtStartQty" value="<%=StartQty%>" readonly/>
		</td>
		<td nowrap>Sample Qty 抽样数</td>
		<td ><input  type="text" name="txtSampleQty" id="txtSampleQty" size="7" value="<%=SampleQty%>" readonly/>		
	</tr>
<%if isSave = "1" then%>
	<tr><td colspan="8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td height="" align="center" valign="top">
			<table width="100%" height="74" border="1" cellPadding=0 cellSpacing=0>
				<tr>
				  <td width="43%"><span id="inner_JobsList">Master Jobs </span>主工单</td>
				  <td width="57%">Sub Job Number 子工单号</td>
				</tr>				
				<tr align="center" height="35">
					<td height="23"><%=jobNumber%> &nbsp;</td>	
			      <td><%=sheetNumber%></td>
				</tr>
			</table>	
		</td>
		<td width="60%" valign="top">
		<table id="tb_defect_code" width="100%" border=1 cellSpacing=0 cellPadding=0 >
			<tr><td colspan=2 align=center>Defect Code 缺陷代码</td><td align=center>Quantity 数量</td></tr>
			<tr><td height="49" colspan=2 ><input type=text id="defect_code" name="defect_code" onChange="removeDefect()"  onKeyDown="if(this.value&&event.keyCode==13){event.keyCode=9;}"></td>
				<td ><input type=text size=3 id="defect_qty" name="defect_qty" onKeyDown="addDefectQty()" >&nbsp;
				<input style="height:30px; width:65px; font-size:small; font-weight:normal" type=button value='Add 添加' onClick="addDefect()">
				</td>
			</tr>
		</table>	
		</td>
	</tr>	
	</table></td></tr>
<%end if%>	
</table>
<input type="hidden" name="Action" id="Action">
<input type="hidden" name="oqcResult" id="oqcResult">
     <input type="hidden" ID="computername" name="computername" value="<%=ComputerName %>" />
<br>
<center>
<%if isSave = "1" then%>
<input type="hidden" id="OQCNID" name="OQCNID" value="<%=OQCNID%>" />
    
<input name="btnPass" type="button"  id="btnPass" value="Pass 通过" onclick="submitForm('3','Pass')" />&nbsp;
<input name="btnReject" type="button"  id="btnReject" value="Reject 拒绝" onclick="submitForm('3','Reject')" />
<%else%>
 
<input name="btnSave" type="button"  id="btnSave" value="Save 保存" onclick="submitForm('2')" />&nbsp;
<input name="btnPrint" type="button"  id="btnPrint" value="Print 打印"  onclick="printRetestTicket()" />
<%end if%>
<input name="btnClose" type="button"  id="btnClose" value="Close 关闭"  onclick="window.close();" />
</center>
</form>
</body>
<script type="text/javascript">

function printRetestTicket(){
	var batchNo=document.all.rejBatch.value;
	if(batchNo){
		window.showModalDialog('PrintRetestTicket.asp?batch_no='+batchNo,'','dialogHeight:800px;dialogWidth:800px;status:no');
	}else{
		alert("Reject Batch No cannot be blank. \n拒绝批号不能为空。")
		form1.rejBatch.focus();
		return false;
	}
}

function removeDefect(){
	if(document.all.defect_code.value){
    	var objTable = document.getElementById("tb_defect_code");
    	for(var i=2;i<objTable.rows.length;i++){
			if(document.all.defect_code.value == objTable.rows(i).cells(0).innerText){
				objTable.deleteRow(i);
				document.all.defect_code.value="";
				document.all.defect_code.blur();
				document.all.defect_code.focus();				
				return;
			}
    	}
   	}	
}

function addDefectQty(){
	if(event.keyCode==13){//Key:Enter
		event.keyCode=8;//Key:Backspace
		addDefect();		
	}
}

function addDefect(){
	if(!document.all.defect_code.value){
		alert("Defect Code can not be blank!\n缺陷代码不得为空！");
		document.all.defect_code.focus();
		return false;
	}else if(!document.all.defect_qty.value){
		alert("Quantity can not be blank!\n数量不得为空！");
		document.all.defect_qty.focus();
		return false;	
	}else if(isNaN(document.all.defect_qty.value)){
		alert("Input Quantity is not a number!\n输入的数量不是数字！");
		document.all.defect_qty.select();
		return false;		
	}else{
		var deftInfo = window.showModalDialog("../KES1/GetValueByKey.asp?key=DefectCode&keyValue="+document.all.defect_code.value,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		if(deftInfo.indexOf("Error-")==0){	
			alert(deftInfo.substring(6));
			document.all.defect_qty.value="";
			document.all.defect_code.select();
			return false;
		}else{
			var aryDeftInfo = deftInfo.split("$");
			var objTable = document.getElementById("tb_defect_code");
			var newRow = objTable.insertRow(objTable.rows.length); 
			var newCell0 = newRow.insertCell(newRow.cells.length);
			var newCell1 = newRow.insertCell(newRow.cells.length);
			var newCell2 = newRow.insertCell(newRow.cells.length);
			var newCell3 = newRow.insertCell(newRow.cells.length);
			var newCell4 = newRow.insertCell(newRow.cells.length);
			newCell0.innerHTML = document.all.defect_code.value;
			newCell1.innerHTML = aryDeftInfo[1];//defect name
			newCell2.innerHTML = document.all.defect_qty.value;
			newCell3.innerHTML = aryDeftInfo[2];//station id
			newCell3.style.display ="none";
			newCell4.innerHTML = aryDeftInfo[0];//defect code id
			newCell4.style.display ="none";
			document.all.defect_code.value = "";
			document.all.defect_qty.value = "";
			document.all.defect_code.focus();
		}
	}	
}

function submitDefect(result){
	var objTable = document.getElementById("tb_defect_code");
	var deftInfoStr="";
	var deftCount=0;
	var totalDefectQty=0;
	for(var i=2;i<objTable.rows.length;i++){    
		deftInfoStr=deftInfoStr+"&defect_code"+deftCount+"="+objTable.rows(i).cells(0).innerText;
		deftInfoStr=deftInfoStr+"&defect_name"+deftCount+"="+objTable.rows(i).cells(1).innerText;
		deftInfoStr=deftInfoStr+"&defect_quantity"+deftCount+"="+objTable.rows(i).cells(2).innerText;
		deftInfoStr=deftInfoStr+"&station_id"+deftCount+"="+objTable.rows(i).cells(3).innerText;
		deftCount++;
		totalDefectQty = totalDefectQty + new Number(objTable.rows(i).cells(2).innerText);
	}
	if(result=="Pass" && deftCount>0){
		alert("Can not Pass OQC when input defect information. \n有缺陷不能通过OQC。");
		return false;
	}else if(result=="Reject" && deftCount==0){
		alert("Please input defect information when OQC Reject. \n请输入OQC拒绝的缺陷信息。");
		return false;
	}else if(totalDefectQty>document.form1.txtStartQty.value){
		alert("Total Defect Code quantity exceeds Total Qty.\n缺陷数量合计超过总数量。");
		return false;
	}
	
	deftInfoStr = "defect_count="+deftCount+deftInfoStr;
	document.form1.action=document.form1.action+"?"+deftInfoStr;
	return true;
}
</script>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
