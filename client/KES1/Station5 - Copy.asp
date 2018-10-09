<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/GetStationName.asp" -->
<!--#include virtual="/Functions/GetPartFactory.asp" -->
<!--#include virtual="/Functions/getValidActionValue.asp" -->
<%
session.Contents.Remove("2DCodeInfo")
pagename="Station5.asp"
has2DCode=false
station_close_time=now()
'get job's properties.
SQL="select J.PART_NUMBER_ID,J.START_TIME,J.CURRENT_STATION_ID,S.STATION_NAME,S.STATION_CHINESE_NAME,OPENED_STATIONS_ID from JOB J inner join STATION S on J.CURRENT_STATION_ID=S.NID where J.JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	part_number_id=rs("PART_NUMBER_ID")
	station_name=rs("STATION_NAME")
	station_chinese_name=rs("STATION_CHINESE_NAME")
	job_start_time=rs("START_TIME")
	'current_station_id=rs("CURRENT_STATION_ID")
	current_station_id=session("CURRENT_STATION_ID")
	session("DEFECT_STATIONS")=left(rs("OPENED_STATIONS_ID"),instr(rs("OPENED_STATIONS_ID"),current_station_id)+len(current_station_id))
end if
rs.close

factory=getPartFactory(part_number_id)
if session("STATIONS_ROUTINE")="1" then
	current_station_id=session("CURRENT_STATION_ID")
end if

'get stations' properties.
SQL="select START_TIME,STATION_START_QUANTITY from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	station_start_time=rs("START_TIME")
	station_start_quantity=rs("STATION_START_QUANTITY")
end if
rs.close

%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JValueCheck.js" type="text/javascript"></script>
</head>

<body onKeyPress="keyhandler()"  bgcolor="#339966" onLoad="pageload()">
<span id="erroralarm"></span>
<form action="/KES1/Station6.asp" method="post" name="form1" target="_self" onsubmit="return formSubmit()"  >
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-t-DarkBlue">Select defect code and input quantity 选择缺陷代码并输入数量</td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Operator 操作员:
      <% =session("operator") %>
(
<% =session("code") %>
)</td>
  </tr>
  <tr>
    <td height="20" colspan="2"><div align="center"><span id="errorinsert" class="strongred"></span>&nbsp;</div></td>
  </tr>
  <tr>
    <td  height="20">Job Number 工单号</td>
    <td  height="20"><% =session("JOB_NUMBER") %></td>
    </tr>
  <tr>
    <td height="20">Sheet Number 分批号 </td>
    <td height="20"><% =session("SHEET_NUMBER") %></td>
  </tr>
  <tr>
    <td height="20">Part Number 型号 </td>
    <td height="20"><% =session("PART_NUMBER_TAG") %></td>
    </tr>
  <tr>
    <td height="20">Line Name 线别</td>
    <td height="20"><% =session("LINE_NAME") %>
	<input type="hidden" name="LINENAME"  id="LINENAME" value="<% =session("LINE_NAME")%>">
	
	<%
	
	 SQL="select * from Line where Line_name='"&session("LINE_NAME")&"'"
  rs.open SQL,conn,1,3
  if not rs.eof then
	'response.Write(rs("CODE_LINENAME"))
 %>
      
      <input type="hidden" name="CODE_LINENAME"  id="CODE_LINENAME" value="<% =rs("CODE_LINENAME")%>">
      <input type="hidden" name="CODE_NAME"  id="CODE_NAME" value="<% =rs("CODE_NAME2")%>">
      <input type="hidden" name="FACTORY_CODE"  id="FACTORY_CODE" value="<% =rs("FACTORY_CODE")%>">
      <input type="hidden" name="VERSION_NUMBER"  id="VERSION_NUMBER" value="<% =rs("VERSION_NUMBER")%>">
     
     
     <%  end if
  rs.close%>
     
      </td>
  </tr>
  <tr>
    <td height="20">Job Start Time 工单开始时间</td>
    <td height="20"><% =job_start_time %></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Station Defect 本站缺陷</td>
    </tr>
  <tr>
    <td height="20">Station Name 本站站名</td>
    <td height="20"><span class="red"><% =session("CURRENT_STATION_NAME")%>&nbsp;<% =session("REPEATED_SEQUENCE") %></span></td>
  </tr>
  <tr>
    <td height="20"> Start Time 开始时间</td>
    <td height="20"><% =station_start_time%></td>
    </tr>
  <tr>
    <td height="20">Close Time 结束时间</td>
    <td height="20"><% =station_close_time%>&nbsp;</td>
    </tr>
  <tr>
    <td height="20">Passed Time 用时</td>
    <td height="20"><% =datediff("n",station_start_time,station_close_time)%> minutes&nbsp;</td>
  </tr>
  <tr>
    <td height="20">Job Quantity 工单数量</td>
    <td height="20"><%=station_start_quantity%>&nbsp;</td>
  </tr>
  <%if trim(session("JOB_TYPE"))="N" or trim(session("JOB_TYPE"))="S" or trim(session("JOB_TYPE"))="R" or trim(session("JOB_TYPE"))="C" then%>
  <tr>
    <td height="20">Defect Code 缺陷代码</td>
    <td height="20">
		<table id="tb_defect_code" border=1 cellSpacing=0 cellPadding=0 >
			<tr><td colspan=2 align=center>Defect Code 缺陷代码</td><td align=center>Quantity 数量</td></tr>
			<tr><td colspan=2 ><input type="hidden" name="recent_defect_id" id="recent_defect_id">            	
				<input type=text id="defect_code" name="defect_code" onChange="removeDefect()"  onKeyDown="if(this.value&&event.keyCode==13){event.keyCode=9;}"></td>
				<td ><input type=text size=3 id="defect_qty" name="defect_qty" onKeyDown="addDefectQty()" >&nbsp;
				<input style="height:30px; width:65px; font-size:small; font-weight:normal" type=button value='Add 添加' onClick="addDefect()">
				</td>
			</tr>
		</table>	
	</td>
  </tr>
 <%
  end if
  'get action list
  SQL="select TRANSACTION_TYPE,ACTIONS_INDEX from STATION where NID='"&current_station_id&"'"
  rs.open SQL,conn,1,3
  if not rs.eof then
	  ACTIONS_INDEX=rs("ACTIONS_INDEX")
	  if not isnull(ACTIONS_INDEX) then
	  	aaction=split(ACTIONS_INDEX,",")
	  end if
	  station_default_transaction_type=rs("TRANSACTION_TYPE")
  else
  	  station_default_transaction_type="0"
  end if
  rs.close
  
  if not isnull(ACTIONS_INDEX) then
  SQL="select NID,ACTION_PURPOSE,STATION_POSITION,ACTION_NAME,ACTION_CHINESE_NAME,NULL_ALLOW,ELEMENT_TYPE,ELEMENT_NUMBER,ACTION_TYPE,WITH_LOT from ACTION where STATION_POSITION=1 and STATUS=1 and NID in ('"&replace(ACTIONS_INDEX,",","','")&"')"
  SQL=SQL+" order by instr('"&ACTIONS_INDEX&"',NID) "
  
  rs.open SQL,conn,1,3
  i=1
  while not rs.eof 
	max_quantity=0
	min_quantity=0	
  %>
    <tr>
      	<td height="20"><span class="<%if rs("NULL_ALLOW")<>"0" then%>StrongGreen<%end if%>"><%=rs("ACTION_NAME")%><br>(<%=rs("ACTION_CHINESE_NAME")%>)</span>
          <input name="action_id<%=i%>" type="hidden" id="action_id<%=i%>" value="<%=rs("NID")%>">
		  <input name="action_name<%=i%>" type="hidden" id="action_name<%=i%>" value="<%=rs("ACTION_NAME")%>(<%=rs("ACTION_CHINESE_NAME")%>)">
          <input name="action_purpose<%=i%>" type="hidden" id="action_purpose<%=i%>" value="<%=rs("ACTION_PURPOSE")%>">
		  <input name="action_null<%=i%>" type="hidden" id="action_null<%=i%>" value="<%=rs("NULL_ALLOW")%>">
		  <input name="action_number<%=i%>" type="hidden" id="action_number<%=i%>" value="<%=rs("ELEMENT_NUMBER")%>">
          <input name="valid_value<%=i%>" type="hidden" id="valid_value<%=i%>" value="<%=getValidActionValue(session("jobnumber"),part_number_id,rs("NID"),max_quantity,min_quantity)%>">
          <input name="max_quantity<%=i%>" type="hidden" id="max_quantity<%=i%>" value="<%=max_quantity%>">
          <input name="min_quantity<%=i%>" type="hidden" id="min_quantity<%=i%>" value="<%=min_quantity%>">
		</td>
      	<td height="20" nowrap>
    <%select case UCase(rs("ELEMENT_TYPE"))
	  case "TEXT"
	  	if rs("ACTION_PURPOSE")="7" then
			has2DCode=true
		end if
		if rs("ACTION_PURPOSE")="8" then
			hasDefect2DActn=true
		end if
		if rs("ACTION_PURPOSE")="5" then 'Job quantity
			if initail_quantity_type="Con" then
				initail_quantity=getInitialQuantity(session("jobnumber"),session("SHEET_NUMBER"),session("JOB_TYPE"),previous_station_id)
			%>
        	<input name="action_value<%=i%>" type="hidden" value="<%=initail_quantity%>"><%=initail_quantity%>
        <%  else%>
        	<input name="action_value<%=i%>" id="action_value<%=i%>" type="text" size="15" autocomplete="off" onChange="numbercheck2(this,<%=i%>);">
        <%	end if
		elseif rs("ACTION_PURPOSE")="6" then 'Rework quantity%>
       		<input name="action_value<%=i%>" id="action_value<%=i%>" type="text" size="15" autocomplete="off" onChange="numbercheck2(this,<%=i%>);">
        <%else %>
			<table width="0" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<%txtSize = 100/cint(rs("ELEMENT_NUMBER"))
				  if txtSize >30 then
					txtSize = 30
				  elseif txtSize <10 then
					txtSize = 10
				  end if
				for k=1 to cint(rs("ELEMENT_NUMBER"))%>
				<td><input name="action_boolean<%=i%>_<%=k%>" type="hidden" value="0">
					<input name="action_value<%=i%>_<%=k%>" id="action_value<%=i%>_<%=k%>" type="text" onFocus="this.select();" 
					onKeyDown="if(event.keyCode==13){event.keyCode=9;}"
					onChange="<%if rs("ACTION_PURPOSE")="4" then%>numbercheck2(this,<%=i%>);<%end if%>valuecheck(<%=i%>,<%=k%>,'<%=rs("ACTION_PURPOSE")%>','<%=rs("ACTION_NAME")%>');<%if rs("ACTION_PURPOSE")="7" then %>check2DCode(this,action_boolean<%=i%>_<%=k%>);<%end if%>;<%if rs("ACTION_PURPOSE")="8" then %>defect2DCode(this);<%end if%>" 
					<%if rs("ACTION_PURPOSE")="4" then%>value="0"<%end if%> size="<%=txtSize%>" autocomplete="off">
				</td>
				<%if(k mod 5)=0 then%></tr><tr><%end if%>
				<%next%>
			  </tr>
			</table>
		<%
		end if
	  end select
	%><span class="strongred" id="errorinsert<%=i%>"></span></td>
    </tr>
	<%i=i+1
	rs.movenext
	wend
  rs.close
  end if
  set rs=nothing%>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>	
  </tr>
  <%if has2DCode then%>
  <tr>
    <td height="20">2D Code Remain Qty<br>二维码剩余数量</td>
	<td height="20"><%=get2DCodeRemainQty()%></td>	
  </tr>
  <%end if%>
  <tr>
    <td height="20" colspan="2"><div align="center">
	 
      <input name="station_default_transaction_type" type="hidden" id="station_default_transaction_type" value="<%=station_default_transaction_type%>">
      <input name="action_count" type="hidden" id="action_count" value="<%=i-1%>">
      <input name="current_station_id" type="hidden" id="current_station_id" value="<%=current_station_id%>">
	  <input name="part_number_id" type="hidden" id="part_number_id" value="<%=part_number_id%>">
	  <input name="station_start_time" type="hidden" id="station_start_time" value="<%=station_start_time%>" >
  <%
  
  function get2DCodeRemainQty()
	'get 2D code count of job
	set rs2DCode=server.createobject("adodb.recordset")
	sql2DCode="select count(1) from job_2d_code where job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"'"
	rs2DCode.open sql2DCode,conn,1,3
	job2DCodeQty = clng(rs2DCode(0))
	rs2DCode.close
	
	sql2DCode="select nvl(sum(defect_quantity),0) from job_defectcodes_temp where job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' and enter_station_id='"&current_station_id&"' " 
	rs2DCode.open sql2DCode,conn,1,3
	jobDefectQty = clng(rs2DCode(0))
	rs2DCode.close
	set rs2DCode=nothing
	get2DCodeRemainQty = clng(station_start_quantity) - job2DCodeQty - jobDefectQty
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
	matRecord=false
  	SQL="Select MATERIAL_COUNT from MATERIAL_COUNT where ROUTING_ID='"&getNewRoutineId(part_number_id)&"' and STATION_ID='"&getNewStationId(current_station_id)&"'" 		 
	set rsF=server.createobject("adodb.recordset")
	rsF.open SQL,conn,1,3
	if not rsF.eof then	
		if cint(rsF("MATERIAL_COUNT").value)>0 then
			matRecord=true
		end if
	end if
	rsF.close
	if matRecord then
  %>
   	<input name="Record" id="Record" type="button" value="Material Info"  onClick="ToRecord();">&nbsp;
    
   	<input name="Next" type="submit" id="Next" value="OK 确定" disabled="disabled">	
   	<%
	else%>
    <input name="Record" id="Record" type="button" value="请输入胶水&铜线"  onClick="ToRecordNew();">
   	<input name="Next" type="submit" id="Next" value="OK 确定">
    <%end if%>
	&nbsp;</div>
	</td></tr>
</table>
</form>
<script language="javascript">
function ToRecord()
{
	currentStationId=document.getElementById("current_station_id").value;
	part_number_id=document.getElementById("part_number_id").value;
	toUrl="Material_Count_Record.asp?currentStationId="+currentStationId+"&part_number_id="+part_number_id;
	window.showModalDialog(toUrl,window,"dialogHeight:400px;dialogWidth:500px;resizable=yes");
}
function ToRecordNew()
{
	
	toUrl="Material_Count_Record_New.asp";
	window.showModalDialog(toUrl,window,"dialogHeight:500px;dialogWidth:1200px;resizable=yes");
}
function formSubmit() {
    if (formcheck()&&submitDefect()) {		
        submitonce("form1");
    } else {
        return false;
    }
    return true;
}

function removeDefect(){
	if(document.all.defect_code.value){
    	var objTable = document.getElementById("tb_defect_code");
    	for(var i=2;i<objTable.rows.length;i++){
			if(document.all.defect_code.value == objTable.rows(i).cells(0).innerText){
				totalDefectQty=totalDefectQty - parseInt(objTable.rows(i).cells(2).innerText);
				objTable.deleteRow(i);
				document.all.defect_code.value="";
				document.all.recent_defect_id.value="";
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
var totalDefectQty=0;
function addDefect(){
	//var objTable = document.getElementById("tb_defect_code");
	  //var  dd=objTable.rows(2).cells(1).innerText ;
	  <%if hasDefect2DActn then%>
	  var defect2DCode=document.all.recent_defect_id.value;
	  if(defect2DCode!=""){
		alert("此工站一次只能输入一个不良代码");
		document.all.defect_code.value = "";
		document.all.defect_qty.value = "";
		return false;	
	           }	
	<%end if%>
	  	
	  
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
	}else if(document.all.defect_qty.value<0){
		alert("Input Quantity is not a negative!\n输入的数量不是为负数！");
		document.all.defect_qty.select();
		return false;				
	}else{
		var deftInfo = window.showModalDialog("GetValueByKey.asp?key=DefectCode&keyValue="+document.all.defect_code.value,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
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
			document.all.recent_defect_id.value=aryDeftInfo[0];
			totalDefectQty=totalDefectQty+parseInt(document.all.defect_qty.value);
			document.all.defect_code.value = "";
			document.all.defect_qty.value = "";
			document.all.defect_code.focus();
		}
	}	
}

function submitDefect(){
	var objTable = document.getElementById("tb_defect_code");
	var deftInfoStr="";
	var deftCount=0;
	var totalDefectQty=0;
	for(var i=2;i<objTable.rows.length;i++){    
		deftInfoStr=deftInfoStr+"&defect_code"+deftCount+"="+objTable.rows(i).cells(4).innerText;
		//deftInfoStr=deftInfoStr+"&defect_name"+deftCount+"="+objTable.rows(i).cells(1).innerText;
		deftInfoStr=deftInfoStr+"&defect_quantity"+deftCount+"="+objTable.rows(i).cells(2).innerText;
		deftInfoStr=deftInfoStr+"&station_id"+deftCount+"="+objTable.rows(i).cells(3).innerText;
		deftCount++;
		totalDefectQty = totalDefectQty + new Number(objTable.rows(i).cells(2).innerText);
	}
	//check total defect count
	if(totalDefectQty > <%=station_start_quantity%>){
		document.all.errorinsert.innerText="Total Defect Code Quantity exceeds Job Quantity!\nDefect Code数量合计超过Job Quantity！";
		document.all.defect_code.focus();
		return false;
	}
<%if has2DCode then%>	
	if(ary2DObj.length+totalDefectQty><%=get2DCodeRemainQty()%>){
		alert("Input 2D code qty exceeds remain qty!\n输入的二维码数量超过剩余数量!");
		return false;
	}
<%end if%>
	<%if hasDefect2DActn then%>
	if(defect2DCodeQty!=totalDefectQty){
		alert("Input defect 2D code qty does not equal total defect qty.\n输入的二维码数量与总的不良数不一致。");
		return false;	
	}	
	<%end if%>
	deftInfoStr = "defect_count="+deftCount+deftInfoStr;
	document.form1.action=document.form1.action+"?"+deftInfoStr;
	return true;
}
var defect2DCodeQty=0;
function defect2DCode(obj){
	isValid = true;
	var deftId=document.all.recent_defect_id.value;	
	if(obj.value){
		if(!deftId){
			alert("Please input defect first.\n请先输入不良.");
			obj.value="";
			obj.blur();
			obj.focus();
			return false;
		}else if(totalDefectQty<=defect2DCodeQty){
			alert("Input 2D code qty exceeds total defect qty.\n输入二微码的数量已经达到不良总数.");
			obj.value="";
			obj.blur();
			obj.focus();
			return false;
		}
		
		
		if(str2DValue.indexOf(obj.value+",") == -1){			
		
				str2DValue = str2DValue + obj.value + ",";
			//}
		}else{
			isValid = false;
		}			
		
		if(!isValid){
		alert("This 2D code("+obj.value+") is existed.\n该二维码("+obj.value+")已经存在。");
		obj.style.background="#CC99FF";
		obj.blur();
		obj.select();
		objActBL.value="3";
		return false;
	    }else{
		obj.style.background="#FFFFFF";		
	    }	
		
		
		
		
		
		
		
		
		var rtnValue = window.showModalDialog("GetValueByKey.asp?key=Defect2DCode&keyValue="+obj.value+"&defect_id="+deftId,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		if(rtnValue.indexOf("Error-")==0){
			alert(rtnValue.substring(6));
			obj.value="";
			obj.blur();
			obj.focus();
			return false;
		}else{
			defect2DCodeQty=defect2DCodeQty+1;	
		}
	}	
}

var str2DValue="";
function check2DCode(object,objActBL){
	isValid = true;
	if(object.value != ""){
		object.value = object.value.toUpperCase().trim();	
		 if(object.value.length!=17) 
              {
              	alert("对不起，2D码位数不对！");
	          	object.style.background="#CC99FF";
				object.blur();
				object.select();
				objActBL.value="3";
				return false;
              } 
		
	 var FACTORY_CODE =document.all.FACTORY_CODE.value;
	 var FACTORYCODE= left(object.value,3);
	     FACTORYCODE= FACTORY_CODE.indexOf(FACTORYCODE);
		// alert(FACTORYCODE);
	   if (FACTORYCODE == -1)
             {
            alert("对不起！北京工厂代码应为:DYD");
	        object.style.background="#CC99FF";
			object.blur();
			object.select();
			objActBL.value="3";
			return false;
             } 
		
	//	var jobValue=document.form1.txtSubJobList.value;
		
	 var CODE_LINENAME =document.all.CODE_LINENAME.value;
	
	 var CodeLineName= mid(object.value,7,1);
	
	     CodeLineName= CODE_LINENAME.indexOf(CodeLineName);
	     
	   if (CodeLineName ==-1)
	   {
		alert("对不起！线别不对！");
		object.style.background="#CC99FF";
		object.blur();
		object.select();
		objActBL.value="3";
		return false;
	   } 
	
	 var CODE_NAME =document.all.CODE_NAME.value; 
	 var CodeName= mid(object.value,11,4);
		 CodeName= CODE_NAME.indexOf(CodeName);
	   if (CodeName ==-1)
	   {
		alert("产品不对，此产品不是这个项目的！");
		object.style.background="#CC99FF";
		object.blur();
		object.select();
		objActBL.value="3";
		return false;
	   } 
	 var VERSION_NUMBER =document.all.VERSION_NUMBER.value; 
	 var VERSIONNUMBER= mid(object.value,15,1);
		 VERSIONNUMBER= VERSION_NUMBER.indexOf(VERSIONNUMBER);
	   if (VERSIONNUMBER ==-1)
	   {
		alert("版本号不对！");
		object.style.background="#CC99FF";
		object.blur();
		object.select();
		objActBL.value="3";
		return false;
	   } 
		if(str2DValue.indexOf(object.value+",") == -1){			
			//var rtnValue = window.showModalDialog("GetValueByKey.asp?key=2DCode&keyValue="+object.value,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			//if(rtnValue != ""){
			//	isValid = false;
			//}esle{
				str2DValue = str2DValue + object.value + ",";
			//}
		}else{
			isValid = false;
		}				
	}
	if(!isValid){
		alert("This 2D code("+object.value+") is existed.\n该二维码("+object.value+")已经存在。");
		object.style.background="#CC99FF";
		object.blur();
		object.select();
		objActBL.value="3";
		return false;
	}else{
		object.style.background="#FFFFFF";		
	}	
}

var all2DCode="";
var ary2DObj=new Array();
function checkAll2DCode(){
	if(all2DCode!=""){	
		var rtnValue = window.showModalDialog("GetValueByKey.asp?key=2DCode&keyValue="+all2DCode,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		if(rtnValue != ""){
			var ary2DCode = all2DCode.split(",");			
			for(var i=0;i<ary2DCode.length-1;i++){
				if(rtnValue.indexOf(ary2DCode[i]) != -1){
					document.getElementById(ary2DObj[i]).style.background="#CC99FF";					
				}
			}
			alert("You have inputed exists 2D code!\n输入了已存在的二维码！");
			return false;
		}
	}
	return true;	
}

function formcheck()
{	
	//check action value
	var actCount = document.all.action_count.value;
	all2DCode="";
	for(var i=1;i<=actCount;i++){
		actPurpose = eval("document.all.action_purpose"+i).value;
		actName = eval("document.all.action_name"+i).value;
		actNumber = eval("document.all.action_number"+i).value;
		actNull = eval("document.all.action_null"+i).value;
		//Job quantity or Rework quantity
		if(actPurpose=="5"||actPurpose=="6"){
			obj = eval("document.all.action_value"+i);
			if(obj.value == ''){
				alert(actName+" cannot be blank!\n"+actName+"不得为空！");
				obj.blur();
				eobj.select();
				return false;
			}
		}else {
			var hasValue=false;			
			for(var j=1;j<=actNumber;j++){
				obj = eval("document.all.action_value"+i+"_"+j);
				if(obj.value){				
					hasValue=true;					
				}
				if(eval("document.all.action_boolean"+i+"_"+j).value=='2'){
					alert("Value of "+actName+" is invalid!\n"+actName+"输入值错误！");
					obj.blur();
					obj.focus();
					return false;
				}else if(eval("document.all.action_boolean"+i+"_"+j).value=='3'){
					alert("You have inputed exists 2D code!\n输入了已存在的二维码！");
					return false;
				}
				if(actPurpose=="7" && obj.value){
					all2DCode = all2DCode + obj.value + ",";
					ary2DObj[ary2DObj.length] = obj.id;
				}
			}
			if(actNull=="0" && !hasValue){
				alert(actName+" cannot be blank!\n"+actName+" 不得为空！");				
				return false;
			}			
		}	
	}	
	if (confirm("请确定输入的不良代码和数量是否正确！"))
   { 
  
		return checkAll2DCode();
	
	}
	else
	{
	
	  return false;
	}
	

	
	
	
	
	
}

function pageload()
{
	if(document.all.action_count.value>0 && typeof document.all.action_value1_1 != "undefined"){
		document.all.action_value1_1.focus();
	}else{	
		document.form1.defect_code.focus();	
	}
}


function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}



function left(mainStr,lngLen) {
    if (lngLen>0) {return mainStr.substring(0,lngLen)}
    else{return null}
}

function right(mainStr,lngLen) {
    if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) {
    return mainStr.substring(mainStr.length-lngLen,mainStr.length)}
    else{return null}
}


function mid(mainStr,starnum,endnum){
    if (mainStr.length>=0){
        return mainStr.substr(starnum,endnum)
    }else{return null}
} 
function trim(str){ //删除左右两端的空格 
return str.replace(/(^\s*)|(\s*$)/g, ""); 
} 


</script>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>