<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/GetStationName.asp" -->
<!--#include virtual="/Functions/GetInitialQuantity.asp" -->
<!--#include virtual="/Functions/GetAutoQuantity.asp" -->
<!--#include virtual="/Functions/getValidActionValue.asp" -->
<%
pagename="Station3.asp"
job_start_time=now()

if session("STATIONS_ROUTINE")="1" and request.Form("selected_station")="" then'if routine is repeated
response.Redirect("Station1.asp?errorstring=At lease one station is selected!<br>至少要选择一个站！")
end if
'get job's properties.
current_station_id=session("CURRENT_STATION_ID")
SQL="select J.PART_NUMBER_ID,J.STATUS,J.START_TIME,J.FIRST_STATION_ID,J.CURRENT_STATION_ID,nvl(J.PREVIOUS_STATION_ID,J.CURRENT_STATION_ID) as PREVIOUS_STATION_ID,J.PREVIOUS_STATION_TRANSACTION,J.PREVIOUS_STATION_CLOSE_TIME,INTERVAL_SKIP_STATION_ID from JOB J where J.JOB_NUMBER='"&session("JOB_NUMBER")&"' and J.SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	job_start_time=rs("START_TIME")
	part_number_id=rs("PART_NUMBER_ID")
	first_station_id=rs("FIRST_STATION_ID")
	'current_station_id=rs("CURRENT_STATION_ID")
	previous_station_id=rs("PREVIOUS_STATION_ID")
	previous_station_transaction=rs("PREVIOUS_STATION_TRANSACTION")
	previous_station_close_time=rs("PREVIOUS_STATION_CLOSE_TIME")
	this_interval_skip_station_id=rs("INTERVAL_SKIP_STATION_ID")
else
	job_start_time=now()
	part_number_id=session("PART_NUMBER_ID")
	first_station_id=session("FIRST_STATION_ID")
	'current_station_id=session("CURRENT_STATION_ID")
	previous_station_id=""
	previous_station_transaction=""
	previous_station_close_time=""
	this_interval_skip_station_id=null
end if
rs.close

'if prevoius station is conjuctive, check whether normal and twin sub job is closed
if previous_station_transaction="2" then	
	SQLO="select SheetNumber from tbl_MES_LotMasterSub where WipEntityName='"&session("JOB_NUMBER")&"' and SheetType=1 order by SheetNumber"
	rs.open SQLO,connTicket,1,3

	if not rs.eof then
		i=1
		while not rs.eof
			if i=cint(session("SHEET_NUMBER")) then
				conjunctive_sheetnumber=trim(rs("SheetNumber"))
			end if
		i=i+1
		rs.movenext
		wend
	end if
	rs.close

	SQL="select status from job where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&conjunctive_sheetnumber&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
	conjuctive_status=rs("status")
	end if
	rs.close
	
	if conjuctive_status<>"1" then
	response.Redirect("ActionError.asp?alerttype=conjuctive&alertmessage="&session("JOB_NUMBER")&"-"&repeatstring(conjunctive_sheetnumber,"0",3))
	end if
end if

if isnull(this_interval_skip_station_id) then' if job is unlocked, which station will not check max interval.
interval_skip_station_id=""
else
interval_skip_station_id=this_interval_skip_station_id
end if

'get job quantity's property in this station.
SQL="select TRANSACTION_TYPE,INITAIL_QUANTITY_TYPE from STATION where NID='"&current_station_id&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	initial_quantity_type=rs("INITAIL_QUANTITY_TYPE")
	station_default_transaction_type=rs("TRANSACTION_TYPE")
else
	initial_quantity_type=0
	station_default_transaction_type="0" 'default is 0 (Compulsory)
end if
rs.close

'get max interval property of this part.
SQL="select STATIONS_INDEX,MAX_INTERVAL,INITIAL_QUANTITY from PART where NID='"&part_number_id&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	stations_index=rs("STATIONS_INDEX")
	max_interval=rs("MAX_INTERVAL")
	first_initial_quantity=rs("INITIAL_QUANTITY")
end if
rs.close

if session("STATIONS_ROUTINE")="1" then
	if request.Form("sequence_type"&request.Form("selected_station"))="new" then 'if stations_routine of Part is 1, set current stations as selected station.
		current_station_id=request.Form("station_id"&request.Form("selected_station"))
		session("CURRENT_STATION_ID")=current_station_id
		SQL="select REPEATED_SEQUENCE from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' order by REPEATED_SEQUENCE desc"'get the REPEATED_SEQUENCE of this station in the same job and station list.
		rs.open SQL,conn,1,3
		if not rs.eof then
		session("REPEATED_SEQUENCE")=cint(rs("REPEATED_SEQUENCE"))+1
		else
		session("REPEATED_SEQUENCE")=1
		end if
		rs.close
	else
	current_station_id=request.Form("station_id"&request.Form("selected_station"))
	session("CURRENT_STATION_ID")=current_station_id
	session("REPEATED_SEQUENCE")=request.Form("repeated_sequence"&request.Form("selected_station"))
	end if
else
	SQL="select STATUS,REPEATED_SEQUENCE from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		current_station_status=rs("STATUS")
		station_repeated_sequence=rs("REPEATED_SEQUENCE")
	else
		current_station_status="0"
	end if
	rs.close
	
	if station_default_transaction_type="1" and current_station_status="0" then
		SQL="select REPEATED_SEQUENCE from JOB_ACTIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"'" 'get the default transaction_type of current station
		rs.open SQL,conn,1,3
		if not rs.eof then
			session("REPEATED_SEQUENCE")=cint(rs("REPEATED_SEQUENCE"))+1
		else
			session("REPEATED_SEQUENCE")=1
		end if
		rs.close
	elseif station_default_transaction_type="1" and current_station_status<>"0" then
		session("REPEATED_SEQUENCE")=station_repeated_sequence
	else
		session("REPEATED_SEQUENCE")=1
	end if
end if

'get station's properties.
SQL="select STATUS from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
rs.open SQL,conn,1,3
if rs.eof then 'no record, this station is new opened
	if current_station_id<>first_station_id then
		'get the max interval time between current station and previous station
		if max_interval<>"" and current_station_id<>first_station_id and interval_skip_station_id<>current_station_id then 
			amax=split(max_interval,",")
			astations=split(stations_index,",")
			for j=0 to ubound(astations)
				if current_station_id=astations(j) then
					current_max_interval=amax(j-1)
				end if
			next
		end if
		if previous_station_close_time<>"" and current_max_interval<>"" then 'whether interval between cuurent station and previous station exceed max interval
			if datediff("n",previous_station_close_time,job_start_time)>cint(current_max_interval) then
			response.Redirect("Station_Interval.asp")
			end if
		end if
	end if
else 'this station is opened
	if rs("STATUS")<>"0" then 'this station has started
	    response.Redirect("Station5.asp")
	end if
end if
rs.close

%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JValueCheck.js" type="text/javascript"></script>
<script language="javascript">
    function formSubmit() {
        if (formcheck()) {
            submitonce("form1");
        } else {
            return false;
        }
        return true;
    }
	
	function ToRecordNew()
{
	
	toUrl="Material_Count_Record_New.asp";
	window.showModalDialog(toUrl,window,"dialogHeight:600px;dialogWidth:1200px;resizable=yes");
}
</script>


</head>

<body onLoad="pageload()" onKeyPress="keyhandler()"  bgcolor="#339966">
<span id="erroralarm"></span>
<form action="/KES1/Station4.asp" method="post" name="form1" target="_self" onsubmit="return formSubmit()">
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td colspan="4" class="t-t-DarkBlue">Scan station info 扫描站信息</td>
  </tr>
  <tr>
    <td colspan="4" class="t-t-Borrow">Operator 操作员:<% =session("operator") %>(<% =session("code") %>)</td>        
  </tr>
  <tr>
    <td width="26%">Job Number 工单号</td>
    <td width="25%"><% =session("JOB_NUMBER") %>-<% =repeatstring(session("SHEET_NUMBER"),"0",3) %></td>
    <td width="24%">Part Number 型号  </td>
    <td width="25%"><% =session("PART_NUMBER_TAG") %></td>
  </tr>
  <tr>
    <td>Line Name 线别 </td>
    <td><% =session("LINE_NAME") %></td>
    <td height="20">Job Start Time 工单开始时间 </td>
    <td height="20"><% =job_start_time%></td>
  </tr>
</table>
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
    <td width="17%" height="20">Station Name 站名</td>
    <td height="20" colspan="2"><span class="red"><%=getStationName(true,current_station_id,chinesename)%>&nbsp;<%session("CURRENT_STATION_NAME")=chinesename%><%=chinesename%>&nbsp;<% =session("REPEATED_SEQUENCE") %></span></td>
    </tr>
<%if session("NEW_JOB")="true" then%>	
	<tr>
		<td height="20">Shift 班别</td>
		<td height="20" colspan="2">
			<select id="sltShift" name="sltShift" style="width:60px">
				<option value=""></option>
				<option value="A">A</option>
				<option value="B">B</option>
				<option value="C">C</option>
				<option value="D">D</option>
			</select>
		</td>
	</tr>
	<tr>
		<td height="20">Plastic Supplier<br>塑料件供应商</td>
		<td height="20" colspan="2">
			<select id="sltPlasticSupplier" name="sltPlasticSupplier">
				<option value=""></option>
				<option value="H">H-hip</option>
				<option value="S">S-swiftronics</option>
				<option value="Y">Y-yingcheng</option>
				<option value="C">C-concraft</option>
			</select>
		</td>
	</tr>
<%end if%>	
  <%
  SQL="select ACTIONS_INDEX from STATION where NID='"&current_station_id&"'"
  rs.open SQL,conn,1,3
  if not rs.eof then
  	ACTIONS_INDEX=rs("ACTIONS_INDEX")	
  end if
  rs.close

  if not isnull(ACTIONS_INDEX) then
	aaction=split(ACTIONS_INDEX,",")
  	SQL="select NID,ACTION_PURPOSE,ACTION_NAME,ACTION_CHINESE_NAME,NULL_ALLOW,ELEMENT_TYPE,ELEMENT_NUMBER,ACTION_TYPE,WITH_LOT from ACTION where STATION_POSITION=0 and STATUS=1 and NID in ('"&replace(ACTIONS_INDEX,",","','")&"')"
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
	<input name="action_name<%=i%>"  type="hidden" id="action_name<%=i%>" value="<%=rs("ACTION_NAME")%>">
	<input name="action_purpose<%=i%>" type="hidden"  id="action_purpose<%=i%>" value="<%=rs("ACTION_PURPOSE")%>">
	<input name="action_null<%=i%>"  type="hidden" id="action_null<%=i%>" value="<%=rs("NULL_ALLOW")%>">
	<input name="action_number<%=i%>" type="hidden" id="action_number<%=i%>" value="<%=rs("ELEMENT_NUMBER")%>">
	<input name="valid_value<%=i%>" type="hidden" id="valid_value<%=i%>" value="<%=getValidActionValue(session("JOB_NUMBER"),part_number_id,rs("NID"),max_quantity,min_quantity)%>">
	<input name="max_quantity<%=i%>" type="hidden" id="max_quantity<%=i%>" value="<%=max_quantity%>">
	<input name="min_quantity<%=i%>" type="hidden" id="min_quantity<%=i%>" value="<%=min_quantity%>">
	</td>
    <td height="20" nowrap>
	<%select case rs("ELEMENT_TYPE")
		case "TEXT"			
			if rs("ACTION_PURPOSE")="5" then 'Job Quantity
				if initial_quantity_type="Con" and current_station_id<>first_station_id and session("STATIONS_ROUTINE")="0" then
					initial_quantity=getInitialQuantity(session("JOB_NUMBER"),session("SHEET_NUMBER"),session("JOB_TYPE"),previous_station_id)
					if previous_station_transaction="2" then'conjuctive station
						conjuctive_quantity=getInitialQuantity(session("JOB_NUMBER"),conjunctive_sheetnumber,session("JOB_TYPE"),previous_station_id)
						mod_quantity=(cint(initial_quantity)+cint(conjuctive_quantity)) mod 2
						initial_quantity=(cint(initial_quantity)+cint(conjuctive_quantity)-cint(mod_quantity))/2
					end if
				%>
					<input name="action_value<%=i%>" type="hidden" value="<%=initial_quantity%>"><%=initial_quantity%><%if mod_quantity<>"" then%>余<%=mod_quantity%><%end if%>
				<%else 'first station
					initial_quantity=getAutoQuantity(session("JOB_NUMBER"),session("SHEET_NUMBER"),session("JOB_TYPE"))						
					if initial_quantity<>"" then
						session("JOB_START_QUANTITY")=initial_quantity
					end if
				%>
					<input name="action_value<%=i%>" type="text" id="action_value<%=i%>" value="<%=initial_quantity%>" size="15" autocomplete="off"  readonly="true">	
				<%end if
			else%>
				<table width="0" border="0" cellpadding="0" cellspacing="0">
				  <tr>
					<%txtSize = 60/cint(rs("ELEMENT_NUMBER"))
					  if txtSize >30 then
					  	txtSize = 30
					  elseif txtSize <10 then
						txtSize = 10
					  end if
					for k=1 to cint(rs("ELEMENT_NUMBER"))%>
					<td><input name="action_boolean<%=i%>_<%=k%>" type="hidden" value="0">
						<input name="action_value<%=i%>_<%=k%>" id="action_value<%=i%>_<%=k%>" type="text" onFocus="this.select();" 
							onKeyDown="if(event.keyCode==13){event.keyCode=9;}"				
							onChange="<%if rs("ACTION_PURPOSE")="4" then%>numbercheck2(this,<%=i%>);<%end if%>valuecheck(<%=i%>,<%=k%>,'<%=rs("ACTION_PURPOSE")%>','<%=rs("ACTION_NAME")%>')" 
						<%if rs("ACTION_PURPOSE")="4" then%>value="0"<%end if%> size="<%=txtSize%>" autocomplete="off">						
					</td>
					<%if(k mod 5)=0 then%></tr><tr><%end if%>
					<%next%>
				  </tr>
				</table>
	  	<%end if
		end select
		%><span class="strongred" id="errorinsert<%=i%>"></span></td>
    	<td><div align="center"><%if rs("ACTION_TYPE")="Key" then%>输入<%else%>扫描<%end if%></div></td>
  	</tr>
  <%i=i+1
	rs.movenext
	wend
  	rs.close
  end if
  set rs=nothing%>
  <tr>
    <td height="20" colspan="3">&nbsp;</div></td>
  </tr>
  <tr>
	<td height="20" colspan="3"><div align="center">
    	<input name="station_default_transaction_type" type="hidden" id="station_default_transaction_type" value="<%=station_default_transaction_type%>">
      	<input name="action_count" type="hidden" id="action_count" value="<%=i-1%>">
      	<input name="stations_routine" type="hidden" id="stations_routine" value="<%=stations_routine%>">
        <%if session("NEW_JOB")="true" then
		
		
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
		

        
      	<input name="Record" id="Record" type="button" value="请输入胶水&铜线"  onClick="ToRecordNew();">
        <input name="Next" type="submit" value="Next 下一步" disabled="disabled">
        <%else%>
        <input name="Next" type="submit" value="Next 下一步" >
        <%end if
		else%>		
		 <input name="Next" type="submit" value="Next 下一步" >
		
 <% end if%>
		

		&nbsp;
		<input type="reset" name="Reset" value="Reset 重置">
    	&nbsp;
    	<input type="button" name="Button" value="Back 返回" onClick="javascript:location.href='/KES1/Station1.asp'">
    </div></td>
   </tr>
</table>
</form>
<script language="javascript">
function pageload()
{
	var actCount = document.all.action_count.value;
	try{
		if(actCount>1 && document.all.action_purpose2.value != "5"){
			document.all.action_value2_1.focus();		
		}else{
			document.all.Next.focus();
		}
	}catch(e){}
}
function formcheck()
{
<%if session("NEW_JOB")="true" then%>
	if(!document.all.sltShift.value){
		alert("Shift cannot be blank!\n班别不能为空");
		return false;
	}
	if(!document.all.sltPlasticSupplier.value){
		alert("Plastic Supplier cannot be blank!\n塑料件供应商不能为空");
		return false;
	}
<%end if%>	
	//check action value
	var actCount = document.all.action_count.value;
	for(var i=1;i<=actCount;i++){
		actPurpose = eval("document.all.action_purpose"+i).value;
		actName = eval("document.all.action_name"+i).value;
		actNumber = eval("document.all.action_number"+i).value;
		actNull = eval("document.all.action_null"+i).value;
		//Job quantity
		if(actPurpose=="5"){
			var objAct = eval("document.all.action_value"+i);			
			if(typeof objAct == "text" && objAct.value == ''){
				alert(actName+" cannot be blank!\n"+actName+"不得为空！");
				objAct.select();
				return false;
			}
		}else if(actNull=="0"){
			hasValue=false;
			for(var j=1;j<=actNumber;j++){
				value = eval("document.all.action_value"+i+"_"+j).value;
				if(value){				
					hasValue=true;					
				}
				if(eval("document.all.action_boolean"+i+"_"+j).value=='2'){
					alert("Value of "+actName+" is invalid!\n"+actName+"输入值错误！");
					eval("document.all.action_value"+i+"_"+j).select();
					return false;
				}
			}
			if(!hasValue){
				alert(actName+" cannot be blank!\n"+actName+" 不得为空！");				
				return false;
			}			
		}	
	}		
	return true;
}
</script>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->