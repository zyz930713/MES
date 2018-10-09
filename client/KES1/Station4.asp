<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStationName.asp" -->
<!--#include virtual="/Functions/GetJobUVSBOM.asp" -->
<%
pagename="Station4.asp"
station_start_time=now()
islocked=false
lockalert=""

'get station start quantity
'response.write "count:"+request.Form("action_count")+"<br>"
for i=1 to cint(request.Form("action_count"))
	'response.write "purpose:"+request.Form("action_purpose"&i)+"<br>"
	'response.write "value:"+request.Form("action_value"&i)+"<br>"
	if request.form("action_purpose"&i)="5" then
		station_start_quantity=clng(request.Form("action_value"&i))
	end if
next

'is machine, lot, code locked
for i=1 to cint(request.Form("action_count"))
	select case	request.form("action_purpose"&i)
		case "1"
		SQL="select LOCKED from MACHINE where MACHINE_NUMBER='"&trim(request.Form("action_value"&i))&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
			if rs("LOCKED")="1" then
				islocked=true
				lockalert=lockalert&trim(request.Form("action_value"&i))&"<br>"
			end if
		end if
		rs.close
		case "3"
		SQL="select LOCKED_LOT from MATERIAL where MATERIAL_NUMBER='"&trim(request.Form("action_value"&i))&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
			for j=1 to cint(request.form("action_number"&i))
				if instr(rs("LOCKED_LOT"),trim(request.Form("action_value"&i&"_"&j)))<>0 then
				islocked=true
				lockalert=lockalert&trim(request.Form("action_value"&i))&"<br>"
				end if
			next
		end if
		rs.close
	end select
next
if islocked=true then
response.Redirect("ActionError.asp?alerttype=lock&alertmessage="&lockalert)
end if

'is materila in BOM
if session("FACTORY_ID")="FA00000002" or session("FACTORY_ID")="FA00000003" then'only KE need to validate
for i=1 to cint(request.Form("action_count"))
	if request.form("action_purpose"&i)="2" then
		SQLProd="select B.Inventory_Item_Name from tbl_Lot_BOM_Item B inner join tbl_MES_LotMaster M on B.WIP_ENTITY_ID=M.WipEntityId where M.WipEntityName='"&session("JOB_NUMBER")&"' and B.Inventory_Item_Name='"&request.Form("action_value"&i)&"'"
		rsProd.open SQLProd,connTicket,1,3
		if  rsProd.eof then
			response.Redirect("ActionError.asp?alerttype=part&alertmessage="&request.Form("action_name"&i))
		end if
		rsProd.close
	end if
next
end if

'create a new record in JOB_MASTER
SQL="select * from JOB_MASTER where JOB_NUMBER='"&session("JOB_NUMBER")&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("ORGANIZATION_ID")=session("ORGANIZATION_ID")
	rs("JOB_NUMBER")=session("JOB_NUMBER")
	rs("PART_NUMBER_ID")=session("PART_NUMBER_ID")
	rs("PART_NUMBER_TAG")=session("PART_NUMBER_TAG")
	rs("FACTORY_ID")=session("FACTORY_ID")
	rs("LINE_NAME")=session("LINE_NAME")
	session("aerror")=session("JOB_TOTAL_START_QUANTITY")
	rs("ASSEMBLY_INPUT_QUANTITY")=session("JOB_TOTAL_START_QUANTITY")
	rs("START_QUANTITY")=clng(session("JOB_TOTAL_START_QUANTITY"))
	rs("INPUT_TIME")=now()
	rs("CREATE_CODE")=session("CREATE_CODE")
	rs("CREATE_NAME")=session("CREATE_NAME")
	rs("ERP_CREATE_TIME")=session("ERP_CREATE_TIME")
	rs("ERP_CREATE_BY")=session("ERP_CREATE_BY")
	rs("ERP_LAST_UPDATE_TIME")=session("ERP_LAST_UPDATE_TIME")
	rs("ERP_LAST_UPDATE_BY")=session("ERP_LAST_UPDATE_BY")	
	rs("ERP_LAST_UPDATE_BY")=session("ERP_LAST_UPDATE_BY")	
	rs("BOM_LABEL")=session("BOM_LABEL")							
	rs.update
end if
rs.close

'save to job table
if(isnull(session("JOB_TYPE"))) then
	session("JOB_TYPE")="N"
end if
SQL="select JOB_NUMBER,SHEET_NUMBER,JOB_TYPE,PART_NUMBER_ID,PART_NUMBER_TAG,FACTORY_ID,LINE_NAME,STATIONS_ROUTINE,JOB_START_QUANTITY,FIRST_STATION_ID,CURRENT_STATION_ID,LAST_STATION_ID,STATIONS_INDEX,STATIONS_TRANSACTION,START_TIME,LEAD_TIME,JOB_PRIORITY,NEW_JOB_TYPE,OPENED_STATIONS_ID,REPEATED_STATIONS_SEQUENCE,JOB_GOOD_QUANTITY,SHIFT,SUPPLIER from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("JOB_NUMBER")=session("JOB_NUMBER")
	rs("SHEET_NUMBER")=session("SHEET_NUMBER")
	rs("JOB_TYPE")=session("JOB_TYPE")
	rs("PART_NUMBER_ID")=session("PART_NUMBER_ID")
	rs("PART_NUMBER_TAG")=session("PART_NUMBER_TAG")
	rs("FACTORY_ID")=session("FACTORY_ID")
	rs("LINE_NAME")=session("LINE_NAME")
	rs("STATIONS_ROUTINE")=session("STATIONS_ROUTINE")
	rs("JOB_START_QUANTITY")=session("JOB_START_QUANTITY")
	rs("JOB_GOOD_QUANTITY")=session("JOB_START_QUANTITY")
	rs("FIRST_STATION_ID")=session("FIRST_STATION_ID")
	rs("CURRENT_STATION_ID")=session("CURRENT_STATION_ID")
	rs("LAST_STATION_ID")=session("LAST_STATION_ID")
	rs("STATIONS_INDEX")=session("STATIONS_INDEX")
	rs("STATIONS_TRANSACTION")=session("STATIONS_TRANSACTION")
	rs("LEAD_TIME")=session("LEAD_TIME")
	rs("START_TIME")=now()
	if(session("JOB_TYPE")="N") then
		rs("JOB_PRIORITY")=session("JOB_PRIORITY")
	else
		rs("JOB_PRIORITY")="5"
	end if 
	rs("NEW_JOB_TYPE")=session("NEW_JOB_TYPE")	
	rs("SHIFT")=request("sltShift")	
	rs("SUPPLIER")=request("sltPlasticSupplier")			 
end if
current_station_id=session("CURRENT_STATION_ID")
job_start_time=rs("START_TIME")
rs("OPENED_STATIONS_ID")=rs("OPENED_STATIONS_ID")&current_station_id&","
rs("REPEATED_STATIONS_SEQUENCE")=rs("REPEATED_STATIONS_SEQUENCE")&session("REPEATED_SEQUENCE")&","
rs.update
rs.close

'get station's properties
SQL="select JOB_NUMBER,SHEET_NUMBER,JOB_TYPE,STATION_ID,START_TIME,STATUS,STATION_START_QUANTITY,GOOD_QUANTITY,OPERATOR_CODE,REPEATED_SEQUENCE from JOB_STATIONS where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
rs.open SQL,conn,3,3
if rs.eof then' it is a new station record
	rs.addnew
	rs("JOB_NUMBER")=session("JOB_NUMBER")
	rs("SHEET_NUMBER")=session("SHEET_NUMBER")
	rs("JOB_TYPE")=session("JOB_TYPE")
	rs("STATION_ID")=current_station_id
	rs("STATUS")=1
	rs("STATION_START_QUANTITY")=station_start_quantity
	rs("GOOD_QUANTITY")=station_start_quantity
	rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
	rs("OPERATOR_CODE")=session("code")
	rs("START_TIME")=station_start_time
	rs.update
else
	rs("START_TIME")=station_start_time
	rs.update
end if
rs.close

'Add by Jack Zhang 2010-8-4 only for Rework and Slapping
if session("JOB_TYPE")<>"N" then
	SQL="select FIRST_STATION_ID from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
	rs.open SQL,conn,3,3
	if rs.recordcount>0 then
		if rs("FIRST_STATION_ID")=current_station_id then
			SQL2="update JOB set JOB_START_QUANTITY=" & station_start_quantity & " where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'" 
			set rs2=server.createobject("adodb.recordset")
			rs2.open SQL2,conn,3,3			
			set rs2=nothing
		end if
	end if
	rs.close
end if
'end add

'if user fresh this page, before re-save actions' info, all old records need to be deleted.
SQL="delete from JOB_ACTIONS where STATION_POSITION=0 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"'"
if session("STATIONS_ROUTINE")="1" then
	SQL=SQL+" and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
end if
rs.open SQL,conn,3,3

'save all actions' info.
SQL="select * from JOB_ACTIONS where STATION_POSITION=0 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
rs.open SQL,conn,3,3
if rs.eof then
	for i=1 to cint(request.Form("action_count"))
		rs.addnew
		rs("JOB_NUMBER")=session("JOB_NUMBER")
		rs("SHEET_NUMBER")=session("SHEET_NUMBER")
		rs("JOB_TYPE")=session("JOB_TYPE")
		rs("STATION_ID")=current_station_id
		rs("ACTION_ID")=request.Form("action_id"&i)
		rs("STATION_POSITION")=0
		rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
		ACTION_VALUE=""
		if request.form("action_purpose"&i)="5" then 'Job quantity
			if request.Form("valid_value"&i)<>"" and request.Form("action_boolean"&i)<>"1" then
				response.Redirect("/Functions/ValueError.asp?erroractions="&request.Form("action_name"&i))
			end if
			ACTION_VALUE=request.Form("action_value"&i)		 
		else
			for j=1 to cint(request.form("action_number"&i))
				if request.Form("action_value"&i&"_"&j)<>"" then
					if request.Form("valid_value"&i)<>"" and request.Form("action_boolean"&i&"_"&j)<>"1" then
						response.Redirect("/Functions/ValueError.asp?erroractions="&request.Form("action_name"&i))
					end if
					if request.Form("action_value"&i&"_"&j)<>"NULL" and request.Form("action_value"&i&"_"&j)<>"0" then
						ACTION_VALUE=ACTION_VALUE&request.Form("action_value"&i&"_"&j)&","
					end if
				end if
			next
			if ACTION_VALUE<>"" then
				ACTION_VALUE=left(ACTION_VALUE,len(ACTION_VALUE)-1)
			end if
			'2D code action
			if request.form("action_purpose"&i)="7" then
				str2DCode = str2DCode& ucase(ACTION_VALUE) &","
			end if
		end if
		rs("ACTION_VALUE")=ucase(ACTION_VALUE)
		rs.update
	next
end if
rs.close

'if user fresh this page, before re-save actions' info, all old records need to be deleted.
if request.Form("station_default_transaction_type")="1" then
	SQL="delete from JOB_ACTIONS_REPEATED where STATION_POSITION=0 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
	rs.open SQL,conn,3,3
	
	'save all actions' info into JOB_ACTIONS_REPEATED, if station's transaction_type is optional.
	SQL="select * from JOB_ACTIONS_REPEATED where STATION_POSITION=0 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
	rs.open SQL,conn,3,3
	if rs.eof then
		for i=1 to cint(request.Form("action_count"))
			rs.addnew
			rs("JOB_NUMBER")=session("JOB_NUMBER")
			rs("SHEET_NUMBER")=session("SHEET_NUMBER")
			rs("JOB_TYPE")=session("JOB_TYPE")
			rs("STATION_ID")=current_station_id
			rs("ACTION_ID")=request.Form("action_id"&i)
			rs("STATION_POSITION")=0
			rs("OPERATOR_CODE")=session("code")
			rs("RECORDED_TIME")=now()
			rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
			ACTION_VALUE=""
			if request.form("action_purpose"&i)<>"3" and request.form("action_purpose"&i)<>"4" then 'not Material Lot Number and Material Quantity, but is Machine Code or Material Part Number or Job Quantity
				if request.Form("valid_value"&i)<>"" and request.Form("action_boolean"&i)<>"1" then
				response.Redirect("/Functions/ValueError.asp?erroractions="&request.Form("action_name"&i))
				end if
				rs("ACTION_VALUE")=ucase(request.Form("action_value"&i))
			else
				for j=1 to cint(request.form("action_number"&i))
					if request.Form("action_value"&i&"_"&j)<>"" then
						if request.Form("valid_value"&i)<>"" and request.Form("action_boolean"&i&"_"&j)<>"1" then
							response.Redirect("/Functions/ValueError.asp?erroractions="&request.Form("action_name"&i))
						end if
						if request.Form("action_value"&i&"_"&j)<>"NULL" and request.Form("action_value"&i&"_"&j)<>"0" then
							ACTION_VALUE=ACTION_VALUE&request.Form("action_value"&i&"_"&j)&","
						end if
					end if
				next
				if ACTION_VALUE<>"" then
					ACTION_VALUE=left(ACTION_VALUE,len(ACTION_VALUE)-1)
				end if
				rs("ACTION_VALUE")=ucase(ACTION_VALUE)
			end if
			rs.update
		next
	end if
	rs.close
end if
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=3;
adCount=0;
function showPopup()
{
adCount+=1
	if(adCount<timePopup)
	{
	setTimeout("showPopup()",1000);
	document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
	else
	{
	closePopup()
	}
}
function closePopup()
{
	location.href="Station_Close.asp"
}

function submitonce(theform)
{
	if (document.all||document.getElementById)
	{
		for (i=0;i<document.form1.length;i++)
		{
			var tempobj=document.form1.elements[i]
			if(tempobj.type.toLowerCase() =="submit"||tempobj.type.toLowerCase()=="reset")
			tempobj.disabled=true
		}
	}
}
</script>
</head>

<body onLoad="showPopup();"  bgcolor="#339966">
<form action="/Production/Sisonic/Station5.asp" method="post" name="form1" target="_self" onSubmit="return submitonce()">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2"><div align="center" class="strongred">System will close window in 3 seconds.&nbsp;<br>
    系统将在3秒钟后关闭窗口。
      <span id="countinsert"></span></div></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-DarkBlue">Station is in processing 本站正在进行中</td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Operator 操作员:
      <% =session("operator") %>
(
<% =session("code") %>
)</td>
  </tr>
  <tr>
    <td width="48%" height="20">Job Number 工单号</td>
    <td width="52%" height="20"><% =session("JOB_NUMBER") %></td>
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
    <td height="20"><% =session("LINE_NAME") %></td>
  </tr>
  <tr>
    <td height="20">Job Start Time 工单开始时间</td>
    <td height="20"><% =job_start_time%></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Station Progress 本站进程</td>
    </tr>
  <tr>
    <td height="20">Station Name 本站名</td>
    <td height="20"><span class="red"><%=getStationName(true,current_station_id,chinesename)%><%=chinesename%>&nbsp;<% =session("REPEATED_SEQUENCE")%></span></td>
  </tr>
  <tr>
    <td height="20">Station Start Time 本站开始时间</td>
    <td height="20"><% =station_start_time%>      </td>
    </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->
<%=session.Abandon()%>