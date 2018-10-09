<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/KES1/SessionCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetStationName.asp" -->
<%
pagename="Station6.asp"
station_start_time=request("station_start_time")
station_close_time=now()
islocked=false
lockalert=""
str2DCode=""
strDisplay2DCodeInfo=""
job2DCodeQty=0
jobGoodQty=0
has2DCode=false

'is machine, lot, code locked
for i=1 to cint(request.Form("action_count"))
	select case	request.form("action_purpose"&i)
		case "3" 'Material Lot Number
		SQL="select LOCKED_LOT from MATERIAL where LOCKED_LOT is not null"
		rs.open SQL,conn,1,3
		if not rs.eof then
		while not rs.eof
			for j=1 to cint(request.form("action_number"&i))
				if instr(rs("LOCKED_LOT"),trim(request.Form("action_value"&i&"_"&j)))<>0 then
				islocked=true
				lockalert=lockalert&trim(request.Form("action_value"&i&"_"&j))&"<br>"
				end if
			next
		rs.movenext
		wend
		end if
		rs.close
	end select
next

if islocked=true then
response.Redirect("ActionError.asp?alerttype=lock&alertmessage="&lockalert)
end if

current_station_id=request.Form("current_station_id")
session("current_station_id")=current_station_id

'if user fresh this page, before re-save actions' info, all old records need to be deleted.
SQL="delete from JOB_ACTIONS where STATION_POSITION=1 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"'"
if session("STATIONS_ROUTINE")="1" then
	SQL=SQL+" and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
end if
rs.open SQL,conn,3,3

'save all actions' info.
SQL="select * from JOB_ACTIONS where STATION_POSITION=1 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
rs.open SQL,conn,3,3
if rs.eof then
	for i=1 to cint(request.Form("action_count"))
		ACTION_VALUE=""		
		'Job quantity or Rework quantity
		if request.form("action_purpose"&i)="5" and request.form("action_purpose"&i)="6" then
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
		end if
		if request.form("action_purpose"&i)="7" then
			has2DCode=true
		end if
		
		if request.Form("action_purpose"&i)="8" then
		hasDefect2DActn=true
		end if
		
		
		
		if ACTION_VALUE<>"" then
			'2D code action
			
			if hasDefect2DActn then
			strDefect2DCode=strDefect2DCode& ucase(ACTION_VALUE) &","
			end if
			
			if request.form("action_purpose"&i)="7"  then
				strDisplay2DCodeInfo=strDisplay2DCodeInfo & "<tr><td>" & request.Form("action_name"&i) &"</td><td colspan='3'>"&ACTION_VALUE&"</td></tr>"
				str2DCode = str2DCode& ucase(ACTION_VALUE) &","				
			else
				rs.addnew
				rs("JOB_NUMBER")=session("JOB_NUMBER")
				rs("SHEET_NUMBER")=session("SHEET_NUMBER")
				rs("JOB_TYPE")=session("JOB_TYPE")
				rs("STATION_ID")=current_station_id
				rs("ACTION_ID")=request.Form("action_id"&i)
				rs("STATION_POSITION")=1
				rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")		
				rs("ACTION_VALUE")=ucase(ACTION_VALUE)
				rs.update
			end if
		end if
	next
end if
rs.close


'save Defect2code

if hasDefect2DActn then	

  defect_code_id=request("defect_code0")
	set rsDefect2DCode = server.createobject("adodb.recordset")
	if strDefect2DCode <>"" and strDefect2DCode <> session("2DCodeInfo") then 'for refresh this page
		'session("2DCodeInfo") = strDefect2DCode
		aryDefect2DCode = split(strDefect2DCode,",")
		
		
		for i=0 to Ubound(aryDefect2DCode)-1
		    
			if aryDefect2DCode(i) <>"" then
			keyValue=cstr(aryDefect2DCode(i))
	        strSql = "select defect_code_id,lm_user,lm_time from job_2d_code where code = '"& trim(keyValue) & "' and job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' "
	        
	        rsDefect2DCode.open strSql,conn,1,3
	        if not rsDefect2DCode.eof then
		    	if rsDefect2DCode("defect_code_id") <> "" then
			
		   		else
				
				rsDefect2DCode("defect_code_id")=defect_code_id
				rsDefect2DCode("lm_user")=session("code")
				rsDefect2DCode("lm_time")=now()
				rsDefect2DCode.update
		    	end if		
			end if
			rsDefect2DCode.close
			end if
		next
		
	end if
	
end if








'save 2D code info
if has2DCode then	
	set rs2DCode = server.createobject("adodb.recordset")
	if str2DCode <>"" and str2DCode <> session("2DCodeInfo") then 'for refresh this page
		session("2DCodeInfo") = str2DCode
		ary2DCode = split(str2DCode,",")
		lmTime=now()
		sql2DCode="select code,job_number,sheet_number,lm_user,lm_time from job_2d_code where 1=2"		
		rs2DCode.open sql2DCode,conn,1,3
		for i=0 to Ubound(ary2DCode)-1
			if ary2DCode(i) <>"" then
				rs2DCode.addnew
				rs2DCode("code")=cstr(ary2DCode(i))
				rs2DCode("job_number")=session("JOB_NUMBER")
				rs2DCode("sheet_number")=session("SHEET_NUMBER")
				rs2DCode("lm_user")=session("code")
				rs2DCode("lm_time")=lmTime
				rs2DCode.update
			end if
		next
		rs2DCode.close
	end if
	'get 2D code count of job
	sql2DCode="select count(1) from job_2d_code where job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"'"
	rs2DCode.open sql2DCode,conn,1,3
	job2DCodeQty = clng(rs2DCode(0))
	rs2DCode.close	
end if

'if user fresh this page, before re-save actions' info, all old records need to be deleted.
if request.Form("station_default_transaction_type")="1" then
	SQL="delete from JOB_ACTIONS_REPEATED where STATION_POSITION=1 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
	rs.open SQL,conn,3,3
	'save all actions' info into JOB_ACTIONS_REPEATED, if station's transaction_type is optional.
	SQL="select * from JOB_ACTIONS_REPEATED where STATION_POSITION=1 and JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"' and STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
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
			'Job quantity or Rework quantity
			if request.form("action_purpose"&i)="5" and request.form("action_purpose"&i)="6" then
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
			end if
			rs("ACTION_VALUE")=ucase(ACTION_VALUE)
			rs.update
		next
	end if
	rs.close
end if

'get job's properties.
SQL="select START_TIME,JOB_GOOD_QUANTITY,SHIFT from JOB where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and JOB_TYPE='"&session("JOB_TYPE")&"'"
rs.open SQL,conn,1,3
	job_start_time=rs("START_TIME")
	jobGoodQty=clng(rs("JOB_GOOD_QUANTITY"))
	SHIFT=trim(rs("SHIFT"))
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
</head>

<body onLoad="pageload()" onKeyPress="keyhandler()"  bgcolor="#339966">
<form action="/KES1/Station7.asp" method="post" name="form1" target="_self" onsubmit="return formSubmit()">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="4" class="t-t-DarkBlue">Confirm station info 确认本站信息</td>
  </tr>
  <tr>
    <td height="20" colspan="4" class="t-t-Borrow">Operator 操作员:
      <% =session("operator") %>
(
<% =session("code") %>
)</td>
  </tr>
  <tr>
    <td height="20">Job Number 工单号</td>
    <td height="20"><% =session("JOB_NUMBER") %></td>
    <td>Sheet Number 分批号 </td>
    <td><% =session("SHEET_NUMBER") %></td>
  </tr>
  
  <tr>
    <td height="20">Part Number  型号 </td>
    <td height="20"><% =session("PART_NUMBER_TAG") %></td>
    <td>Line Name 线别</td>
    <td><% =session("LINE_NAME") %></td>
  </tr>
  
  <tr>
    <td height="20">Job Start Time  工单开始时间</td>
    <td height="20" colspan="3"><%=job_start_time%>&nbsp;</td>
    </tr>
  <tr>
    <td height="20" colspan="4" class="t-t-Borrow">Station Summary 本站小结</td>
    </tr>
  <tr>
    <td height="20">Station Name 本站站名</td>
    <td height="20"><span class="red"><%=getStationName(true,current_station_id,chinesename)%><%=chinesename%>&nbsp;<% =session("REPEATED_SEQUENCE") %></span></td>
    <td>Station Start Time 开始时间</td>
    <td><%=station_start_time%></td>
  </tr>
  
  <tr>
    <td height="20">Station Stop Time 结束时间</td>
    <td height="20"><%=station_close_time%>&nbsp;</td>
    <td>Passed Time 用时</td>
    <td><%=datediff("n",station_start_time,station_close_time)%> minutes&nbsp; <input id="Cminutes"  type="hidden" value="<%=datediff("n",station_start_time,station_close_time)%>" ></td>
  </tr>
  
  <%
  jobQty=0
  SQL="select ACTIONS_INDEX from STATION where NID='"&current_station_id&"'"
  rs.open SQL,conn,1,3
  if not rs.eof then
	  ACTIONS_INDEX=rs("ACTIONS_INDEX")  
	  if not isnull(ACTIONS_INDEX) then
		  aaction=split(ACTIONS_INDEX,",")
	  end if
  end if
  rs.close
  if not isnull(ACTIONS_INDEX) then
  	SQL="select 1,A.ACTION_ID,A.JOB_NUMBER,I.ACTION_NAME,I.ACTION_CHINESE_NAME,I.ACTION_PURPOSE,A.ACTION_VALUE from JOB_ACTIONS A inner join ACTION I on A.ACTION_ID=I.NID where A.JOB_NUMBER='"&session("JOB_NUMBER")&"' and A.SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and A.JOB_TYPE='"&session("JOB_TYPE")&"' and A.STATION_ID='"&current_station_id&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
  	rs.open SQL,conn,1,3
  	if not rs.eof then
  	  for j=0 to ubound(aaction)
		while not rs.eof
		  if rs("ACTION_ID")=aaction(j) then
  		%>
		  <tr>
			<td height="20"><% =rs("ACTION_NAME")%>(<%=rs("ACTION_CHINESE_NAME")%>)</td>
			<td height="20" colspan="3">
			<%
			if rs("ACTION_PURPOSE")="5" then
			  	jobQty=clng(rs("ACTION_VALUE"))
			end if
			  			
			if rs("ACTION_PURPOSE")="3" then
				if not isnull(rs("ACTION_VALUE")) then
				%>
				<%=replace(rs("ACTION_VALUE"),","," ; ")%>
				<%end if
			  elseif rs("ACTION_PURPOSE")="4" then
				if not isnull(rs("ACTION_VALUE")) then%>
					<%=replace(rs("ACTION_VALUE"),","," + ")%>
				<%end if
			  else%>
				<% =rs("ACTION_VALUE")%>
			<%end if%>&nbsp;</td>
			</tr>
		  <%end if
		  rs.movenext
		wend
	  rs.movefirst
	  next
    end if
    rs.close
  end if
  
  'display 2D code info
  if strDisplay2DCodeInfo <> "" then
  	response.Write(strDisplay2DCodeInfo)
  end if
  
  defectQty=0
  'save defect data
  set rsDeftemp = server.createobject("adodb.recordset")
  for i=0 to cint(request("defect_count"))
		deftQty=request("defect_quantity"&i)	
		SQL="select JOB_NUMBER,SHEET_NUMBER,STATION_ID,DEFECT_CODE_ID,REPEATED_SEQUENCE,DEFECT_QUANTITY from JOB_DEFECTCODES "
		SQL=SQL+"where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and STATION_ID='"&request("station_id"&i)&"' and DEFECT_CODE_ID='"&request("defect_code"&i)&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
		rs.open SQL,conn,1,3
		
		'save data to defect code temp
		strDefSQL="select JOB_NUMBER,SHEET_NUMBER,STATION_ID,DEFECT_CODE_ID,REPEATED_SEQUENCE,DEFECT_QUANTITY,ENTER_STATION_ID from JOB_DEFECTCODES_TEMP "
		strDefSQL=strDefSQL+"where JOB_NUMBER='"&session("JOB_NUMBER")&"' and SHEET_NUMBER='"&session("SHEET_NUMBER")&"' and STATION_ID='"&request("station_id"&i)&"' "
		strDefSQL=strDefSQL+" and DEFECT_CODE_ID='"&request("defect_code"&i)&"' and REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"' and ENTER_STATION_ID='"&current_station_id&"'"
		rsDeftemp.open strDefSQL,conn,1,3
		
		if deftQty<>"" and deftQty<>"0" then
			if rs.eof then
				rs.addnew
				rs("JOB_NUMBER")=session("JOB_NUMBER")
				rs("SHEET_NUMBER")=session("SHEET_NUMBER")
				rs("STATION_ID")=request("station_id"&i)
				rs("DEFECT_CODE_ID")=request("defect_code"&i)
				rs("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
				rs("DEFECT_QUANTITY")=deftQty
			else 
				rs("DEFECT_QUANTITY")=clng(rs("DEFECT_QUANTITY"))+clng(deftQty)
			end if			
			rs.update
			
			if rsDeftemp.eof then
				rsDeftemp.addnew
				rsDeftemp("JOB_NUMBER")=session("JOB_NUMBER")
				rsDeftemp("SHEET_NUMBER")=session("SHEET_NUMBER")
				rsDeftemp("STATION_ID")=request("station_id"&i)
				rsDeftemp("DEFECT_CODE_ID")=request("defect_code"&i)
				rsDeftemp("REPEATED_SEQUENCE")=session("REPEATED_SEQUENCE")
				rsDeftemp("DEFECT_QUANTITY")=deftQty
				rsDeftemp("ENTER_STATION_ID")=current_station_id
			else 
				rsDeftemp("DEFECT_QUANTITY")=clng(rsDeftemp("DEFECT_QUANTITY"))+clng(deftQty)
			end if					
			rsDeftemp.update
			
		elseif not rs.eof then'输入数量为0，则删除该不良
			'rs.delete
		end if
		rs.close
		rsDeftemp.close
  next
  
  'get defect info
  SQL="SELECT A.DEFECT_CODE_ID,A.STATION_ID,A.DEFECT_QUANTITY,B.DEFECT_CODE,B.DEFECT_NAME||'('||B.DEFECT_CHINESE_NAME||')' AS DEFECT_NAME FROM JOB_DEFECTCODES_TEMP A INNER JOIN DEFECTCODE B ON A.DEFECT_CODE_ID = B.NID "
  SQL=SQL+"WHERE A.JOB_NUMBER='"&session("JOB_NUMBER")&"' AND A.SHEET_NUMBER='"&session("SHEET_NUMBER")&"' AND A.ENTER_STATION_ID='"&current_station_id&"' AND A.REPEATED_SEQUENCE='"&session("REPEATED_SEQUENCE")&"'"
  rs.open SQL,conn,1,3

  if rs.recordcount >0 then
  	i=0
 %>
 <tr>
    <td height="20">Defect Code 缺陷代码</td>
	<td height="20" colspan="3">		
		<table id="tb_defect_code" border=1 cellSpacing=0 cellPadding=0 >
		<tr><td>Seq 序号</td><td align=center colspan="2">Defect Code 缺陷代码</td><td align=center>Quantity 数量</td></tr>
		<%while not rs.eof
			defectQty = defectQty + clng(rs("DEFECT_QUANTITY"))
		%>
  		<tr><td><%=(i+1)%></td><td><%=rs("DEFECT_CODE")%></td><td><%=rs("DEFECT_NAME")%></td><td><%=rs("DEFECT_QUANTITY")%><td>
    	<td>
			<input type="hidden" id="defect_code<%=i%>" name="defect_code<%=i%>" value="<%=rs("DEFECT_CODE_ID")%>">
			<input type="hidden" id="defect_quantity<%=i%>" name="defect_quantity<%=i%>" value="<%=rs("DEFECT_QUANTITY")%>">
			<input type="hidden" id="station_id<%=i%>" name="station_id<%=i%>" value="<%=rs("STATION_ID")%>">
		</td>
  		</tr>
  		<%rs.movenext
		  i=i+1
		wend
		rs.close%>
		</table>
	</td>
  </tr>
<%end if%>
  <tr>
    <td height="20" colspan="4" align="center">
	<input type="hidden" id="defect_count" name="defect_count" value="<%=i%>">
	&nbsp;</td>
  </tr>
  <tr>
    <td height="20" colspan="4"><div align="center">      
      <input name="station_close_time" type="hidden" id="station_close_time" value="<%=station_close_time%>">
      <%
	  set rsDQ = server.createobject("adodb.recordset")
sql="select * from JOB_DEFECTCODES_TEMP where Job_Number ='"&session("JOB_NUMBER")&"' and Sheet_Number='"&session("SHEET_NUMBER")&"' and  Enter_Station_Id in (select NID  from  station where mother_station_id='SA00000989' )"
rsDQ.open sql,conn ,1,3
if not  rsDQ.eof then
Defect_Quantity=rsDQ("Defect_Quantity")
end if
rsDQ.close

sql=" select * from (select NID  from  station where mother_station_id='SA00000989') where  NID='"&current_station_id&"'"
rsDQ.open sql ,conn,1,3
if not rsDQ.eof then
NID=rsDQ("NID")
end if
rsDQ.close
sql="select PRODUCT  from line where LINE_NAME='"&session("LINE_NAME")&"'"
rsDQ.open sql ,conn,1,3
if not rsDQ.eof then
PRODUCT=trim(rsDQ("PRODUCT"))
end if
rsDQ.close
sql="select get_sub_job_number('"&session("JOB_NUMBER") &"','"&session("SHEET_NUMBER") &"') as job_number from dual"

rsDQ.open sql ,conn,1,3
if not rsDQ.eof then
jobN=rsDQ("job_number")
end if
rsDQ.close
	  %>
      <input name="Next" type="submit" id="Next" value="OK 工单结束" >
  &nbsp;
  <input type="button" name="btnBack" value="Back 继续录入" onClick="javascript:location.href='/KES1/Station1.asp'">
  <input type="hidden" name="Defect_Quantity" id="Defect_Quantity" value="<%=Defect_Quantity%>">
  <input type="hidden" name="NID" id="NID" value="<%=NID%>">
  <input type="hidden" name="PRODUCT" id="PRODUCT" value="<%=PRODUCT%>">
   <input type="hidden" name="JobN" id="jobN" value="<%=jobN%>">
    </div></td>
    </tr>
</table>
</form>
<div id="msg_2d_code" align="center" />
</body>
<script language="javascript">
function pageload()
{
	<%if has2DCode and jobGoodQty > (job2DCodeQty+defectQty) then
		remainQty = jobGoodQty - job2DCodeQty - defectQty
	%>	
		document.form1.Next.disabled=true;
		document.form1.btnBack.focus();
		document.all.msg_2d_code.innerHTML="<font color='red'>Please scan 2D codes that in other tray, 2D Code remain qty is <%=remainQty%><br>请继续扫描其他料盘的二维码,二维码剩余数量是<%=remainQty%></font>";
	<%elseif session("checkCurrentStaion")<> "" then%>
		document.form1.Next.disabled=true;
		document.form1.btnBack.focus();
		document.all.msg_2d_code.innerHTML="<font color='red'><%=session("checkCurrentStaion")%></font>";
	<%elseif defectQty> jobQty then%>
		document.form1.Next.disabled=true;
		document.form1.btnBack.focus();
		document.all.msg_2d_code.innerHTML="<font color='red'>Total Defect Code Quantity exceeds Job Quantity!<br>Defect Code数量合计超过Job Quantity！</font>";
	<%else%>
		document.form1.Next.focus();
	<%end if%>
}

function formSubmit()
{
	var Defect_Quantity=document.form1.Defect_Quantity.value; 
	var NID=document.form1.NID.value; 
	var Cminutes=document.form1.Cminutes.value;
	var JobN=document.form1.JobN.value;
	var PRODUCT=document.form1.PRODUCT.value;
	if  (Cminutes<1 )
	{
		alert("关站时间不能小于开站时间,请点击继续录入")
		 return false;
		
		
	}
	 if (NID !="") 
	{ 
	      if (PRODUCT=="RA" || PRODUCT=="Slim" ||PRODUCT=="Franklin"  ||PRODUCT=="Petra"  )
		  
		  
		  {
			 
			  var rtnValue = window.showModalDialog("GetValueByKey.asp?key=JobNumber&keyValue="+JobN,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			
				if(rtnValue != ""){
					
					if (confirm("确定已完成此工站了吗？OK -此工站关闭，不能再输入不良 NO -此工站不关闭，只记录不良!!!!"))
				   { 		   
					  
				  
					submitonce("form1");
					
					}
					else
					{
					
					  return false;
					}
					
					
					
				}
				else
				{
					alert("请添写报告后，再关站点");
					
					window.open("InputHV.asp?ProductName=<%=PRODUCT%>&LineName=<%=session("LINE_NAME")%>&JobNumber=<%=jobN%>&ShiftName=<%=SHIFT%>");
					
					return false;
					
					
					}
		
		  }
		  
		  else
		  
		  {
	 
				if  ( !Defect_Quantity>0  )
				
				{
					alert("声学站点不良不能为空，请输入声学不良");
					return false;
					
				}
				else	
				
				{
				
			
					if (confirm("确定已完成此工站了吗？OK -此工站关闭，不能再输入不良 NO -此工站不关闭，只记录不良"))
				   { 
				   
					  
				  
					submitonce("form1");
					
					}
					else
					{
					
					  return false;
					}
				}
		  }
	
	}
	else
	
	{
		
	if (confirm("确定已完成此工站了吗？OK -此工站关闭，不能再输入不良 NO -此工站不关闭，只记录不良"))
   { 
   
      
  
	submitonce("form1");
	
	}
	else
	{
	
	  return false;
	}
		
		
		
		
		}
}

</script>





</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->