<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
session.Abandon()
pagename="Station1.asp"
errorstring=request.QueryString("errorstring")
cookiesExpiresDate = Date+3650 '10 years

'set computerName to  cookies
computerName=request("computername")
if computerName<>"" then
	Response.Cookies("computer_name")=computerName
	Response.Cookies("computer_name").Expires=cookiesExpiresDate
else
	computerName = Request.Cookies("computer_name")
end if
'get current station from cookies
currentStation = request("radStationId")
if currentStation = "" then
	currentStation = Request.Cookies("current_station")
	if currentStation = "" then
		response.Redirect("Station.asp")
	end if
else
	aryStation = split(currentStation,"$")
	if Ubound(aryStation)=1 then
		Response.Cookies("current_station_id") = aryStation(0)
		Response.Cookies("current_station") = aryStation(1)
		Response.Cookies("current_station_id").Expires=cookiesExpiresDate
		Response.Cookies("current_station").Expires=cookiesExpiresDate
		currentStation = aryStation(1)
	else
		response.Redirect("Station.asp")
	end if
end if
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->


<% set rsdistrictName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(computerName)+"'"
	
		rsdistrictName.open SQL,conn,1,3
	 	
		if rsdistrictName.eof  then
	
		word="此电脑不是BPS系统专用电脑！"
		action="location.href='station.asp'"
		%>
		
		<script language="javascript">
        alert("<%=word%>");
        <%=action%>;
        </script>
		
		
	<%
		else
		
		
		if isnull(rsdistrictName("LINE_NAME")) then
		%>
		<script language="javascript">
       alert("线别设置不正确！");
        location.href="station.asp";
        </script>
		
	<%	end if
	    end if
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="javascript">
function pageload()
{
	document.form1.code.focus();
}
function formcheck()
{
	document.all.errorinsertcode.innerText="";
	document.all.errorinsertjob.innerText="";
	if(document.form1.code.value==""){
		document.all.errorinsertcode.innerText="Operator Code cannot be blank!\n操作员工号不能为空！";
		document.form1.code.focus();
		result=false;
	}else if(document.form1.jobnumber.value==""){
		document.all.errorinsertjob.innerText="Job Number cannot be blank!\n工单号不能为空！";
		document.form1.jobnumber.focus();
		result=false;
	}else{
		document.form1.Next.disabled=true;
		document.form1.submit();
	}
}

function ShowJobQueue()
{
	var jobnumber=document.getElementById("jobnumber").value;
	window.showModalDialog("JobQueue.asp?jobnumber="+jobnumber,window,"dialogHeight:800px;dialogWidth:800px");
}

function getJobByTray()
{
	if(document.form1.trayId.value){
		var jobNumber = window.showModalDialog("GetValueByKey.asp?key=TrayId&keyValue="+document.form1.trayId.value,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		if(jobNumber.indexOf("Error-")==0){			
			document.form1.jobnumber.value = "";
			alert(jobNumber.substring(6));
			document.form1.trayId.blur();
			document.form1.trayId.select();
			return false;
		}else{
			document.form1.jobnumber.value = jobNumber;
		}
	}
}

function getJobByComputer()
{
	if(document.form1.code.value && !document.form1.trayId.value){
		var jobNumber = window.showModalDialog("GetValueByKey.asp?key=computer&keyValue=<%=Request.Cookies("computer_name")%>",window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		if(jobNumber.indexOf("Error-")==0){			
			document.form1.jobnumber.value = "";
			alert(jobNumber.substring(6));
			document.form1.trayId.blur();
			document.form1.trayId.select();
			return false;
		}else{
			document.form1.jobnumber.value = jobNumber;
		}
	}
}

function chgFocus(obj,objNext){	
	if(event.keyCode==13 && obj.value){	
		//event.keyCode=9;//Key: Tab
		objNext.focus();
	}
}

</script>
</head>

<body onLoad="pageload()" onKeyPress="keyhandler();" bgcolor="#339966">

<form action="Station2.asp" method="post" name="form1" target="_self">
<table width="90%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2" class="t-t-DarkBlue">Scan job info 扫描工单信息
	&nbsp;&nbsp;<a style="color:#6495ED;" href="Station.asp"><%=currentStation%></a>
	</td>
  </tr>
	<%if errorstring<>"" then%>
	  <tr>
		<td height="20" colspan="2" class="strongred"><div align="center"><%=errorstring%>&nbsp;</div></td>
		</tr>
	  <tr>
  	<%end if%>
	<td width="31%">Tray ID 料盘号</td>
    <td width="69%">
      <input name="trayId" type="text" id="trayId" onKeyPress="chgFocus(this,form1.code);" tabindex="1" size="20" autocomplete="off" onChange="getJobByTray()">
	  </td>
  </tr>
    <td width="31%"><span class="d_link">O</span>perator Code 工号 <span class="red">*</span> </td>
    <td width="69%">
      <input name="code" type="text" id="code" onChange="getJobByComputer();" onKeyPress="chgFocus(this,form1.jobnumber);" tabindex="2" size="20"  autocomplete="off">
	 <br><span id="errorinsertcode" class="strongred"></span></td>
  </tr>
  <tr>
    <td height="20"><span class="d_link">J</span>ob Number  工单号<span class="red"> *</span></td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" onFocus="this.select();" tabindex="3" onKeyPress="chgFocus(this,form1.Next);" size="20" autocomplete="off" >
	<br><span id="errorinsertjob" class="strongred"></span>	
	</td>
  </tr>
  <%if request.QueryString("shift")="Y" then%>
  <tr>
    <td height="20">Shift In 开线</td>
    <td height="20"><input name="shiftin" type="checkbox" id="shiftin" value="1" checked>
      确认开线</td>
  </tr>
  <%end if%>
  <tr>
    <td rowspan="2" class="red">Locked Items 锁定项目 </td>
    <td height="20">Following code, machine or lot is locked, please don't use them<br>
      下列工号、机器或批号已被锁定，请不要使用：</td>
  </tr>
  <tr>
    <td height="20">
	<%	SQL="select LOCKED_LOT from MATERIAL where LOCKED_LOT is not null"
		rs.open SQL,conn,1,3
		while not rs.eof
			lockalert=lockalert&rs("LOCKED_LOT")&"<br>"
			rs.movenext
		wend
		rs.close
	%>
	<%=lockalert%>&nbsp;
	</td>
  </tr>
  <tr>
    <td height="20" colspan="2">&nbsp;</td>
    </tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
        <input name="Next" type="button" accesskey="n" tabindex="5" value="Next 下一步" onClick="formcheck()">
        &nbsp;&nbsp;
        <input name="Reset" type="reset" accesskey="r" tabindex="6" value="Reset 重置">
        &nbsp;
        <input name="Close" type="button" id="Close" tabindex="7" onClick="javascript:window.close()" value="Close 关闭"> 
		&nbsp;	 
		<input name="ShiftOut" type="button" id="Pause" tabindex="8" value="Shift Out 停线" onClick="javascript:document.form1.action='/KES1/Station_ShiftOut.asp';document.form1.submit()">		
    </div></td>
  </tr>    
</table>
<br>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">  
  <tr>
    <td colspan="2" class="darkbule">支持电话：18910736955 (24小时) </td>
    </tr>
  <tr>
    <td><div align="left"><span class="darkbule style1">IP Address: <%=request.ServerVariables("REMOTE_HOST")%></span></div></td>
    <td><div align="right"><span class="darkbule style1">Host Name: <%=request.ServerVariables("HTTP_HOST")%></span></div></td>
  </tr>
  <tr><td>&nbsp;</td></tr>

  <tr>
  <td height="20" colspan="2"></td>
  </tr>

</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->