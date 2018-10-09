<%
response.charset = "UTF-8"
ShiftEditUser = session("ShiftEditUser")
if ShiftEditUser = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
if session("ShiftEditType") <> "HV" then response.redirect "default.asp"
%>
<!--#include file= "conn_open.asp"-->
<!--#include file= "include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>HV - Editor - Line</title>
<link rel="stylesheet" type="text/css" href="Styles/basic.css">
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
</head>
<style type="text/css">
table{
border:1px solid #000;
}
table tr{
	height:26px;
	line-height:26px;
	background:#F2F2F2;
}
table tr td{
	text-align:center;
}

</style>
<script type="text/javascript">
function beforeSubmit(form)
{	
	if(document.form1.ProductName.value.length==0){
     alert("产品不能为空!");
     document.form1.ProductName.focus();
     return false;
    }
	if(document.form1.LineName.value.length==0){
     alert("产品不能为空!");
     document.form1.LineName.focus();
     return false;
    }
	if(document.form1.DName.value.length==0){
     alert("D/N 不能为空!");
     document.form1.DName.focus();
     return false;
    }
	if(document.form1.Adate.value.length==0){
     alert("日期不能为空!");
     document.form1.Adate.focus();
     return false;
    }
	if(document.form1.d11.value.length==0){
     alert("开始时间不能为空!");
     document.form1.d11.focus();
     return false;
    }
	if(document.form1.Detime.value.length==0){
     alert("终止时间不能为空!");
     document.form1.Detime.focus();
     return false;
    }
	if(document.form1.Dtime.value.length==0){
     alert("停机时间不能为空!");
     document.form1.Dtime.focus();
     return false;
    }
	if(isNaN(document.form1.Dtime.value)){
     alert("停机时间必须为数字!");
     document.form1.Dtime.focus();
     return false;
    }
	if(document.form1.Dequipment.value.length==0){
     alert("设备不能为空!");
     document.form1.Dequipment.focus();
     return false;
    }
	if(document.form1.DPos1.value.length==0){
     alert("工位不能为空!");
     document.form1.DPos1.focus();
     return false;
    }
	if(document.form1.DPro1.value.length==0){
     alert("故障描述不能为空!");
     document.form1.DPro1.focus();
     return false;
    }
	if(document.form1.DRes1.value.length==0){
     alert("故障原因不能为空!");
     document.form1.DRes1.focus();
     return false;
    }
	if(document.form1.DSol1.value.length==0){
     alert("解决方法不能为空!");
     document.form1.DSol1.focus();
     return false;
    }
return true;
}
function dis(){
	var t2=$('d12').value;
	var arr2=t2.split(":");
	var t1=$('d11').value;
	var arr1=t1.split(":");
	var arr3=new Array();
for(var i=0;i<arr2.length;i++){
	arr3[i]= parseInt(arr2[i])- parseInt(arr1[i]);
}
if(i==2){
	var result=arr3[0]+":"+arr3[1];
}
else if(i==3){
	var result=arr3[0]+":"+arr3[1]+":"+arr3[2];
	var dd13=arr3[0]*60+arr3[1];
}
	$('Dtime').value=dd13;
}
function $(obj){
   return document.getElementById(obj);
}
function cycode(obj)
{
	var rng = document.body.createTextRange();
	rng.moveToElementText(obj);
	rng.select();
	rng.execCommand('Copy');
}
</script>
<body style="margin:20px;font-size:14px;background-color:#339966;">
<%
	ProductName = request("ProductName")
	if ProductName = "" then ProductName = ShiftEditUser
	
	LineName = request("LineName")
	Adate = request("Adate")
	DName = request("DName")
	ShiftName = request("ShiftName")
	SearchEdit = request("SearchEdit")

	Bid = request("Bid")
	Modi = request("Modi")
	MachineSub = request("MachineSub")
	PCSub = request("PCSub")

	MdId = request("MdId")
	XlType = request("XlType")
	Dhtime=request("Dhtime")
	Detime=request("Detime")
	Dequipment=request("Dequipment")
	DPosition=request("DPosition")
	Dproblem=request("Dproblem")
	DReason=request("DReason")
	Dsolution=request("Dsolution")
	Dtime=request("Dtime")
	PcSum = request("PcSum")
	PcQty = request("PcQty")
	PlanStop = request("PlanStop")

	if Adate = "" then
		Adate  = FormatTime(date(),2)
	end if
	
	TestSub = request("TestSub")
	OnlineId = request("OnlineId")
	QtySM = request("QtySM")
	QtyOC = request("QtyOC")
	QtyTS = request("QtyTS")
	QtyMT = request("QtyMT")
	QABlock = request("QABlock")

	if Modi = "d" then
		ConnSql.execute("delete FROM [dbo].[Day_Data_Machine] where Rid = '"&Bid&"'")
		SearchEdit = "刷新"
	end if

	'增加线上数据信息'
	if TestSub = "增加" then
		if ProductName = ""  or LineName = "" or DName = "" or Adate = "" then
			call errorHistoryBack("有空数据或停机时间不为空,请重新填写!")
		else
			Sqlstr = "INSERT INTO [dbo].[Day_Data_Online](Adate,Udate,ProductName,LineName,DName,[QtySM],[QtyOC],[QtyTS],[QtyMT],[QABlock])"
			Sqlstr = Sqlstr & " VALUES('"&Adate&"',getdate(),'"&ProductName&"','"&LineName&"','"&DName&"','"&QtySM&"','"&QtyOC&"','"&QtyTS&"','"&QtyMT&"','"&QABlock&"')"
			ConnSql.execute(Sqlstr)
		end if
		call sussLoctionHref("添加成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'增加线上数据信息'
	'修改线上数据信息
	if TestSub = "修改" then
		Sqlstr = "update [dbo].[Day_Data_Online] set Udate=getdate(),QtySM='"&QtySM&"',QtyOC='"&QtyOC&"',QtyTS='"&QtyTS&"',QtyMT='"&QtyMT&"',QABlock='"&QABlock&"' where Rid = '"&OnlineId&"'"
		ConnSql.execute(Sqlstr)
		call sussLoctionHref("修改成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'修改线上数据信息'
	
	'增加设备维护信息'
	if MachineSub = "增加" then
		if ProductName = ""  or LineName = "" or DName = "" or Adate = "" or XlType = "" or Dhtime = "" or Detime = "" or Dequipment = "" or DPosition = "" or Dproblem = "" or DReason = "" or Dsolution = "" or Dtime = "" or (NOT IsNumeric(Dtime)) then
			call errorHistoryBack("有空数据或停机时间不为空,请重新填写!")
		else
			Sqlstr = "INSERT INTO [dbo].[Day_Data_Machine](Adate,Udate,ProductName,LineName,DName,Dlossn,Dhtime,Detime,Dequipment,DPosition,Dproblem,DReason,Dsolution,Dtime,Dqty)"
			Sqlstr = Sqlstr & " VALUES('"&Adate&"',getdate(),'"&ProductName&"','"&LineName&"','"&DName&"','"&Xltype&"','"&Dhtime&"','"&Detime&"',N'"&Dequipment&"',N'"&DPosition&"',N'"&Dproblem&"',N'"&DReason&"',N'"&Dsolution&"','"&Dtime&"','0')"
			ConnSql.execute(Sqlstr)
		end if
		call sussLoctionHref("添加成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'增加设备维护信息'
	'修改设备维护信息'
	if MachineSub = "修改" then
		Sqlstr = "update [dbo].[Day_Data_Machine] set Dlossn='"&Xltype&"',Dhtime='"&Dhtime&"',Detime='"&Detime&"',Dequipment=N'"&Dequipment&"',DPosition=N'"&DPosition&"',Dproblem=N'"&Dproblem&"',DReason=N'"&DReason&"',Dsolution=N'"&Dsolution&"',Dtime='"&Dtime&"' where Rid = '"&MdId&"'"
		ConnSql.execute(Sqlstr)
		call sussLoctionHref("修改成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'修改设备维护信息'
	
	'增加主载体信息'
	if PCSub = "增加" then
		if ProductName = "" or LineName = "" or DName = "" or Adate = "" then
			call errorHistoryBack("有空数据,请重新填写!")
		else
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT top 1 * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and Adate = '"&Adate&"' and DName = '"&DName&"'"
			rs.open sql,ConnSql,1,1
			if rs.eof then
				Sqlstr = "INSERT INTO [dbo].[Day_Data_PC](Adate,Udate,ProductName,LineName,DName,PlanStop,PcSum,PcQty)"
				Sqlstr = Sqlstr & " VALUES('"&Adate&"',getdate(),'"&ProductName&"','"&LineName&"','"&DName&"','"&PlanStop&"','"&PcSum&"','"&PcQty&"')"
			'response.write Sqlstr
			'response.end
				ConnSql.execute(Sqlstr)
			end if
			rs.close
			set rs = nothing
		end if
		call sussLoctionHref("添加成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'增加主载体信息'
	'修改主载体信息'
	if PCSub = "修改" then
		Sqlstr = "update [dbo].[Day_Data_PC] set PcSum='"&PcSum&"',PcQty='"&PcQty&"',PlanStop='"&PlanStop&"' where Adate = '"&Adate&"' and DName = '"&DName&"' and ProductName = '"&ProductName&"' and LineName = '"&LineName&"'"
		ConnSql.execute(Sqlstr)
		call sussLoctionHref("修改成功","HvMachine.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName)
	end if
	'修改主载体信息'
	
'提取Reamrk'
	if SearchEdit = "remark" then
		dim RemarkStr,RemarkStrSum
		RemarkStr = ""
		RemarkStrSum = ""
		Dlineno = ProductName & LineName
		
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND Adate = '"&Adate&"'"
			rs.open sql,ConnSql,1,1
			if not rs.eof then
				PcSum = rs("PcSum")
				PcQty = rs("PcQty")
				PlanStop = rs("PlanStop")
			end if
			rs.close
			set rs = nothing
		
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND Adate = '"&Adate&"' order by Rid"
		rs.open sql,ConnSql,1,1
		if not rs.eof then
			RemarkStrSum = Adate & "　" & ProductName & " " & LineName & "　D/N: " & rs("DName") & " </BR></BR>"
			RemarkStrSum = RemarkStrSum & "计划停机: " & PlanStop & " 小时</br></br>"
			do while not rs.eof
				Dlossn = rs("Dlossn")
				Adate = FormatTime(rs("Adate"),2)
				Dhtime = rs("Dhtime")
				Detime = rs("Detime")
				Dtime = rs("Dtime")
				Dequipment = rs("Dequipment")
				DPosition = rs("DPosition")
				Dproblem = rs("Dproblem")
				DReason = rs("DReason")
				Dsolution = rs("Dsolution")
					
					RemarkStr = Dhtime & " - " & Detime & "　停机: " & Dtime & " 分钟</br>"
					RemarkStr = RemarkStr & "位置:" & Dequipment & " " & DPosition & "</br>"
					RemarkStr = RemarkStr & "问题:" & Dproblem & "</br>"
					RemarkStr = RemarkStr & "原因:" & DReason & "</br>"
					RemarkStr = RemarkStr & "解决:" & Dsolution & "</br>"
					RemarkStr = RemarkStr & "</br>"
			
				RemarkStrSum = RemarkStrSum & RemarkStr
			rs.movenext
			loop

			'提取封存数据'
			set RsOnline = Server.CreateObject("adodb.recordset")
			sql = "SELECT [QtySM],[QtyOC] ,[QtyTS],[QtyMT],[QABlock],[QARelease] FROM [dbo].[Day_Data_Online] where ProductName = '"&ProductName&"' and LineName = '"&LINENAME&"' and DName = '"&DNAME&"' and Adate = '"&Adate&"' order by Rid"
			RsOnline.open sql,ConnSql,1,1
				QtySM = 0
				QtyOC = 0
				QtyTS = 0
				QtyMT = 0
				QABlock = 0
				'QARelease = 0
			if not RsOnline.eof then
				QtySM = RsOnline("QtySM")
				QtyOC = RsOnline("QtyOC")
				QtyTS = RsOnline("QtyTS")
				QtyMT = RsOnline("QtyMT")
				QABlock = RsOnline("QABlock")
				'QARelease = RsOnline("QARelease")
				RemarkStrSum = RemarkStrSum & "采样: " & QtySM & "</br>"
				RemarkStrSum = RemarkStrSum & "目测废品: " & QtyOC & "</br>"
				RemarkStrSum = RemarkStrSum & "清零废品: " & QtyTS & "</br>"
				RemarkStrSum = RemarkStrSum & "废膜/顶片: " & QtyMT & "</br>"
				RemarkStrSum = RemarkStrSum & "封存: " & QABlock & "</br></br>"
			end if
			RsOnline.close
			set RsOnline = nothing
			'提取封存数据'

			RemarkStrSum = RemarkStrSum & "在线主载体数量: " & PcSum & "　待修主载体数量: " & PcQty & "</br>"
			rs.close
			set rs = nothing

		end if

		response.write "<input type='button' value='返回' onClick='history.back(-1)' />　"
		response.write "<input type=button value='复制到剪切板' onclick='cycode(ip);'>"
		response.write "<div id='ip' style='background:#F1F1F1;'>" & RemarkStrSum & "</div>"
		response.end
	end if
	'提取Reamrk'
%>
<center>
<form name="form1" method="post" action="HvMachine.asp">
<table width="1160" border="1" cellspacing="0" cellpadding="0"><tr><td align="center" class="TitleLable"><p style="height:34px;line-height:34px;"><strong>线上设备维护及测试数据</strong></p></td></tr>
		<tr>
			<td align="center">　产品：<select name="ProductName" id="TmProduct"></select>
				　线别：<select name="LineName" id="TmLine"></select>
				<select name="CellName" id="TmCell" style="display:none" ></select>
					<script type="text/javascript">
						MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
					</script>
				　日期：<input type="text" name="Adate" class="Wdate" id="Adate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('Adate_1').value=$dp.cal.getP('W','W');$dp.$('Adate_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />
				　D/N：<select name="DName" >
							<option value="D" <%if DName = "D" then response.write "selected"%> >Day</option>
							<option value="N" <%if DName = "N" then response.write "selected"%> >Night</option>
						</select>
				<input type="submit" name="SearchEdit" value="刷新">
			</td>
		</tr>
	</table>
</br>
<%
if SearchEdit = "刷新" or SearchEdit = "ok" or Modi = "y" then
set rs = Server.CreateObject("adodb.recordset")
sql = "SELECT [Rid],[QtySM],[QtyOC] ,[QtyTS],[QtyMT],[QABlock],[QARelease] FROM [dbo].[Day_Data_Online] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Rid"
rs.open sql,ConnSql,1,1
	QtySM = 0
    QtyOC = 0
    QtyTS = 0
    QtyMT = 0
    QABlock = 0
    QARelease = 0
	TestSubStr = "<input type='submit' name='TestSub' value='增加' style='background:#FFBFBF;'/>"
	if not rs.eof then
		OnlineId = rs("Rid")
		QtySM = rs("QtySM")
		QtyOC = rs("QtyOC")
		QtyTS = rs("QtyTS")
		QtyMT = rs("QtyMT")
		QABlock = rs("QABlock")
		QARelease = rs("QARelease")
		TestSubStr = "<input type='submit' name='TestSub' value='修改' style='background:#BFFFBF;'/>"
	end if
rs.close
%>
	<table cellpadding="1" border="1" cellspacing="0" id="table" width="1160">
		<tr>
			<td colspan="9">
				<strong>产品：<%=ProductName%>　线别：<%=LineName%>　日期：<%=Adate%>　D/N：<%=DName%></strong>
			</td>
		</tr>
		<tr style="background:#F7FFFF;">
		<th>在线数据:</th>
	        <th colspan="8" style="text-align:left;">
				<input type="hidden" name="OnlineId" value="<%=OnlineId%>" />
				采样：<input type="text" name="QtySM" value="<%=QtySM%>" size="4" />　
				目测废品：<input type="text" name="QtyOC" value="<%=QtyOC%>" size="4" />　
				清零废品：<input type="text" name="QtyTS" value="<%=QtyTS%>" size="4" />　
				废膜/顶片：<input type="text" name="QtyMT" value="<%=QtyMT%>" size="4" />　
				封存：<input type="text" name="QABlock" value="<%=QABlock%>" size="4" />　
				<%=TestSubStr%>
			</th>
		</tr>
		<tr><td colspan="9" height="5"></td></tr>
		<tr>
			<th width="130">维护项目</th>
			<th width="140">起止时间</th>
			<th WIDTH="50">时长</th>
			<th WIDTH="60">设备</th>
			<th width="90">工位</th>
			<th width="180">故障描述</th>
			<th width="180">故障原因</th>
			<th width="180">解决方法</th>
			<th width="60">操作</th>
		</tr>
<%
	dim SumDtime
	SumDtime = 0
	Dlineno = ProductName & LineName
	
	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Rid"
	rs.open sql,ConnSql,1,1
			do while not rs.eof
			Rid = rs("Rid")
			Dlossn = rs("Dlossn")
			DlossName = XlName(right(Dlossn,1))
			Adate = FormatTime(rs("Adate"),2)
			Dequipment = rs("Dequipment")
			DPosition = rs("DPosition")
			Dhtime = rs("Dhtime")
			Detime = rs("Detime")
			Dtime = rs("Dtime")
			Dproblem = rs("Dproblem")
			DReason = rs("DReason")
			Dsolution = rs("Dsolution")
			SumDtime = SumDtime + Dtime
		%>
			<tr>
				<td><%=DlossName%><a href="HvMachine.asp?Bid=<%=Rid%>&ProductName=<%=ProductName%>&LineName=<%=LineName%>&Adate=<%=Adate%>&DName=<%=DName%>&Modi=d" ><img src="images/del.jpg" alt="删除记录" border="0" id="DelSub"/></a></td>
				<td><%=Dhtime%> - <%=Detime%></td>
				<td><%=Dtime%></td>
				<td><%=Dequipment%></td>
				<td><%=DPosition%></td>
				<td><%=Dproblem%></td>
				<td><%=DReason%></td>
				<td><%=Dsolution%></td>
				<td><a href="HvMachine.asp?Bid=<%=Rid%>&ProductName=<%=ProductName%>&LineName=<%=LineName%>&Adate=<%=Adate%>&DName=<%=DName%>&Modi=y" >修改</a></td>
			</tr>
		<%
	rs.movenext
	loop
	rs.close
	set rs = nothing
'------------------------------------------------'	

	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' and Adate = '"&Adate&"' order by Rid"
	rs.open sql,ConnSql,1,1
	PcSum = 0
	PcQty = 0
	PlanStop = 0
	if not rs.eof then
		PCRid = rs("Rid")
		PcSum = rs("PcSum")
		PcQty = rs("PcQty")
		PlanStop = rs("PlanStop")
		PcButton = "<input type='Submit' name='PCSub' value='修改' style='background:#BFFFBF;'/>"
	else
		PcButton = "<input type='Submit' name='PCSub' value='增加' style='background:#FFBFBF;'/>"
	end if
	rs.close
	set rs = nothing
	
	if Modi = "y" then
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT top 1 * FROM [dbo].[Day_Data_Machine] where Rid = '"&Bid&"'"
		'response.write Sql
		rs.open sql,ConnSql,1,1
		if not rs.eof then
			Bid = rs("Rid")
			MDlossn = rs("Dlossn")
			MDhtime = rs("Dhtime")
			MDetime = rs("Detime")
			MDequipment = rs("Dequipment")
			MDPosition = rs("DPosition")
			MDproblem = rs("Dproblem")
			MDReason = rs("DReason")
			MDsolution = rs("Dsolution")
			MDtime = rs("Dtime")
		end if
		rs.close
		set rs = nothing
		AddMachineValue = "<input type='Submit' name='MachineSub' value='修改' style='background:#BFFFBF;' onClick='return beforeSubmit()'/>"
	else
		AddMachineValue = "<input type='Submit' name='MachineSub' value='增加' style='background:#FFBFBF;' onClick='return beforeSubmit()'/>"
	end if

%>
		<tr>
			<td>
			<input type="hidden" id="BjId" name="MdId" value="<%=Bid%>" size="1" readonly>
				<select name="XlType">
					<option value="XL1" <%if MDlossn = "XL1" then response.write "selected" %> >更换零备件</option>
					<option value="XL2" <%if MDlossn = "XL2" then response.write "selected" %> >维修</option>
					<option value="XL3" <%if MDlossn = "XL3" then response.write "selected" %> >设备调整</option>
					<option value="XL4" <%if MDlossn = "XL4" then response.write "selected" %> >计划性检测</option>
					<option value="XL5" <%if MDlossn = "XL5" then response.write "selected" %> >单位加工时间</option>
					<option value="XL6" <%if MDlossn = "XL6" then response.write "selected" %> >速度不匹配</option>
					<option value="XL7" <%if MDlossn = "XL7" then response.write "selected" %> >小停顿</option>
					<option value="XL8" <%if MDlossn = "XL8" then response.write "selected" %> >过程废品</option>
					<option value="XL9" <%if MDlossn = "XL9" then response.write "selected" %> >初始不良品</option>
				</select>
			</td>
			<td>
			<input id="d11" type="text" onFocus="var d12=$dp.$('d12');WdatePicker({onpicked:function(){d12.focus();},skin:'whyGreen',dateFmt:'HH:mm:ss'})" name="Dhtime" size="4" value="<%=MDhtime%>" />-<input id="d12" type="text" onFocus="WdatePicker({skin:'whyGreen',dateFmt:'HH:mm:ss'})" name="Detime" size="4"  value="<%=MDetime%>" onBlur="dis();"/></td>
			<td><input id="Dtime" type="text" name="Dtime" size="1" value="<%=MDtime%>" /></td>
			<td width="4">
				<select name="Dequipment" />
					<option value="" <%if MDequipment = "WIMA" then response.write "selected" %> ></option>
					<option value="WIMA" <%if MDequipment = "WIMA" then response.write "selected" %> >Wima</option>
					<option value="Cell1" <%if MDequipment = "Cell1" then response.write "selected" %> >Cell1</option>
					<option value="Cell2" <%if MDequipment = "Cell2" then response.write "selected" %> >Cell2</option>
					<option value="Cell3" <%if MDequipment = "Cell3" then response.write "selected" %> >Cell3</option>
					<option value="Cell4" <%if MDequipment = "Cell4" then response.write "selected" %> >Cell4</option>
					<option value="Cell5" <%if MDequipment = "Cell5" then response.write "selected" %> >Cell5</option>
					<option value="Zima" <%if MDequipment = "Zima" then response.write "selected" %> >Zima</option>
				</select>
			</td>
			<td width="4"><input id="DPos1" type="text" name="DPosition" value="<%=MDPosition%>" size="10"/></td>
			<td><input id="DPro1" type="text" name="Dproblem" value="<%=MDproblem%>" size="26"/></td>
			<td><input id="DRes1" type="text" name="DReason" value="<%=MDReason%>" size="26"/></td>
			<td><input id="DSol1" type="text" name="Dsolution" value="<%=MDsolution%>" size="26"/></td>
			<td><%=AddMachineValue%></td>
		</tr>
		<tr>
			<td colspan="9">
				&nbsp;合计停机时间:&nbsp;<font color="red"><%=SumDtime%></font>&nbsp;分钟
				&nbsp;计划停机时间：<input type="text" name="PlanStop" style="text-align:center;" size="1" value="<%=PlanStop%>" />小时
				&nbsp;在线载体总数：<input type="text" name="PcSum" style="text-align:center;" size="1" value="<%=PcSum%>" />
				&nbsp;待修载体数量：<input type="text" name="PcQty" style="text-align:center;" size="1" value="<%=PcQty%>" />
				&nbsp;<%=PcButton%>
			</td>
		</tr>
		<tr><td colspan="9"><input type="submit" name="SearchEdit" value="remark" /></td></tr>
</table>
</form>
<%end if%>
<table width="1160" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="9" align="center">
				<input type="button" value="修改密码" onClick="window.location.reload('modify_pass.asp');" />　　　　
				<input type="button" value="产出数据" onClick="window.location.reload('InputHV.asp');" />　　　　
				<input type="button" value="Close关闭" onClick="window.location.reload('user_logout.asp');">
			</td>
	</tr>
</table>
</center>

<%
function XlName(Xltype)
	select case Xltype
		case 1
			XlName = "更换零备件"
		case 2
			XlName = "维修"
		case 3
			XlName = "设备调整"
		case 4
			XlName = "计划性检测"
		case 5
			XlName = "单位加工时间"
		case 6
			XlName = "速度不匹配"
		case 7
			XlName = "小停顿"
		case 8
			XlName = "过程废品"
		case 9
			XlName = "初始不良品"
	end select
end function
%>
</body>
</html>