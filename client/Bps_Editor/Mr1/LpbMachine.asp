<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("请登录后操作","../user_login.asp")
if session("ShiftEditType") = "OEM" then response.redirect "../InputOem.asp"
%>
<!--#include file= "conn.asp"-->
<!--#include virtual= "/BpsEditor/include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>LPB</title>
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
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
a{/* 统一设置所以样式 */
     font-family:Arial;
     font-size:12px;
     text-align:center;
     margin:1px;
}
a:link,a:visited{  /* 超链接正常状态、被访问过的样式 */
    color:#A62020;
    padding:4px 10px 4px 10px;
    background-color:#ecd8bd;
    text-decoration:none;
    border-top:1px solid #EEEEEE; /* 边框实现阴影效果 */
    border-left:1px solid #EEEEEE;
    border-bottom:1px solid #717171;
    border-right:1px solid #717171;
}
a:hover{       /* 鼠标指针经过时的超链接 */
    color:#821818;     /* 改变文字颜色 */
    padding:1px 8px 1px 8px;  /* 改变文字位置 */
    background-color:#e2c4c9;  /* 改变背景色 */
    border-top:1px solid #717171; /* 边框变换，实现“按下去”的效果 */
    border-left:1px solid #717171;
    border-bottom:1px solid #EEEEEE;
    border-right:1px solid #EEEEEE;
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
	if(document.form1.ShiftName.value.length==0){
     alert("班次不能为空!");
     document.form1.ShiftName.focus();
     return false;
    }
	if(document.form1.Ddate.value.length==0){
     alert("日期不能为空!");
     document.form1.Ddate.focus();
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
<OBJECT id="WebBrowser" height="0" width="0" classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2" VIEWASTEXT> </OBJECT>
<%
	ProductName = request("ProductName")
	LineName = request("LineName")
	DName = request("DName")
	ShiftName = request("ShiftName")
	Ddate = request("Ddate")
	
	Bid = request("Bid")
	Modi = request("Modi")
	AddMachine = request("AddMachine")
	MdId = request("MdId")
	'LineNo = ProductName & LineName
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

	'取Cookies缓存数据'
	if ProductName = "" then
		ProductName = request.Cookies("Lpb")("ProductName")
	else
		response.Cookies("Lpb")("ProductName") = ProductName
	end if
	
	if LineName = "" then
		LineName = request.Cookies("Lpb")("LineName")
	else
		response.Cookies("Lpb")("LineName") = LineName
	end if
	
	if DName = "" then
		DName = request.Cookies("Lpb")("DName")
	else
		response.Cookies("Lpb")("DName") = DName
	end if
	
	if ShiftName = "" then
		ShiftName = request.Cookies("Lpb")("ShiftName")
	else
		response.Cookies("Lpb")("ShiftName") = ShiftName
	end if
	
	if Ddate = "" then
		Ddate = request.Cookies("Lpb")("Ddate")
	else
		response.Cookies("Lpb")("Ddate") = Ddate
	end if
	'取Cookies缓存数据'
	
	if DName = "" then
		if 8 <= hour(now()) and hour(now()) <= 20 then
			DName = "D"
			response.Cookies("Lpb")("DName") = "D"
		else
			DName = "N"
			response.Cookies("Lpb")("DName") = "N"
		end if
	end if
	
	if Ddate = "" then
		Ddate  = FormatTime(date(),2)
		response.Cookies("Lpb")("Ddate") = Ddate
	end if
	
	if DName = "D" then
		Stime = cdate(Ddate & " 08:15:00")
		Etime = dateadd("h",12,Stime)
	else
		Stime = cdate(Ddate & " 20:15:00")
		Etime = dateadd("h",12,Stime)
	end if

	'增加设备维护信息'
	if AddMachine = "增加" then
		if ProductName = ""  or LineName = "" or DName = "" or Ddate = "" or ShiftName = "" or XlType = "" or Dhtime = "" or Detime = "" or Dequipment = "" or DPosition = "" or Dproblem = "" or DReason = "" or Dsolution = "" or Dtime = "" or (NOT IsNumeric(Dtime)) then
			call errorHistoryBack("有空数据或停机时间不为空,请重新填写!")
		else
			Sqlstr = "INSERT INTO [dbo].[Day_Data_Machine](Ddate,Udate,ProductName,LineName,DName,ShiftName,Dlossn,Dhtime,Detime,Dequipment,DPosition,Dproblem,DReason,Dsolution,Dtime,Dqty)"
			Sqlstr = Sqlstr & " VALUES('"&Ddate&"',getdate(),'"&ProductName&"','"&LineName&"','"&DName&"','"&ShiftName&"','"&Xltype&"','"&Dhtime&"','"&Detime&"','"&Dequipment&"','"&DPosition&"','"&Dproblem&"','"&DReason&"','"&Dsolution&"','"&Dtime&"','0')"
			SqlConn.execute(Sqlstr)
		end if
		call sussLoctionHref("添加成功","LpbMachine.asp")
	end if
	'增加设备维护信息'
	
	'增加主载体信息'
	if AddMachine = "增 加" then
		if ProductName = ""  or LineName = "" or DName = "" or Ddate = "" or ShiftName = ""  then
			call errorHistoryBack("有空数据,请重新填写!")
		else
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT top 1 * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and Ddate = '"&Ddate&"' and DName = '"&DName&"' and ShiftName = '"&ShiftName&"'"
			rs.open sql,SqlConn,1,1
			if rs.eof then
				Sqlstr = "INSERT INTO [dbo].[Day_Data_PC](Ddate,Udate,ProductName,LineName,DName,ShiftName,PlanStop,PcSum,PcQty)"
				Sqlstr = Sqlstr & " VALUES('"&Ddate&"',getdate(),'"&ProductName&"','"&LineName&"','"&DName&"','"&ShiftName&"','"&PlanStop&"','"&PcSum&"','"&PcQty&"')"
			'response.write Sqlstr
			'response.end
				SqlConn.execute(Sqlstr)
			end if
			rs.close
			set rs = nothing
		end if
		call sussLoctionHref("添加成功","LpbMachine.asp")
	end if
	'增加主载体信息'
	
	'修改设备维护信息'
	if AddMachine = "修改" then
		Sqlstr = "update [dbo].[Day_Data_Machine] set Dlossn='"&Xltype&"',Dhtime='"&Dhtime&"',Detime='"&Detime&"',Dequipment='"&Dequipment&"',DPosition='"&DPosition&"',Dproblem='"&Dproblem&"',DReason='"&DReason&"',Dsolution='"&Dsolution&"',Dtime='"&Dtime&"' where Rid = '"&MdId&"'"
		SqlConn.execute(Sqlstr)
		call sussLoctionHref("修改成功","LpbMachine.asp")
	end if
	'修改设备维护信息'
	
	'修改主载体信息'
	if AddMachine = "修 改" then
		Sqlstr = "update [dbo].[Day_Data_PC] set PcSum='"&PcSum&"',PcQty='"&PcQty&"',PlanStop='"&PlanStop&"' where Ddate = '"&Ddate&"' and DName = '"&DName&"' and ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and ShiftName = '"&ShiftName&"'"
		SqlConn.execute(Sqlstr)
		call sussLoctionHref("修改成功","LpbMachine.asp")
	end if
	'修改主载体信息'
	
	'提取Reamrk'
	if AddMachine = "remark" then
		dim RemarkStr,RemarkStrSum
		RemarkStr = ""
		RemarkStrSum = ""
		Dlineno = ProductName & LineName
		
			set rs = Server.CreateObject("adodb.recordset")
			sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND ShiftName = '"&ShiftName&"' and Ddate = '"&Ddate&"'"
			rs.open sql,SqlConn,1,1
			if not rs.eof then
				PcSum = rs("PcSum")
				PcQty = rs("PcQty")
				PlanStop = rs("PlanStop")
			end if
			rs.close
			set rs = nothing
		
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND ShiftName = '"&ShiftName&"' and Ddate = '"&Ddate&"' order by Rid"
		rs.open sql,SqlConn,1,1
		if not rs.eof then
			RemarkStrSum = Ddate & "　" & ProductName & " " & LineName & "　Shift: " & rs("ShiftName") & "　D/N: " & rs("DName") & " </BR></BR>"
			RemarkStrSum = RemarkStrSum & "计划停机: " & PlanStop & " 小时</br></br>"
			do while not rs.eof
				ShiftName = rs("ShiftName")
				Dlossn = rs("Dlossn")
				Ddate = FormatTime(rs("Ddate"),2)
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
			BQTY = 0
			RQTY = 0
			set rsq = Server.CreateObject("adodb.recordset")
			SqlStr = "SELECT TOP 1 * FROM [dbo].[Day_Data_QABlock] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND ShiftName = '"&ShiftName&"' and Ddate = '"&Ddate&"'"
			rsq.open SqlStr,SqlConn,1,1
			if not rsq.eof then
				BQTY = rsq("Block")
				RQTY = rsq("Release")
			end if
			rsq.close
			set rsq = nothing
			
			if BQTY <> 0 then RemarkStrSum = RemarkStrSum & "封存数量: " & BQTY & "</br></br>"
			if RQTY <> 0 then RemarkStrSum = RemarkStrSum & "放行数量: " & RQTY & "</br></br>"
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
</br>
<form name="form1" method="post" action="">
	<table width="1020" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>生产中心信息:</td>
			<td>产品：<select name="ProductName" id="TmProduct"></select></td>
			<td>线别：<select name="LineName" id="TmLine"></select></td>
			<select name="CellName" id="TmCell" style="display:none" ></select></td>
			<script type="text/javascript">
				MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=LineName%>','1');
			</script>
			<td>D/N：<select name="DName" >
						<option value="D" <%if DName = "D" then response.write "selected"%> >D</option>
						<option value="N" <%if DName = "N" then response.write "selected"%> >N</option>
					</select>
			</td>
			<td>Shift:<select name="ShiftName" >
							<option value="A" <%if ShiftName = "A" then response.write "selected"%> >A</option>
							<option value="B" <%if ShiftName = "B" then response.write "selected"%> >B</option>
							<option value="C" <%if ShiftName = "C" then response.write "selected"%> >C</option>
						</select>
			</td>
			<td>日期<input type="text" name="Ddate" id="Ddate" onChange="form1.submit();" value="<%=Ddate%>" size="6"/><img onClick="WdatePicker({el:'Ddate'})" src="../include/DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">
			</td>
			<td>
				<input type="submit" name="AddNew" value="提交">
			</td>
		</tr>
	</table>
</br>
	<table cellpadding="1" border="1" cellspacing="0" id="table" width="1020">

		<tr>
			<th>维护项目</th>
			<th>起止时间</th>
			<th>停机时间</th>
			<th>设备</th>
			<th>工位</th>
			<th>故障描述</th>
			<th>故障原因</th>
			<th>解决方法</th>
			<th>操作</th>
		</tr>
<%	
	dim SumDtime
	SumDtime = 0
	Dlineno = ProductName & LineName
	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM [dbo].[Day_Data_Machine] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND ShiftName = '"&ShiftName&"' and Ddate = '"&Ddate&"' order by Rid"

	rs.open sql,SqlConn,1,1
	do while not rs.eof
	Rid = rs("Rid")
	Dlossn = rs("Dlossn")
	DlossName = XlName(right(Dlossn,1))
	Ddate = FormatTime(rs("Ddate"),2)
	Dhtime = rs("Dhtime")
	Detime = rs("Detime")
	Dtime = rs("Dtime")
	Dproblem = rs("Dproblem")
	DReason = rs("DReason")
	Dsolution = rs("Dsolution")
	SumDtime = SumDtime + Dtime
%>
	<tr>
		<td><%=DlossName%></td>
		<td><%=Dhtime%> - <%=Detime%></td>
		<td><%=Dtime%></td>
		<td><%=rs("Dequipment")%></td>
		<td><%=rs("DPosition")%></td>
		<td><%=Dproblem%>&nbsp</td>
		<td><%=DReason%>&nbsp</td>
		<td><%=Dsolution%>&nbsp</td>
		<td>
		<%if (cdate(Stime) < Now()) and (Now() < cdate(Etime)) then %>
		<a href="LpbMachine.asp?Bid=<%=Rid%>&Modi=y" >修改</a>
		<%else
			response.write FormatTime(rs("Udate"),7)
		end if%>
		</td>
	</tr>
<%
	rs.movenext
	loop
	rs.close
	set rs = nothing
'------------------------------------------------'	

	set rs = Server.CreateObject("adodb.recordset")
	sql = "SELECT * FROM [dbo].[Day_Data_PC] where ProductName = '"&ProductName&"' and LineName = '"&LineName&"' and DName = '"&DName&"' AND ShiftName = '"&ShiftName&"' and Ddate = '"&Ddate&"' order by Rid"
	rs.open sql,SqlConn,1,1
	PcSum = 0
	PcQty = 0
	PlanStop = 0
	if not rs.eof then
		PCRid = rs("Rid")
		PcSum = rs("PcSum")
		PcQty = rs("PcQty")
		PlanStop = rs("PlanStop")
		PcButton = "修 改"
	else
		PcButton = "增 加"
	end if
	rs.close
	set rs = nothing
	
	if Modi = "y" then
		set rs = Server.CreateObject("adodb.recordset")
		sql = "SELECT top 1 * FROM [dbo].[Day_Data_Machine] where Rid = '"&Bid&"'"
		response.write Sql
		rs.open sql,SqlConn,1,1
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
		AddMachineValue = "<input type='Submit' name='AddMachine' value='修改' onClick='return beforeSubmit()'/>"
	else
		AddMachineValue = "<input type='Submit' name='AddMachine' value='增加' onClick='return beforeSubmit()'/>"
	end if

%>		

		<%if (cdate(Stime) < Now()) and (Now() < cdate(Etime)) then %>
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
			<input id="d11" type="text" onFocus="var d12=$dp.$('d12');WdatePicker({onpicked:function(){d12.focus();},maxDate:'#F{$dp.$D(\'d12\')}',skin:'whyGreen',dateFmt:'HH:mm:ss'})" name="Dhtime" size="4"  value="<%=MDhtime%>" />
			-
			<input id="d12" type="text" onFocus="WdatePicker({minDate:'#F{$dp.$D(\'d11\')}',skin:'whyGreen',dateFmt:'HH:mm:ss'})" name="Detime" size="4"  value="<%=MDetime%>"  onBlur="dis();" /></td>
			<td><input id="Dtime" type="text" name="Dtime" size="1" value="<%=MDtime%>" />Min</td>
			<td width="4">
				<select name="Dequipment" />
					<option value="WIMA">Wima</option>
					<option value="Cell1">Cell1</option>
					<option value="Cell2">Cell2</option>
					<option value="Cell3">Cell3</option>
					<option value="Cell4">Cell4</option>
					<option value="Cell21">Cell5</option>
					<option value="Zima">Zima</option>
				</select>
			</td>
			<td width="4"><input id="DPos1" type="text" name="DPosition" size="4" value="<%=MDPosition%>" /></td>
			<td><input id="DPro1" type="text" name="Dproblem" value="<%=MDproblem%>" /></td>
			<td><input id="DRes1" type="text" name="DReason" value="<%=MDReason%>" /></td>
			<td><input id="DSol1" type="text" name="Dsolution" value="<%=MDsolution%>" /></td>
			<td><%=AddMachineValue%></td>
		</tr>
		<%end if%>
				<tr>
			<td  colspan="9">
			</br>
			&nbsp;合计停机时间:&nbsp;<%=SumDtime%>&nbsp;分钟
			&nbsp;计划停机时间：<input type="text" name="PlanStop" style="text-align:center;" size="1" value="<%=PlanStop%>" />小时
			&nbsp;在线载体总数：<input type="text" name="PcSum" style="text-align:center;" size="1" value="<%=PcSum%>" />
			&nbsp;待修载体数量：<input type="text" name="PcQty" style="text-align:center;" size="1" value="<%=PcQty%>" />
			<%if (cdate(Stime) < Now()) and (Now() < cdate(Etime)) then %>
			&nbsp;<input type="Submit" name="AddMachine" value="<%=PcButton%>" />
			<%end if%>
			</td>
		</tr>
		<tr>
		<td  colspan="9"><input type="submit" name="AddMachine" value="remark"></td>
	</tr>

	<tr>
		<td colspan="9" align="center">
				<%if session("ShiftEditRole") = "0" then%><input type="button" value="设置Target" onClick="window.location.reload('SetHv.asp');" />　　　　<%end if%>
				<input type="button" value="修改密码" onClick="window.location.reload('../modify_pass.asp');" />　　　　
				<input type="button" value="产出数据" onClick="window.location.reload('../InputHV.asp');" />　　　　
				<input type="button" value="外围产出" onClick="window.location.reload('../');" />　　　　
				<input type="button" value="Close关闭" onClick="window.location.reload('../user_logout.asp');">
			</td>
	</tr>
	</table>
	</form>
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
