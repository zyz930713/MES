<%
response.charset = "UTF-8"
if session("ShiftEditUser") = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
if session("ShiftEditType") <> "HVPER" then response.redirect "default.asp"
ReadonlyStr = "readonly"
if session("ShiftEditRole") = "0" then ReadonlyStr = ""
%>
<!--#include file= "conn_open.asp"-->
<!--#include file= "include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>HV - Editor - Periphery</title>
<link rel="stylesheet" type="text/css" href="Styles/basic.css">
<script language="javascript" type="text/javascript" src="../include/ThreeMenuMa.js"></script>
<script language="javascript" type="text/javascript" src="include/DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
function $(obj){
   return document.getElementById(obj);
}
function Dtjs(Mano){
	var ot = 'O' + Mano + 'T';
	var st = 'S' + Mano + 'T';
	var opt = parseFloat($(ot).value);
	t1 = 720 - opt;
	t1 = t1.toFixed(0);
	$(st).value=t1;
}
function Majs(Mano){
	var cg = 'Ma' + Mano + 'G';
	var cb = 'Ma' + Mano + 'B';
	var cf = 'Ma' + Mano + 'F';
	var t1g = parseFloat($(cg).value);
	var t1b = parseFloat($(cb).value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$(cf).value=t1;
}
</script>
</head>
<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
	background:#F2F2F2;
}
table tr td{
	text-align:center;
}
.MaFOR{
	text-align: right;
}
</style>
<body style="margin:20px;font-size:14px;background-color:#339966;">

<%
	ProductName = request("ProductName")
	LineName = request("LineName")
	MaName = request("MaName")
	Adate = request("Adate")
	DName = request("DName")
	ShiftName = request("ShiftName")

	SearchEdit = request("SearchEdit")
	MaAddModi = request("MaAddModi")
	MaPerRe = request("MaPerRe")
	
	if Adate = "" then
		Adate = FormatTime(date(),2)
	end if
	
	'添加修改各数据
	if MaAddModi = "添加" or MaAddModi = "修改" then
		Rid = request("Rid")
		MaNr = request("MaNr")
		Maoc = request("Maoc")
		OpenTime = request("OpenTime")
		StopTime = request("StopTime")
		MaGood = request("MaGood")
		MaBad = request("MaBad")
		
		if Maoc = "on" or MaGood <> 0 then
			Maoc = 1
		else
			Maoc = 0
		end if

		if MaAddModi = "添加" then
			InsertSql = "INSERT INTO [dbo].[Day_Data_Periphery]([Adate],[AName],[Udate],[UName],[ProductName],[DName],[MaName],[MaNr],[MaRun],[OpenTime],[StopTime],[MaGood],[MaBad])"
			InsertSql = InsertSql & " VALUES('"&cdate(Adate)&"',N'"&session("ShiftEditUser")&"',getdate(),N'"&session("ShiftEditUser")&"','"&ProductName&"','"&DName&"','"&MaName&"','"&MaNr&"','"&Maoc&"','"&OpenTime&"','"&StopTime&"','"&MaGood&"','"&MaBad&"')"
			'response.write InsertSql
			'response.end
			ConnSql.execute(InsertSql)
			call sussLoctionHref("添加成功","InputPeriphery.asp?SearchEdit=ok&ProductName="&ProductName&"&MaName="&MaName&"&Adate="&Adate&"&DName="&DName)
		elseif MaAddModi = "修改" then
			UpdateSql = "UPDATE [dbo].[Day_Data_Periphery] set Udate = getdate(),[UName] = N'"&session("ShiftEditUser")&"',MaRun = '"&Maoc&"',OpenTime = '"&OpenTime&"',StopTime = '"&StopTime&"'"
			UpdateSql = UpdateSql & ",MaGood = '"&MaGood&"',MaBad = '"&MaBad&"'"
			UpdateSql = UpdateSql & " WHERE Rid = '"&Rid&"'"
			'response.write UpdateSql
			' response.end
			ConnSql.execute(UpdateSql)
			call sussLoctionHref("修改成功","InputPeriphery.asp?SearchEdit=ok&ProductName="&ProductName&"&MaName="&MaName&"&Adate="&Adate&"&DName="&DName)
		end if
	end if
	'添加修改各数据
	
	'添加修改维护数据
	if MaPerRe = "写入" then
		neirong = request("neirong")
		neirong1 = Replace(neirong,vbCrlf, "<br/>" )
		InsertSql = "INSERT INTO [dbo].[Day_Data_MachinePer]([Adate],[Udate],[ProductName],[DName],[MaName],[Dsolution])"
		InsertSql = InsertSql & " VALUES('"&cdate(Adate)&"',getdate(),'"&ProductName&"','"&DName&"',N'"&MaName&"',N'"&neirong1&"')"
		'response.write InsertSql
		'response.end
		ConnSql.execute(InsertSql)
		call sussLoctionHref("添加成功","InputPeriphery.asp?SearchEdit=ok&ProductName="&ProductName&"&MaName="&MaName&"&Adate="&Adate&"&DName="&DName)
	end if
	
	'添加修改维护数据
	if MaPerRe = "修改" then
		neirong = request("neirong")
		ReID = request("ReID")
		neirong1 = Replace(neirong,vbCrlf, "<br/>" )
		UpdateSql = "UPDATE [dbo].[Day_Data_MachinePer] SET Udate = getdate(),[Dsolution] = N'"&neirong1&"'"
		UpdateSql = UpdateSql & " WHERE Rid = '"&ReID&"'"
		response.write UpdateSql
		response.end
		ConnSql.execute(UpdateSql)
		call sussLoctionHref("修改成功","InputPeriphery.asp?SearchEdit=ok&ProductName="&ProductName&"&MaName="&MaName&"&Adate="&Adate&"&DName="&DName)
	end if
%>
<center>
<form name="InputHv" method="post" action="InputPeriphery.asp" >
	<table width="800" border="1" cellspacing="0" cellpadding="0">
		<tr>
			<td class="TitleLable"><p style="height:34px;line-height:34px;"><strong>外围设备产出及维护</strong></p>
		</td>
	</tr>
		<tr>
			<td align="center">
				　产品：<select name="ProductName" id="TmProduct"></select>
				　设备：<select name="MaName" id="TmLine" ></select>
				<select name="CellName" id="TmCell" style="display:none" ></select>
					<script type="text/javascript">
						MenuInit('TmProduct','TmLine','TmCell','<%=ProductName%>','<%=MaName%>','1');
					</script>
				　日期：<input type="text" name="Adate" class="Wdate" id="Adate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('Adate_1').value=$dp.cal.getP('W','W');$dp.$('Adate_2').value=$dp.cal.getP('W','WW');}})" value="<%=Adate%>" size="10" />

				　D/N：<select name="DName" >
							<option value="D" <%if DName = "D" then response.write "selected"%> >Day</option>
							<option value="N" <%if DName = "N" then response.write "selected"%> >Night</option>
						</select>
				　<input type="Submit" name="SearchEdit" value="刷新" />
			</td>
		</tr>
	</table>
</form>

<%
if SearchEdit = "刷新" or SearchEdit = "ok" then
	MaSum = 1
	PcsTarget = 0
	FORTarget = 0
	set MaRs = Server.CreateObject("adodb.recordset")
	MaSql = "select top 1 [MaName],[MaSum],[TDate],[PcsTarget],[FORTarget] from Cfg_Periphery where ProductName = '"&ProductName&"' AND MaName = '"&MaName&"' and Tdate < = '"&Adate&"' ORDER BY Tdate desc"
	MaRs.open MaSql,ConnSql,1,1
	if not MaRs.eof then
		MaSum = MaRs("MaSum")
		PcsTarget = MaRs("PcsTarget")
		FORTarget = MaRs("FORTarget")
	end if
	'取出设备数量
%>
	<table width="800" border="1" cellspacing="0" cellpadding="0">
		<tr>
			<td colspan="8" align="center">
				<strong>产品：<%=ProductName%>　设备：<%=MaName%>　日期：<%=Adate%>　D/N：<%=DName%></strong>
			</td>
		</tr>
		<tr>
			<th width="10%">Machine</th>
			<th width="15%">Production</th>
			<th width="15%">Good pieces</th>
			<th width="15%">Bad picesc</th>
			<th width="10%">FOR</th>
			<th width="15%">Production Time</th>
			<th width="15%">Down Time</th>
			<th width="10%">操作</th>
		</tr>
		<%
		set MaRs = Server.CreateObject("adodb.recordset")
		for MaLoop = 1 to MaSum
			MaSql = "SELECT [Rid],[Adate],[Udate],[ProductName],[LineName],[DName],[ShiftName],[MaName],[MaNr],[MaRun],[OpenTime],[StopTime],isnull([MaGood],0) [MaGood],isnull([MaBad],0) [MaBad] FROM [dbo].[Day_Data_Periphery]"
			MaSql = MaSql & " where ProductName = '"&ProductName&"' AND MaName = '"&MaName&"' and MaNr = '"&MaLoop&"' AND [Adate] = '"&FormatTime(Adate,2)&"' AND DName = '"&DName&"'"
			'response.write MaSql
			Mars.open MaSql,ConnSql,1,1
			MachineClose = ""
			MaSubmit = "<input type='submit' name='MaAddModi' value='添加' style='background:#FFBFBF;' />"
			Rid = ""
			MaGood = 0
			MaBad = 0
			MaFOR = 0
			MaRun = 0
			MaOpenTime = 720
			MaStopTime = 0
			if not Mars.eof then
				Rid = MaRs("Rid")
				MaGood = Mars("MaGood")
				MaBad = Mars("MaBad")
				MaRun = Mars("MaRun")
				MaOpenTime = Mars("OpenTime")
				MaStopTime = Mars("StopTime")
				MaSubmit = "<input type='submit' name='MaAddModi' value='修改' style='background:#BFFFBF;' />"
			end if
			Mars.close
			MaAll = MaGood + MaBad
			if MaAll > 0 then MaFOR = Formatnumber(MaBad / MaAll * 100,2,-1)
			'response.write MaRun & "</br>"
			if MaRun = "True" then MachineClose = "checked"
			%>
			<form name="form<%=MaLoop%>" method="post" action="InputPeriphery.asp">
			<tr>
				<td align="center"><%=MaName%>s <%=MaLoop%>
					<input type="hidden" name="Rid" value="<%=Rid%>" />
					<input type="hidden" name="MaNr" value="<%=MaLoop%>" />
					<input type="hidden" name="Adate" value="<%=Adate%>" />
					<input type="hidden" name="ProductName" value="<%=ProductName%>" />
					<input type="hidden" name="MaName" value="<%=MaName%>" />
					<input type="hidden" name="DName" value="<%=DName%>" />
				</td>
				<td align="center"><input type="checkbox" name="Maoc" <%=MachineClose%> />Machine runs</td>
				<td align="center"><input type="text" name="MaGood" id="Ma<%=MaLoop%>G" onChange="Majs(<%=MaLoop%>)" value="<%=MaGood%>" size="5" /></td>
				<td align="center"><input type="text" name="MaBad" id="Ma<%=MaLoop%>B" onChange="Majs(<%=MaLoop%>)" value="<%=MaBad%>" size="5" /></td>
				<td align="center"><input type="text" name="MaFOR" id="Ma<%=MaLoop%>F" value="<%=MaFOR%>" size="4" class="MaFOR"/>%</td>
				<td align="center"><input type="text" name="OpenTime" id="O<%=MaLoop%>T" onChange="Dtjs(<%=MaLoop%>)" value="<%=MaOpenTime%>" size="5"/></td>
				<td align="center"><input type="text" name="StopTime" id="S<%=MaLoop%>T" value="<%=MaStopTime%>" size="5" /></td>
				<td align="center"><%=MaSubmit%>
			</tr>
			</form>
			<%
		next
		
		'取维护内容信息
		MaSql = ""
		MaSql = "SELECT [Rid],[Adate],[Udate],[ProductName],[MaName],[DName],[ShiftName],[Dsolution] FROM [dbo].[Day_Data_MachinePer]"
		MaSql = MaSql & " where Adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND DName = '"&DName&"' AND MaName = '"&MaName&"'"
		'response.write MaSql
		Mars.open MaSql,ConnSql,1,1
		MaPerContent = ""
		MaPerSubStr = "<input type='submit' name='MaPerRe' value='写入' style='background:#FFBFBF;'/>"
		if not MaRs.eof then
			ReID = MaRs("Rid")
			MaPerContent = MaRs("Dsolution")
			MaPerSubStr = "<input type='submit' name='MaPerRe' value='修改' style='background:#BFFFBF;'/>"
		end if
		set Mars = nothing
		%>
	<form name="FormMaPer" method="post" action="InputPeriphery.asp">
		<tr>
			<TD align="right" width="120" height="200">维护内容</TD>
			<td colspan="7" >
				<input type="hidden" name="ReID" value="<%=ReID%>" />
				<input type="hidden" name="Adate" value="<%=Adate%>" />
				<input type="hidden" name="ProductName" value="<%=ProductName%>" />
				<input type="hidden" name="MaName" value="<%=MaName%>" />
				<input type="hidden" name="DName" value="<%=DName%>" />
				<textarea name="neirong" style="height:100%;width:100%;" ><%=Replace(MaPerContent, "<br/>",vbCrlf)%></textarea>
			</td>
		</tr>
		<tr>
			<td colspan="8" align="center" height="15" colspan="2">
				<%=MaPerSubStr%>
			</td>
		</tr> 
	</form>
	</table>

<%END IF%>
</br>
	<table width="800" border="1" cellspacing="0" cellpadding="0">
		<tr>
			<td colspan="5" align="center">
				<%if session("ShiftEditRole") = "0" then%><input type="button" value="设置Target" onClick="window.location.reload('SetHvPeriphery.asp');" />　　　　<%end if%>
				<input type="button" value="修改密码" onClick="window.location.reload('modify_pass.asp');" />　　　　
				<input type="button" value="Close关闭" onClick="window.location.reload('user_logout.asp');">
			</td>
		</tr>
	</table>
</center>

</body> 
</html>