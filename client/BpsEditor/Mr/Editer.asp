<!--#include file= "../ManagementReporting/conn.asp"-->
<!--#include file= "../include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>WH</title>
<script language="javascript" type="text/javascript" src="../include/ThreeMenuEdit.js"></script>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
function $(obj){
   return document.getElementById(obj);
}

function jst1f(){
	var t1g = parseFloat($('T1G').value);
	var t1b = parseFloat($('T1B').value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$('T1F').value=t1;
}
function jscf(cellno){
	var cg = 'C' + cellno + 'G';
	var cb = 'C' + cellno + 'B';
	var cf = 'C' + cellno + 'F';
	var t1g = parseFloat($(cg).value);
	var t1b = parseFloat($(cb).value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$(cf).value=t1;
}
function zimajs(zimano)
{
	var zg = 'Z' + zimano + 'G';
	var zb = 'Z' + zimano + 'B';
	var zf = 'Z' + zimano + 'F';
	var t1g = parseFloat($(zg).value);
	var t1b = parseFloat($(zb).value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$(zf).value=t1;
}
</script>
</head>
<style type="text/css">
#table{
border:1px solid #000;
}
#table tr td{
text-align:center;
}
</style>
<body>

<%
	AddWhdata = request("AddWhdata")
	SearchEdit = request("SearchEdit")
	ProductName = request("ProductName")
	LineName = request("LineName")
	Ddate = request("Ddate")
	DDN = request("DDN")
	
	CellAddModi = request("CellAddModi")
	ZimaAddModi = request("ZimaAddModi")
	
	if DDN = "" then
		if 8 <= hour(now()) and hour(now()) <= 20 then
			DDN = "T"
			response.Cookies("Lpb")("DDN") = "T"
		else
			DDN = "N"
			response.Cookies("Lpb")("DDN") = "N"
		end if
	end if
	
	if Ddate = "" then
		Ddate = FormatTime(date(),2)
	end if
	
	if DDN = "T" then
		Stime = cdate(Ddate & " 07:15:00")
		Etime = dateadd("h",12,Stime)
	else
		Stime = cdate(Ddate & " 19:15:00")
		Etime = dateadd("h",12,Stime)
	end if
	
	if ProductName = "RA" then
		ZimaCounts = 11
	elseif ProductName = "Petra" then
		ZimaCounts = 3
	end if
	
	' response.write ProductName
	' response.write "</br>"
	' response.write LineName
	' response.write "</br>"
	'response.write DDN
	
	if AddWhdata = "添加" then
		call sussLoctionHref("添加成功","Editer.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN)
	elseif AddWhdata = "修改" then
		'conn.execute("UPDATE [DWH].[dbo].[Report_Day_RA2] SET RealToWh = '"&realtowh&"' where [day] = '"&Ddate&"' AND ProdlineNr = '"&LineName&"' AND ShiftAcronym = '"&DDN&"'")
		call sussLoctionHref("修改成功","Editer.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN)
	end if
	
	if CellAddModi = "添加" or CellAddModi = "修改" then
		if CellAddModi = "添加" then
			response.write "Cell添加"
			call sussLoctionHref("添加成功","Editer.asp?SearchEdit=ok&Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN)
		elseif CellAddModi = "修改" then
			response.write "Cell修改"
			call sussLoctionHref("修改成功","Editer.asp?SearchEdit=ok&Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN)
		end if
	end if
	
	if ZimaAddModi = "添加" or ZimaAddModi = "修改" then
		zimano = request("zimano")
		zimaoc = request("zimaoc")
		ZimaGood = request("ZimaGood") 
		ZimaBad = request("ZimaBad")
		ZimaFOR = request("ZimaFOR")
		ZimaProdTime = request("ZimaProdTime")
		ZimaDownTime = request("ZimaDownTime")
		if ZimaAddModi = "添加" then
			response.write "ZIMA添加<br/>"
			response.write zimano & "<br/>"
			response.write zimaoc & "<br/>"
			response.write ZimaGood & "<br/>"
			response.write ZimaBad & "<br/>"
			response.write ZimaFOR & "<br/>"
			response.write ZimaProdTime & "<br/>"
			response.write ZimaDownTime & "<br/>"
			response.end
		elseif ZimaAddModi = "修改" then
			response.write "ZIMA修改<br/>"
			response.write "<br/>"
			response.end
		end if
	end if
	
	ZimaJava = ""
	ProductionTime = 0
	DownTime = 0
	BlockedProducts = 0
	realtowh1 = 0
	' set rs = Server.CreateObject("adodb.recordset")
	' sql = "SELECT TOP 1 * FROM [DWH].[dbo].[Report_Day_RA2] where [day] = '"&Ddate&"' AND ProdlineNr = '"&LineName&"' AND ShiftAcronym = '"&DDN&"'"
	' rs.open sql,conn,1,1
	' if rs.eof then
		Adata = "写入"
	' else
		' realtowh1 = rs("RealToWh")
		' Adata = "修改"
	' end if
%>
</br></br>
<center>
<form id="form1" name="form1" method="post" action="Editer.asp">
	<table width="80%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>产品：<select name="ProductName" id="TmProduct"></select></td>
						<td>线别：<select name="LineName" id="TmLine"></select></td>
						<select name="CellName" id="TmCell" style="display:none" ></select></td>
						<script type="text/javascript">
							MenuInit('TmProduct', 'TmLine', 'TmCell', '<%=ProductName%>', '<%=LineName%>', '1');
						</script>
						<td>日期：
								<input type="text" name="Ddate" id="Ddate" value="<%=Ddate%>" size="6"/>
								<img onClick="WdatePicker({el:'Ddate'})" src="../include/DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">
						</td>
						<td>D/N：<select name="DDN" >
									<option value="T" <%if DDN = "T" then response.write "selected"%> >D</option>
									<option value="N" <%if DDN = "N" then response.write "selected"%> >N</option>
								</select>
						</td>
						<td>
							<input type="Submit" name="SearchEdit" value="提交" />
						</td>
					</tr>
				</table>
			</td>
		</tr>
</form>
		<%if SearchEdit = "提交" or SearchEdit = "ok" then
			if LineName = "2" or LineName = "4" or LineName = "5" or LineName = "6" then
			set cellrs = Server.CreateObject("adodb.recordset")
			CellSql = "SELECT [DWShiftID],[Day],[ShiftBegin],[ShiftEnd],[ShiftAcronym],[ProdlineNr],[ExecutionStateAcronym],[ProductionTime],[ScheduledStopTime],[Cell1_PiecesGood],[Cell1_PiecesBad],[Cell2_PiecesGood],[Cell2_PiecesBad],[Cell3_PiecesGood],[Cell3_PiecesBad],[Cell4_PiecesGood],[Cell4_PiecesBad],[Cell5_PiecesGood],[Cell5_PiecesBad],[Transfer1_PiecesGood],[Transfer1_PiecesBad],[BlockedProducts],[RealToWh] FROM [DWH].[dbo].[Report_Day_RA]"
			CellSql = CellSql & " WHERE [Day] = '"&Ddate&"' AND ShiftAcronym = '"&DDN&"' AND ProdlineNr = '"&LineName&"'"
			cellrs.open CellSql,conn,1,1
				ProductionTime = 0
				DownTime = 720
				BlockedProducts = 0
				Transfer1G = 0
				Transfer1B = 0
				Transfer1F = 0
				Cell1PiecesGood = 0
				Cell1PiecesBad = 0
				Cell1FOR = 0
				Cell2PiecesGood = 0
				Cell2PiecesBad = 0
				Cell2FOR = 0
				Cell3PiecesGood = 0
				Cell3PiecesBad = 0
				Cell3FOR = 0
				Cell4PiecesGood = 0
				Cell4PiecesBad = 0
				Cell4FOR = 0
				Cell5PiecesGood = 0
				Cell5PiecesBad = 0
				Cell5FOR = 0
				CellSubmit = "添加"
			if not cellrs.eof then
				CellSubmit = "修改"
				ProductionTime = cellrs("ProductionTime")
				DownTime = cellrs("ScheduledStopTime")
				BlockedProducts = cellrs("BlockedProducts")
				Transfer1G = cellrs("Transfer1_PiecesGood")
				Transfer1B = cellrs("Transfer1_PiecesBad")
				Transfer1F = Formatnumber(Transfer1B / (Transfer1G + Transfer1B) * 100,2)
				
				Cell1PiecesGood = cellrs("Cell1_PiecesGood")
				Cell1PiecesBad = cellrs("Cell1_PiecesBad")
				Cell1FOR = Formatnumber(Cell1PiecesBad / (Cell1PiecesGood + Cell1PiecesBad) * 100,2)
				
				Cell2PiecesGood = cellrs("Cell2_PiecesGood")
				Cell2PiecesBad = cellrs("Cell2_PiecesBad")
				Cell2FOR = Formatnumber(Cell2PiecesBad / (Cell2PiecesGood + Cell2PiecesBad) * 100,2)
				
				Cell3PiecesGood = cellrs("Cell3_PiecesGood")
				Cell3PiecesBad = cellrs("Cell3_PiecesBad")
				Cell3FOR = Formatnumber(Cell3PiecesBad / (Cell3PiecesGood + Cell3PiecesBad) * 100,2)
				
				Cell4PiecesGood = cellrs("Cell4_PiecesGood")
				Cell4PiecesBad = cellrs("Cell4_PiecesBad")
				Cell4FOR = Formatnumber(Cell4PiecesBad / (Cell4PiecesGood + Cell4PiecesBad) * 100,2)
				
				Cell5PiecesGood = cellrs("Cell5_PiecesGood")
				Cell5PiecesBad = cellrs("Cell5_PiecesBad")
				Cell5FOR = Formatnumber(Cell5PiecesBad / (Cell5PiecesGood + Cell5PiecesBad) * 100,2)
			end if
			%>
		<tr>
			<td>
					</br>
					<center>
					<form name="Cell" method="post" action="Editer.asp" >
					<input type="text" name="Ddate" value="<%=Ddate%>" /><input type="text" name="ProductName" value="<%=ProductName%>" /><input type="text" name="LineName" value="<%=LineName%>" /><input type="text" name="DDN" value="<%=DDN%>" />
					<table width="96%" border="1" cellspacing="0" cellpadding="0">
							<tr>
								<td align="center">
									Production Time：<input type="text" name="ProductionTime" value="<%=ProductionTime%>" size="5"/>
								</td>
								<td align="center">
									Down Time：<input type="text" name="DownTime" value="<%=DownTime%>" size="5"/>
								</td>
								<td align="center">
									Blocked Products：<input type="text" name="BlockedProducts" value="<%=BlockedProducts%>" size="5"/>
								</td>
							</tr>
						</table>
					</br>
					<table width="96%" border="1" cellspacing="0" cellpadding="0">
					  <tr>
						<th width="31%">Machine</th>
						<th width="36%">Good pieces</th>
						<th width="22%">Bad picesc</th>
						<th width="11%">FOR</th>
					  </tr>
					  <tr>
						<td align="center">Transfer 1</td>
						<td align="center"><input type="text" name="Transfer1G" id="T1G" onChange="jst1f()" value="<%=Transfer1G%>" size="6" /></td>
						<td align="center"><input type="text" name="Transfer1B" id="T1B" onChange="jst1f()" value="<%=Transfer1B%>" size="6" /></td>
						<td align="center"><input type="text" name="Transfer1F" id="T1F" value="<%=Transfer1F%>" size="6" /></td>
					  </tr>
						<tr>
						<td align="center">Cell 1</td>
						<td align="center"><input type="text" name="Cell1PiecesGood" id="C1G" onChange="jscf(1)" value="<%=Cell1PiecesGood%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell1PiecesBad" id="C1B" onChange="jscf(1)" value="<%=Cell1PiecesBad%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell1FOR" id="C1F" value="<%=Cell1FOR%>" size="6" /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 2</td>
						<td align="center"><input type="text" name="Cell2PiecesGood" id="C2G" onChange="jscf(2)" value="<%=Cell2PiecesGood%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell2PiecesBad" id="C2B" onChange="jscf(2)" value="<%=Cell2PiecesBad%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell2FOR" id="C2F" value="<%=Cell2FOR%>" size="6" /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 3</td>
						<td align="center"><input type="text" name="Cell3PiecesGood" id="C3G" onChange="jscf(3)" value="<%=Cell3PiecesGood%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell3PiecesBad" id="C3B" onChange="jscf(3)" value="<%=Cell3PiecesBad%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell3FOR" id="C3F" value="<%=Cell3FOR%>" size="6" /></td>
					  </tr>
					<%if ProductName = "RA" then%>
					  <tr>
						<td align="center">Cell 4</td>
						<td align="center"><input type="text" name="Cell4PiecesGood" id="C4G" onChange="jscf(4)" value="<%=Cell4PiecesGood%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell4PiecesBad" id="C4B" onChange="jscf(4)" value="<%=Cell4PiecesBad%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell4FOR" id="C4F" value="<%=Cell4FOR%>" size="6" /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 5</td>
						<td align="center"><input type="text" name="Cell5PiecesGood" id="C5G" onChange="jscf(5)" value="<%=Cell5PiecesGood%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell5PiecesBad" id="C5B" onChange="jscf(5)" value="<%=Cell5PiecesBad%>" size="6" /></td>
						<td align="center"><input type="text" name="Cell5FOR" id="C5F" value="<%=Cell5FOR%>" size="6" /></td>
					  </tr>
					<%end if%>
					</table>
					<input type="submit" name="CellAddModi" value="<%=CellSubmit%>" />
					</form>
					</center>
				<%
				end if
				if LineName = "Zima" or LineName = "Zima" then%>
					</br>
					<table width="96%" border="1" cellspacing="0" cellpadding="0">
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
						for ZimaLoop = 1 to ZimaCounts
						set zimars = Server.CreateObject("adodb.recordset")
						ZimaSql = "SELECT [DWShiftID],[Day],[ShiftBegin],[ShiftEnd],[ShiftAcronym],[RMachineID],[MachineAcronym],[MachineNr],[ProductionTime],[NoProductionTime],[ScheduledStopTime],[ProductionStartTime],[MaintenanceTime],[PiecesGood],[PiecesBad],round(convert(FLOAT,isnull(PiecesBad,0)) / (isnull(PiecesGood,1)+1),4) ZimaFOR FROM [DWH].[dbo].[Report_Day_Peripherie_"&ProductName&"] WHERE MachineAcronym = 'Zima' and MachineNr = '"&ZimaLoop&"' AND [Day] = '"&Ddate&"' AND ShiftAcronym = '"&DDN&"'"
						'response.write ZimaSql
						zimars.open ZimaSql,conn,1,1
						MachineClose = "checked"
						ZimaSubmit = "添加"
						ZimaGood = 0
						ZimaBad = 0
						ZimaFOR = 0
						ZimaProdTime = 0
						ZimaDownTime = 720
						if not zimars.eof then
							ZimaGood = zimars("PiecesGood")
							ZimaBad = zimars("PiecesBad")
							ZimaFOR = zimars("ZimaFOR")
							ZimaProdTime = zimars("ProductionTime")
							ZimaDownTime = zimars("NoProductionTime")
							ZimaSubmit = "修改"
						end if
						if ZimaDownTime = 720 then MachineClose = ""
						%>
						<form name="form<%=ZimaLoop%>" method="post" action="Editer.asp">
							<tr>
								<td align="center">Zimas <%=ZimaLoop%><input type="text" name="zimano" value="<%=ZimaLoop%>"/><input type="text" name="Ddate" value="<%=Ddate%>"></td>
								<td align="center">
									<input type="checkbox" name="zimaoc" <%=MachineClose%> />Machine runs
								</td>
								<td align="center"><input type="text" name="ZimaGood" id="Z<%=ZimaLoop%>G" onChange="zimajs(<%=ZimaLoop%>)" value="<%=ZimaGood%>" size="5" /></td>
								<td align="center"><input type="text" name="ZimaBad" id="Z<%=ZimaLoop%>B" onChange="zimajs(<%=ZimaLoop%>)" value="<%=ZimaBad%>" size="5" /></td>
								<td align="center"><input type="text" name="ZimaFOR" id="Z<%=ZimaLoop%>F" value="<%=ZimaFOR%>" size="5" /></td>
								<td align="center"><input type="text" name="ZimaProdTime" value="<%=ZimaProdTime%>" size="5" /></td>
								<td align="center"><input type="text" name="ZimaDownTime" value="<%=ZimaDownTime%>" size="5" /></td>
								<td align="center"><input type="submit" name="ZimaAddModi" value="<%=ZimaSubmit%>" />
							</tr>
						</form>
						<%
						next
						%>
					</table>
				<%elseif (ProductName = "RA" and LineName = "Topfzus") or (ProductName = "Petra" and LineName = "Topfzus") then%>
					</br>
					<table width="96%" border="1" cellspacing="0" cellpadding="0">
						<tr>
							<th width="18%">Machine</th>
							<th width="18%">Good pieces</th>
							<th width="18%">Bad picesc</th>
							<th width="10%">FOR</th>
							<th width="18%">Production Time</th>
							<th width="18%">Down Time</th>
						</tr>
						<tr>
							<td align="center">Topfzus 1</td>
							<td align="center"><input type="text" name="Transfer1" id="C1G" onChange="jsc1f()" <%=Transfer1%>></td>
							<td align="center"><input type="text" name="Transfer1" id="C1B" onChange="jsc1f()" <%=Transfer1%>></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
						</tr>
					</table>
				<%elseif ProductName = "Petra" and LineName = "Pka" then%>
					</br>
					<table width="96%" border="1" cellspacing="0" cellpadding="0">
						<tr>
							<th width="18%">Machine</th>
							<th width="18%">Good pieces</th>
							<th width="18%">Bad picesc</th>
							<th width="10%">FOR</th>
							<th width="18%">Production Time</th>
							<th width="18%">Down Time</th>
						</tr>
						<tr>
							<td align="center">Pkas 1</td>
							<td align="center"><input type="text" name="Transfer1" id="C1G" onChange="jsc1f()" <%=Transfer1%>></td>
							<td align="center"><input type="text" name="Transfer1" id="C1B" onChange="jsc1f()" <%=Transfer1%>></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
							<td align="center"><input type="text" name="Transfer1" id="C1F" <%=Transfer1%> size="5"></td>
						</tr>
					</table>
				<%end if%>
			</td>
		</tr>
		<%end if%>
	</table>
</center>
</br></br></br></br></br></br></br>
</body>
</html>
