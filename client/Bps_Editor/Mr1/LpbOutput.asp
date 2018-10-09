<!--#include file= "conn.asp"-->
<!--#include file= "../include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>LPB</title>
<script language="javascript" type="text/javascript" src="../include/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="../include/DatePicker/WdatePicker.js"></script>
</head>
<body>

<%
	AddLpb2 = request("AddLpb2")
	ProductName = request("ProductName")
	LineName = request("LineName")
	DDN = request("DDN")
	HShift = request("HShift")
	Ddate = request("Ddate")
	
	'写入生产数量及废品率信息
	if AddLpb2 = "临时写入" or AddLpb2 = "确认写入" or AddLpb2 = "修改数据" then
	
		LineNo = ProductName&LineName
		HSSV = request("HSSV")
		HLE = request("HLE")
		Hespeed = request("Hespeed")
		Hahour = request("Hahour")
		Heqty = request("Heqty")
		Haqty = request("Haqty")
		'Hoee = request("Hoee")
		oee = Formatnumber((Haqty / (Hespeed * Hahour)) * 100,2)
		wima = request("wima")
		cell1 = request("cell1")
		cell2 = request("cell2")
		cell3 = request("cell3")
		cell4 = request("cell4")
		cell5 = request("cell5")
		
		sm = request("sm")
		oc = request("oc")
		ts = request("ts")
		mt = request("mt")
		
		Qablock = request("Qablock")
		Qarelease = request("Qarelease")
		
		Hsend = request("Hsend")
		Hback = request("Hback")
		
		if AddLpb2 = "临时写入" then
			' set rs = server.createobject("adodb.recordset") 
				' sql = "select top 1 * FROM [Ministat].dbo.BJLPB2"
				' rs.open sql,conn,1,3
					' rs.addnew
					' rs("HID") = "BU"
					' rs("Hlineno") = LineNo
					' rs("HDate") = Ddate
					' rs("Udate") = now()
					' rs("HD_N") = DDN
					' rs("HShift") = HShift
					' rs("HSSV") = HSSV
					' rs("HLE") = HLE
					' rs("Hespeed") = Hespeed
					' rs("Hahour") = Hahour
					' rs("Heqty") = Heqty
					' rs("Haqty") = Haqty
					' rs("HOEE") = oee
					' rs("Wima") = wima
					' rs("Cell1") = cell1
					' rs("Cell2") = Cell2
					' rs("Cell3") = Cell3
					' rs("cell4") = cell4
					' rs("Cell21") = Cell5
					' rs.update
				' rs.close
				' set rs = nothing
				
				Sqlstr1 = "INSERT INTO [Ministat].dbo.BJLPB2(HID,Hlineno,HDate,Udate,HD_N,HShift,HSSV,HLE,Hespeed,Hahour,Heqty,Haqty,HOEE,Wima,Cell1,Cell2,Cell3,cell4,Cell21)"
				Sqlstr = Sqlstr1 & "VALUES('BU','"&LineNo&"','"+Ddate+"',GETDATE(),'"&DDN&"','"&HShift&"','"&HSSV&"','"&HLE&"','"+Hespeed+"','"+Hahour+"','"+Heqty+"','"+Haqty+"','"+oee+"','"&wima&"','"&cell1&"','"&Cell2&"','"&Cell3&"','"&cell4&"','"&Cell5&"')"
				conn.execute(Sqlstr)
				' response.write Sqlstr
				' response.end
				
				
				'写入在线检测数据
				Sqlstr1 = "INSERT INTO [data-warehouse].dbo.OnlineData2(CID,CDate,Center,CLINE,Cshift,CShiftType,CType,CQTY)"
				Sqlstr = Sqlstr1 & " VALUES('CT','"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','MT','"&MT&"')"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & " VALUES('CT','"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','OC','"&oc&"')"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & " VALUES('CT','"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','SM','"&sm&"')"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & " VALUES('CT','"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','TS','"&ts&"')"
				conn.execute(Sqlstr)
				'写入在线检测数据

				'写入封存数据
				Sqlstr1 = "INSERT INTO [Ministat].[dbo].[QABlock2](CDate,CCenter,CLINE,CShift,CShiftType,CType,CQTY,upddate)"
				Sqlstr = Sqlstr1 & " VALUES('"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Block','"&Qablock&"',GETDATE())"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & " VALUES('"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Release','"&Qarelease&"',GETDATE())"
				conn.execute(Sqlstr)
				'写入封存数据

				Sqlstr1 = "INSERT INTO [Ministat].[dbo].[HGRecheck2](CDate,CCenter,CLINE,CShift,CShiftType,CType,CQTY,upddate)"
				Sqlstr = Sqlstr1 & " VALUES('"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Send','"&Hsend&"',GETDATE())"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & " VALUES('"+Ddate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Back','"+Hback+"',GETDATE())"
				conn.execute(Sqlstr)

				call sussLoctionHref("添加成功","LpbOutput.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
				
			elseif AddLpb2 = "确认写入" or AddLpb2 = "修改数据" then
				'修改在线检测数据
				Sqlstr1 = "UPDATE [data-warehouse].[dbo].[OnlineData2] SET CQTY = "
				Sqlstr = Sqlstr1 & "'"&MT&"' where CType = 'MT' and CDate = '"&Ddate&"' and CLine = '"&LineNo&"' AND CShift = '"&HShift&"'"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & "'"&oc&"' where CType = 'OC' and CDate = '"&Ddate&"' and CLine = '"&LineNo&"' AND CShift = '"&HShift&"'"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & "'"&sm&"' where CType = 'SM' and CDate = '"&Ddate&"' and CLine = '"&LineNo&"' AND CShift = '"&HShift&"'"
				conn.execute(Sqlstr)
				Sqlstr = Sqlstr1 & "'"&ts&"' where CType = 'TS' and CDate = '"&Ddate&"' and CLine = '"&LineNo&"' AND CShift = '"&HShift&"'"
				conn.execute(Sqlstr)
				'修改在线检测数据
			
				call sussLoctionHref("修改成功","LpbOutput.asp?Ddate="&Ddate&"&ProductName="&ProductName&"&LineName="&LineName&"&DDN="&DDN&"&HShift="&HShift)
				response.end
			end if
	end if
	
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
	
	if DDN = "" then
		DDN = request.Cookies("Lpb")("DDN")
	else
		response.Cookies("Lpb")("DDN") = DDN
	end if
	
	if HShift = "" then
		HShift = request.Cookies("Lpb")("HShift")
	else
		response.Cookies("Lpb")("HShift") = HShift
	end if
	
	if Ddate = "" then
		Ddate = FormatTime(request.Cookies("Lpb")("Ddate"),2)
	else
		response.Cookies("Lpb")("Ddate") = Ddate
	end if
	'取Cookies缓存数据'
	
	if DDN = "" then
		if 8 <= hour(now()) and hour(now()) <= 20 then
			DDN = "D"
			response.Cookies("Lpb")("DDN") = "D"
		else
			DDN = "N"
			response.Cookies("Lpb")("DDN") = "N"
		end if
	end if
	
	if Ddate = "" then
		Ddate = FormatTime(date(),2)
		response.Cookies("Lpb")("Ddate") = Ddate
	end if

	if DDN = "D" then
		Stime = cdate(Ddate & " 08:15:00")
		Etime = dateadd("h",12,Stime)
	else
		Stime = cdate(Ddate & " 20:15:00")
		Etime = dateadd("h",12,Stime)
	end if

'获取停机时间'
Dlineno = ProductName & LineName
SumDtime = 0
set rs = Server.CreateObject("adodb.recordset")
Sql = "select sum(Dtime) DtimeSum FROM Ministat.dbo.BJLPBD2 where Dlineno = '"&Dlineno&"' and DD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and Ddate = '"&Ddate&"'"
'response.write sql
rs.open sql,conn,1,1
if not rs.eof then
	SumDtime = rs("DtimeSum")
end if
rs.close
set rs = nothing
'获取停机时间'

'获取并确认是否有产出数据'
set rs = Server.CreateObject("adodb.recordset")
SqlStr = "SELECT TOP 1 * FROM [Ministat].[dbo].[BJLPB2] where Hlineno = '"&Dlineno&"' and HD_N = '"&DDN&"' AND Hshift = '"&HShift&"' and HDate = '"&Ddate&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then
	HID = rs("HID")
	HSSV = rs("HSSV")
	HLE = rs("HLE")
	Hahour = rs("Hahour")
	Haqty = rs("Haqty")
	HOEE = rs("HOEE")
	WimaFOR = rs("Wima")
	Cell1FOR = rs("Cell1")
	Cell2FOR = rs("Cell2")
	Cell3FOR = rs("Cell3")
	Cell4FOR = rs("Cell4")
	Cell5FOR = rs("Cell21")
	if HID = "HB" then
		AddMenuName = "确认写入"
	elseif HID = "BU" then
		AddMenuName = "修改数据"
	end if
else
	Hahour = 12
	Haqty = 0
	HOEE = 0
	WimaFOR = 0
	Cell1FOR = 0
	Cell2FOR = 0
	Cell3FOR = 0
	Cell4FOR = 0
	Cell5FOR = 0
	AddMenuName = "临时写入"
end if
rs.close
'获取并确认是否有产出数据'

SMCQTY = 0
OCCQTY = 0
TSCQTY = 0
MTCQTY = 0
'找SM数量
SqlStr = "select * from [data-warehouse].[dbo].[OnlineData2] where CType = 'SM' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn2,1,1
if not rs.eof then SMCQTY = rs("CQTY")
rs.close
'找OC数量
SqlStr = "select * from [data-warehouse].[dbo].[OnlineData2] where CType = 'OC' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn2,1,1
if not rs.eof then OCCQTY = rs("CQTY")
rs.close
'找TS数量
SqlStr = "select * from [data-warehouse].[dbo].[OnlineData2] where CType = 'TS' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn2,1,1
if not rs.eof then TSCQTY = rs("CQTY")
rs.close
'找MT数量
SqlStr = "select * from [data-warehouse].[dbo].[OnlineData2] where CType = 'MT' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn2,1,1
if not rs.eof then MTCQTY = rs("CQTY")
rs.close

BackQty = 0
SendQty = 0
BQTY = 0
RQTY = 0
SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.HGRecheck2 WHERE CType = 'Back' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then BackQty = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.HGRecheck2 WHERE CType = 'Send' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then SendQty = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock2 where CType = 'Block' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then BQTY = rs("CQTY")
rs.close

SqlStr = "SELECT TOP 1 * FROM Ministat.dbo.QABlock2 where CType = 'Release' and CDate = '"&Ddate&"' and CLine = '"&Dlineno&"' AND CShift = '"&HShift&"'-- and CShiftType = '"&DDN&"'"
rs.open SqlStr,conn,1,1
if not rs.eof then RQTY = rs("CQTY")
rs.close

%>
</br>
<center>
<form id="form1" name="form1" method="post" action="LpbOutput.asp">
	<table width="80%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td>生产中心信息:</td>
			<td>产品：<select name="ProductName" id="TmProduct"></select></td>
			<td>线别：<select name="LineName" id="TmLine"></select></td>
			<select name="CellName" id="TmCell" style="display:none" ></select></td>
			<script type="text/javascript">
				MenuInit('TmProduct', 'TmLine', 'TmCell', '<%=ProductName%>', '<%=LineName%>', '1');
			</script>
			<td>D/N：<select name="DDN" >
						<option value="D" <%if DDN = "D" then response.write "selected"%> >D</option>
						<option value="N" <%if DDN = "N" then response.write "selected"%> >N</option>
					</select>
			</td>
			<td>Shift:<select name="HShift" >
							<option value="A" <%if HShift = "A" then response.write "selected"%> >A</option>
							<option value="B" <%if HShift = "B" then response.write "selected"%> >B</option>
							<option value="C" <%if HShift = "C" then response.write "selected"%> >C</option>
						</select>
			</td>
			<td>日期<input type="text" name="Ddate" id="Ddate" onchange="form1.submit();" value="<%=Ddate%>" size="6"/><img onclick="WdatePicker({el:'Ddate'})" src="../include/DatePicker/skin/datePicker.gif" width="16" height="22" align="absmiddle">
			</td>
			<td>
				<input type="submit" name="AddNew" value="提交">
			</td>
		</tr>
	</table>
</br>
	<table width="80%" border="1" cellspacing="0" cellpadding="1">
	 <tr>
	    <td>生产数据:</td>
      </tr>
	  <tr><td>SSV：<input type="text" name="HSSV" size="5" value="<%=HSSV%>"/>LE：<input type="text" name="HLE" size="5" value="<%=HLE%>"/></td>
		</tr>
	  <tr>
	    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	      <tr>
	        <td align="right">理论产量/小时</td>
			<td><input name="Hespeed" type="text" id="Hespeed" value="6000" size="8" readonly /></td>
	        <td width="15%" align="right">实际生产时间</td>
	        <td width="8%"><input name="Hahour" type="text" id="Hahour" onChange="js1()" value="<%=Hahour%>" size="8" /></td>
	        <td width="11%" align="right">理论总产量</td>
	        <td width="14%"><input name="Heqty" type="text" id="Heqty" size="8" /></td>
	        <td width="10%" align="right">输入损失</td>
	        <td width="13%"><input type="text" name="input" value="0" size="8" /></td>
          </tr>
	      <tr>
	        <td align="right">实际总产量</td>
	        <td><input name="Haqty" type="text" id="Haqty" onBlur="js1()" value="<%=Haqty%>"  size="8" /></td>
	        <td align="right"><font color="red">停机时间</font></td>
	        <td><input type="text" name="Tj" value="0" id="Tj" size="8"  onChange="js1()" />分</td>
	        <td align="right"><font color="red">OEE</font></td>
	        <td><input name="Hoee" type="text" value="<%=HOEE%>" id="Hoee" size="8" />
	          %</td>
	        <td align="right"><font color="red">差异</font></td>
	        <td><input type="text" name="Cy" value="0" id="Cy" size="8" /></td>
          </tr>
        </table></td>
      </tr>
	  <tr>
	    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	      <tr>
	        <td>For:</td>
	        <td>WIMA</td>
	        <td>CELL1</td>
	        <td>CELL2</td>
	        <td>CELL3</td>
	        <td>CELL4</td>
	        <td>CELL5</td>
          </tr>
	      <tr>
	        <td>&nbsp;</td>
	        <td><input type="text" name="wima" size="3" value="<%=WimaFOR%>" /></td>
	        <td><input type="text" name="cell1" size="3" value="<%=Cell1FOR%>" /></td>
	        <td><input type="text" name="cell2" size="3" value="<%=Cell2FOR%>" /></td>
	        <td><input type="text" name="cell3" size="3" value="<%=Cell3FOR%>" /></td>
	        <td><input type="text" name="cell4" size="3" value="<%=Cell4FOR%>" /></td>
	        <td><input type="text" name="cell5" size="3" value="<%=Cell5FOR%>" /></td>
          </tr>
        </table></td>
      </tr>
	  <tr>
	    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	      <tr>
	        <td colspan="4">在线数据:</td>
          </tr>
	      <tr>
	        <td>采样-SM：<input type="text" name="SM" value="<%=SMCQTY%>" size="7" /></td>
	        <td>目测废品-OC：<input type="text" name="OC" value="<%=OCCQTY%>" size="7" /></td>
	        <td>清零废品-TS：<input type="text" name="TS" value="<%=TSCQTY%>" size="7" /></td>
	        <td>废膜/顶片-MT：<input type="text" name="MT" value="<%=MTCQTY%>" size="7" /></td>
          </tr>
        </table></td>
      </tr>
	  <tr>
	    <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
	      <tr>
	        <td colspan="2">质量封存与放行：</td>
	        <td colspan="2">手工组复检：</td>
          </tr>
	      <tr>
	        <td align="right">封存：<input type="text" name="QABlock" value="<%=BQTY%>" size="8" /></td>
	        <td align="right">放行：<input type="text" name="QARelease" value="<%=RQTY%>" size="8" /></td>
	        <td align="right">送交手工组：<input type="text" name="Hsend" value="<%=SendQty%>" size="8" /></td>
	        <td align="right">手工组返回：<input type="text" name="Hback" value="<%=BackQty%>" size="8" /></td>
          </tr>
        </table></td>
      </tr>
	  <%if (cdate(Stime) < Now()) and (Now() < cdate(Etime)) then %>
	  <tr>
	    <td align="right">
			&nbsp;合计停机时间:&nbsp;<input type="text" id="SumDtime" name="SumDtime" value="<%=SumDtime%>" size="2" style="text-align:center;" readonly />&nbsp;分钟
			<input type="Submit" name="AddLpb2" value="<%=AddMenuName%>" />
		</td>
      </tr>
	  <%end if%>
	  </table>
</form>
</center>
 
<script type="text/javascript">
SumDtime=document.getElementById("SumDtime").value;
document.form1.Tj.value=SumDtime;
window.onload=js1();

function js1()
{
Hespeed1=document.getElementById("Hespeed").value;
Hahour1=document.getElementById("Hahour").value;
Tjsj=document.getElementById("Tj").value;
hj=Hespeed1*Hahour1;
document.form1.Heqty.value=hj;
LLZCL1=document.getElementById("Heqty").value;
Haqty1=document.getElementById("Haqty").value;
Hoee=(Haqty1/LLZCL1*100).toFixed(2);
document.form1.Hoee.value=Hoee;

Tjsj = Tjsj/60;
HjCy = Tjsj*Hespeed1;
Llzcl = parseFloat(HjCy) + parseFloat(Haqty1);
OEE1 = parseFloat(Llzcl) / (hj/100);
OEE1 = 100 - parseFloat(OEE1);
OEE1 = OEE1.toFixed(2);
document.form1.Cy.value=OEE1;
}
</script>
</body>
</html> 
