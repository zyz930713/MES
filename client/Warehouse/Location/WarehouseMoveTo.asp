<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_Move"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>移库操作</title>
<link href="Styles/Basic.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript">

</script>
</head>
<body style="margin:2px;padding:0px;">
<center>
<%
SubmitSearch = request("SubmitSearch")

SubMove = request("SubMove")
BtnRemark = request("BtnRemark")

ItemName = UCASE(trim(request("ItemName")))
SitemId = request("SitemId")
SvendorNum = request("SvendorNum")
SDayIn = request("SDayIn")
SDayOut = request("SDayOut")
SLoctorX = request("SLoctorX")
SLoctorY = request("SLoctorY")
SLoctorZ = request("SLoctorZ")
SLoctorT = request("SLoctorT")
SMENGEWH = request("SMENGEWH")

LoctorX2 = request("LoctorX2")
LoctorY2 = request("LoctorY2")
LoctorZ2 = request("LoctorZ2")
LoctorT2 = request("LoctorT2")
MENGEWH2 = request("MENGEWH2")

'response.write ItemName & " - " &SitemId & " - " & SvendorNum & " - " & SDayIn & " - " & SLoctorX & " - " & SLoctorY & " - " & SLoctorZ & " - " & SLoctorT & " - " & SMENGEWH & "<br/>"
'response.write ItemName & " - " &LoctorX2 & " - " & LoctorY2 & " - " & LoctorZ2 & " - " & LoctorT2 & " - " & MENGEWH2 & "<br/>"
'Remark = request("Remark")
IF SubMove = "确认移库" then
	if len(LoctorX2) < 1 then call sussLoctionHrefFrame("还没有选择目标位置.","WarehouseMove.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
	SMENGEWH = csng(SMENGEWH)
	MENGEWH2 = csng(MENGEWH2)
	set RsWh = Server.CreateObject("adodb.recordset")
	if MENGEWH2 > SMENGEWH then
		call sussLoctionHrefFrame("移库数量大于库存,无法移动.","WarehouseMove.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
	else
		
		SqlWh = "select item_id,ITEM_NAME,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,REMARK,VENDOR_NUM from wh_rec_item where ITEM_ID = '"&SitemId&"'"
		'response.write SqlWh
		RsWh.open SqlWh,conn,1,1
		if not RsWh.eof then
			itemID1 = RsWh("item_id")
			itemName1= RsWh("ITEM_NAME")
			lastOutDate1 = RsWh("LAST_OUT_DATE")
			lastInDate1 = RsWh("LAST_IN_DATE")
			menge1 = RsWh("MENGE")
			LoctorX1 = RsWh("LOCTOR_X")
			LoctorY1 = RsWh("LOCTOR_Y")
			LoctorZ1 = RsWh("LOCTOR_Z")
			LoctorT1 = RsWh("LOCTOR_T")
			vendorNum1 = RsWh("VENDOR_NUM")
			RsWh.close
			SqlOut = "update WH_REC_ITEM set MENGE = MENGE - '"&MENGEWH2&"' where ITEM_ID = '"&itemID1&"'"
			conn.execute(SqlOut)
			conn.execute("delete from WH_REC_ITEM where menge = 0")
			
			OperationSqlOut = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'2','"&UserCode&"','"&MENGEWH2&"','"&BESTAND&"','"&LoctorX1&"','"&LoctorY1&"','"&LoctorZ1&"','"&LoctorT1&"','"&vendorNum1&"')"
			conn.execute(OperationSqlOut)
			response.write "<script>alert('执行到2了');</script>"
			
			SqlIn = "insert into WH_REC_ITEM(ITEM_NAME,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NUM) VALUES('"&ItemName&"',SYSDATE,'"&MENGEWH2&"','"&LoctorX2&"','"&LoctorY2&"','"&LoctorZ2&"','"&LoctorT2&"','"&vendorNum1&"')"
			conn.execute(SqlIn)
			
			OperationSqlIn = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'1','"&UserCode&"','"&MENGEWH2&"','"&MENGEWH2&"','"&LoctorX2&"','"&LoctorY2&"','"&LoctorZ2&"','"&LoctorT2&"','"&vendorNum1&"')"
			conn.execute(OperationSqlIn)
			
			call sussLoctionHrefFrame("成功移库：　"&MENGEWH2,"WarehouseMove.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
		else
			call sussLoctionHrefFrame("未找到可以移动的物料.","WarehouseMove.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
		end if
		RsWh.close
		set RsWh = nothing
	end if
end if

subMove = "<center>目标位无效</center>"

set RsWh = Server.CreateObject("adodb.recordset")
SqlWh = "select * from wh_rec_item where item_name = '"&ItemName&"' and loctor_x = '"&LoctorX2&"' and loctor_y = '"&LoctorY2&"' and loctor_z = '"&LoctorZ2&"'-- and loctor_t = '"&LoctorT&"'"
RsWh.open SqlWh,conn,1,1
if RsWh.eof then	'如果目标位没有物料,则直接入库.
	subMove = "<input type='Submit' name='SubMove' value='确认移库' />"
else	'否则取出日期类型,进行判断,如果相同,才允许.
	if SLoctorX = LoctorX2 and SLoctorY = LoctorY2 and SLoctorZ = LoctorZ2 and SLoctorT = LoctorT2 then
		subMove = "<input type='Submit' name='SubMove' value='确认移库' />"
	end if
end if

%>

<table width="1198" border="1" align="center" cellpadding="0" cellspacing="0" style="background-color:#cccccc;">
	<tr>
		<td height="20" colspan="4" class="t-t-DarkBlue">
			<div align="center">移库库位选择</div>
		</td>
	</tr>
	<tr>
		<td width="14%">料号</td>
		<td width="18%">入库日期　</td>
		<td width="18%">出库日期　</td>
	</tr>
	<tr>
		<td><%=ItemName%></td>
		<td><%=SDayIn%></td>
		<td><%=SDayOut%></td>
	</tr>
	<tr>
		<td colspan="4">
		</td>
	</tr>
</table>
			<form name="form1" method="post" action="WarehouseMoveTo.asp">
			<input type="hidden" name="ItemName" value="<%=ItemName%>"/>
			<input type="hidden" name="SitemId" value="<%=SitemId%>"/>
			<input type="hidden" name="SDayIn" size="2" value="<%=SDayIn%>"/>
			<input type="hidden" name="SDayOut" size="2" value="<%=SDayOut%>"/>
				<table width='1198' border='1' cellpadding='0' cellspacing='0' style="background-color:#cccccc;">
					<tr><th>排</th><th>列</th><th>高</th><th>类型</th><th>库存</th><th><font color="RED">移动到</span></font></th><th>排</th><th>列</th><th>高</th><th>类型</th><th>数量</th></tr>
					<tr>
						<td width="8%"><%=SLoctorX%><input type="hidden" name="SLoctorX" size="2" value="<%=SLoctorX%>"/></td>
						<td width="8%"><%=SLoctorY%><input type="hidden" name="SLoctorY" size="2" value="<%=SLoctorY%>"/></td>
						<td width="8%"><%=SLoctorZ%><input type="hidden" name="SLoctorZ" size="2" value="<%=SLoctorZ%>"/></td>
						<td width="10%"><%=SLoctorT%><input type="hidden" name="SLoctorT" size="2" value="<%=SLoctorT%>"/></td>
						<td width="10%"><%=SMENGEWH%><input type="hidden" name="SMENGEWH" size="2" value="<%=SMENGEWH%>"/></td>
						<td width="12%"><font color="RED">&nbsp;>>&nbsp;</span></font></td>
					<td width="8%">
						<select name="LoctorX2" ONCHANGE="this.form.submit()">
							<option value="">&nbsp;</option>
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'X' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								LoctorStr = ""
								while not RsLoctor.eof
									LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
									if LoctorX2 = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
									LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
									RsLoctor.movenext
								wend
								RsLoctor.close
								set RsLoctor = nothing
								response.write LoctorStr
							%>
						</select>
					</td>
					<td width="8%">
						<select name="LoctorY2" ONCHANGE="this.form.submit()">
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Y' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								LoctorStr = ""
								while not RsLoctor.eof
									LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
									if LoctorY2 = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
									LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
									RsLoctor.movenext
								wend
								RsLoctor.close
								set RsLoctor = nothing
								response.write LoctorStr
							%>
						</select>
					</td>
					<td width="8%">
						<select name="LoctorZ2" ONCHANGE="this.form.submit()">
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Z' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								LoctorStr = ""
								while not RsLoctor.eof
									LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
									if LoctorZ2 = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
									LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
									RsLoctor.movenext
								wend
								RsLoctor.close
								set RsLoctor = nothing
								response.write LoctorStr
							%>
						</select>
					</td>
					<td width="10%">
						<select name="LoctorT2" ONCHANGE="this.form.submit()">
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'T' and LOCTOR_TYPE = '1' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								LoctorStr = ""
								while not RsLoctor.eof
									LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
									if LoctorT2 = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
									LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
									RsLoctor.movenext
								wend
								RsLoctor.close
								set RsLoctor = nothing
								response.write LoctorStr
							%>
						</select>
					</td width="10%">
						<td><input type="text" name="MENGEWH2" id="MENGEWH2" value="0" size="3"/></td>
					</tr>
					<tr>
						<td colspan="11">
						<%=subMove%>
						</td>
					</tr>
				</table>
			</form>
</center>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->