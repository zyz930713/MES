<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_Out"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>出库操作</title>
<link href="Styles/Basic.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript">
function lookup(inputString) {
  if(inputString.length == 0) {
   // Hide the suggestion box.
   $('#suggestions').hide();
  } else {
   $.post("showmember.asp", {queryString: ""+escape(inputString)+""}, function(data){
    if(data.length >0) {
     $('#suggestions').show();
     $('#autoSuggestionsList').html(unescape(data));
    }
   });
  }
 } // lookup
 
function fill(thisValue) {
  $('#PART_NUMBER').val(thisValue);
  setTimeout("$('#suggestions').hide();", 200);
 }

function lookup1(inputString) {
  if(inputString.length == 0) {
   // Hide the suggestion box.
   $('#suggestions1').hide();
  } else {
   $.post("showmember.asp", {cstname: ""+escape(inputString)+""}, function(data){
    if(data.length >0) {
     $('#suggestions1').show();
     $('#autoSuggestionsList1').html(unescape(data));
    }
   });
  }
 } // lookup
 
function fille(thisValue) {
  $('#CUSTOMER_NAME').val(thisValue);
  setTimeout("$('#suggestions1').hide();", 200);
 }
</script>
<style type="text/css">
.suggestionsBox {
	position:relative;
	margin:10px 0px 0px 0px;
	width:200px;
	background-color: #212427;
	border:2px solid #000; 
	color:#fff;
}
.suggestionList {
	margin:0px;
	padding:0px;
}
.suggestionList li {
	padding:3px;
	cursor:pointer;
	list-style:none;
}
.suggestionList li:hover {
	background-color:#659CD8;
}
table tr{
	background:#cccccc;
}
tr.change:hover
{
	background-color:#66A6FF
}
</style>
</head>
<body style="margin:20px auto;background-color:#339966;">
<center>
<%
SubmitSearch = request("SubmitSearch")
ItemID = trim(request("ItemID"))
ItemName = UCASE(trim(request("ItemName")))
VendorName = trim(request("VendorName"))
MENGEWH = request("MENGEWH")
LoctorX = request("LoctorX")
LoctorY = request("LoctorY")
LoctorZ = request("LoctorZ")
LoctorT = request("LoctorT")
Remark = request("Remark")
IOWh = request("IOWh")
BtnRemark = request("BtnRemark")
OpCode = "2"

IF (not isnull(IOWh)) and IOWh = "出库" then

	'下行可以进行优化
	if (not IsNumeric(MENGEWH)) or MENGEWH <= 0 then call sussLoctionHref("数量必须为数字并且不能为0","WarehouseOut.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
	
	set RsWh = Server.CreateObject("adodb.recordset")
	SqlWh = "select * from wh_rec_item where item_id = '"&ItemId&"'"
	'response.write SqlWh
	RsWh.open SqlWh,conn,1,1
	if RsWh.eof then
		call sussLoctionHref("没有可以出库的物料","WarehouseOut.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
	else
		menge = csng(RsWh("menge"))
		MENGEWH = csng(MENGEWH)
		
		if MENGEWH > menge then
			call sussLoctionHref("您出库数量大于库存数量,请重新输入!","WarehouseOut.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
		else
			BESTAND = menge - MENGEWH
			OutSql = "update WH_REC_ITEM set menge = '"&BESTAND&"',Remark = '"&Remark&"',LAST_OUT_DATE = SYSDATE where item_id = '"&ItemId&"'"
		end if
		
		conn.execute(OutSql)
		conn.execute("delete from WH_REC_ITEM where menge = 0") '清除零库存的记录
		
		'写入操作记录
		InsertOperationSql = "insert into WH_REC_OPERATION(ITEM_NAME,VENDOR_Name,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T) VALUES('"&ItemName&"','"&VendorName&"',SYSDATE,'"&OpCode&"','"&UserCode&"','"&MENGEWH&"','"&BESTAND&"','"&LoctorX&"','"&LoctorY&"','"&LoctorZ&"','"&LoctorT&"')"
		conn.execute(InsertOperationSql)
		call sussLoctionHref("成功出库：　"&MENGEWH,"WarehouseOut.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
	end if
end if

if BtnRemark = "注" then
	conn.execute("update WH_REC_ITEM set Remark = '"&Remark&"' where item_name = '"&ItemName&"' and loctor_x = '"&LoctorX&"' and loctor_y = '"&LoctorY&"' and loctor_z = '"&LoctorZ&"' and loctor_t = '"&LoctorT&"'")
	call sussLoctionHref("备注修改成功!","WarehouseOut.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
end if
%>
<table width="1300" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">出库操作</div>
		</td>
	</tr>	
	<tr>
		<td>
			<div align="center">
				<form name="form1" method="post" action="WarehouseOut.asp" >
				料号：<input type="text" name="ItemName" value="<%=ItemName%>" id="PART_NUMBER" onkeyup="lookup(this.value);" onblur="fill();"/>&nbsp;
				类型：<select name="LoctorT">
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'T' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								dim SelectStr
								response.write "<option value='ALL' >所有类型</option>"
								while not RsLoctor.eof
									loctor_value = RsLoctor("loctor_value")
									if loctor_value = LoctorT then
										SelectStr = "Selected"
									else
										SelectStr = ""
									end if
									response.write "<option value='"&loctor_value&"' "&SelectStr&">"&loctor_value&"</option>"
								RsLoctor.movenext
								wend
								RsLoctor.close
								set RsLoctor = nothing
							%>
						</select>&nbsp;
						<input type="submit" name="SubmitSearch" value="查询" class="t-b-midautumn" />
					<div class="suggestionsBox" id="suggestions" style="display: none;">
					<div class="suggestionList" id="autoSuggestionsList">&nbsp;</div>
				</form>
			</div>
		</td>
	</tr>
</table>
<table>
<br>
<%
if SubmitSearch = "查询" or SubmitSearch = "ok" then

	SQL = "select wri.Item_ID,wri.item_name,pm.description,wri.vendor_num,s.vendor_name,s.vendor_name_alt,unit,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name left join suppliers s on wri.vendor_num = s.vendor_num where wri.item_name = '"&ItemName&"'"
	IF LoctorT <> "ALL" THEN SQL = SQL & " and LOCTOR_T like '"&LoctorT&"'"
	SQL = SQL & " order by LAST_IN_DATE asc"
	rs.open SQL,conn,1,1
	
	if not rs.eof then
		ItemName = rs("Item_name")
		unit = rs("unit")
		description = rs("description")
%>
<table width="1300" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<th width="140">料号</th>
		<th width="50">单位</th>
		<th width="1138">描述</th>
	</tr>
	  <tr class="t-t-List">
		<td width="156"><%=ItemName%></td>
		<td align="center"><%=unit%></td>
		<td align="left"><%=Description%></td>
	</tr>
</table>
<br>
<table width="1300" border="1" cellpadding="0" cellspacing="0" >
	<tr>
		<th>供应商</th>
		<th width='150'>入库日期</th>
		<th width='150'>出库日期</th>
		<th width="30">排</th>
		<th width="30">列</th>
		<th width="30">高</th>
		<th width="50">类型</th>
		<th width="70">库存</th>
		<th width='50'>数量</th>
		<th width='40'>操作</th>
		<th width='180'>备注</th>
	</tr>
<%
		while not rs.eof
		ItemID = rs("Item_ID")
		VendorNum = rs("vendor_num")
		VendorNameEn = rs("vendor_name")
		VendorNameCn = rs("vendor_name_alt")
		
		if isnull(VendorNameCn) then
			VendorName = VendorNameEN
		else
			VendorName = VendorNameCn
		end if
		MENGE = rs("MENGE")
		DayIn = rs("LAST_IN_DATE")
			if isnull(DayIn) then DayIn = "　"
		DayOut = rs("LAST_Out_DATE")
			if isnull(DayOut) then DayOut = "　"
		LoctorX = rs("LOCTOR_X")
		LoctorY = rs("LOCTOR_Y")
		LoctorZ = rs("LOCTOR_Z")
		LoctorT = rs("LOCTOR_T")
		Remark = rs("Remark")

		%>
		<form name="form2" method="post" action="WarehouseOut.asp">
			<tr class="change">
				<td><%=VendorName%><input type="hidden" name="VendorName" size="2" value="<%=VendorName%>"/><input type="hidden" name="ItemName" size="2" value="<%=ItemName%>"/><input type="hidden" name="ItemID" size="2" value="<%=ItemID%>"/></td>
				<td><%=DayIn%></td>
				<td><%=DayOut%></td>
				<td><%=LoctorX%><input type="hidden" name="LoctorX" size="2" value="<%=LoctorX%>"/></td>
				<td><%=LoctorY%><input type="hidden" name="LoctorY" size="2" value="<%=LoctorY%>"/></td>
				<td><%=LoctorZ%><input type="hidden" name="LoctorZ" size="2" value="<%=LoctorZ%>"/></td>
				<td><%=LoctorT%><input type="hidden" name="LoctorT" size="2" value="<%=LoctorT%>"/></td>
				<td><%=MENGE%></td>
				<td><input type='text' name='MENGEWH' id='MENGEWH' value='0' size='3'/></td>
				<td><input type='Submit' name='IOWh' value='出库' onClick='return beforeSubmit()'/></td>
				<td><input type="text" name="Remark" value="<%=Remark%>"/> <input type="Submit" name="BtnRemark" value="注" /></td>
			</tr>
		</form>
		<%
		rs.movenext
		wend
		%>
		</table>
		<%else
			response.write "没有找到此物料<br/>"
		end if%>
<%end if%>
<br>
  <table width="1300" border="1" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center">
      <input type="button" value="修改密码" onClick="window.location.reload('../../Functions/UserPass_Modify.asp');" />
      　　　　
      <input type="button" value="Close关闭" onClick="window.location.reload('../../Functions/User_Logout.asp');" /></td>
    </tr>
  </table>
</center>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->