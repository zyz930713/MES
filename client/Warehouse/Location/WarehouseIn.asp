<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_In"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>入库操作</title>
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
  $('#ItemName').val(thisValue);
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
</script>

<style type="text/css">
 .suggestionsBox {
  position: relative;
  left: 5px;
  margin: 10px 0px 0px 0px;
  width: 200px;
  background-color: #212427;
  -moz-border-radius: 7px;
  -webkit-border-radius: 7px;
  border: 2px solid #000; 
  color: #fff;
 }
 
 .suggestionList {
  margin: 0px;
  padding: 0px;
 }
 
 .suggestionList li {
  
  margin: 0px 0px 3px 0px;
  padding: 3px;
  cursor: pointer;
 }
 
 .suggestionList li:hover {
  background-color: #659CD8;
 }
table tr{
	background:#cccccc;
}
tr.change:hover
{
	background-color:#66A6FF;
}
</style>
</head>
<body style="margin:20px;background-color:#339966;">

<center>
<%
SubmitSearch = request("SubmitSearch")
itemid = request("itemid")
ItemName = UCASE(trim(request("ItemName")))
nVendroNum = trim(request("nVendroNum"))
nVendroName = trim(request("nVendroName"))
nMENGEWH = request("nMENGEWH")
nLoctorX = request("nLoctorX")
nLoctorY = request("nLoctorY")
nLoctorZ = request("nLoctorZ")
nLoctorT = request("nLoctorT")
nRemark = request("nRemark")
iowh = request("iowh")
BtnRemark = request("BtnRemark")
OpCode = "1"

if iowh = "入库" or iowh = "新入" then
	if not IsNumeric(nMENGEWH) then call sussLoctionHref("入库数必须为数字并且不能为0!","WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
	nMENGEWH = csng(nMENGEWH)

	set RsWh = Server.CreateObject("adodb.recordset")

	'按照入库的方法入库
	if iowh = "入库" then
		SqlWh = "select item_id,ITEM_NAME,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,REMARK,VENDOR_NUM from wh_rec_item where ITEM_ID = '"&itemid&"'"
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
		end if
		RsWh.close
		BESTAND = nMENGEWH + menge1
		insertSql = "update WH_REC_ITEM set MENGE = MENGE + '"&nMENGEWH&"' where ITEM_ID = '"&itemID1&"'"
		InsertOperationSql = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'"&OpCode&"','"&UserCode&"','"&nMENGEWH&"','"&BESTAND&"','"&LoctorX1&"','"&LoctorY1&"','"&LoctorZ1&"','"&LoctorT1&"','"&vendorNum1&"')"
		conn.execute(insertSql)'写入库
		conn.execute(InsertOperationSql)'写日志
		call sussLoctionHref("成功入库：　"&nMENGEWH,"WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
		response.end()
	end if
	'按照入库的方法入库
	
	'按照新入的方法入库
	if iowh = "新入" then
		Sql = "select item_id,ITEM_NAME,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,REMARK,VENDOR_NUM from wh_rec_item"
		SqlWh = Sql + " WHERE LOCTOR_X = '"&nLoctorX&"' and LOCTOR_Y = '"&nLoctorY&"' and LOCTOR_Z = '"&nLoctorZ&"' and LOCTOR_T = '"&nLoctorT&"'"
		RsWh.open SqlWh,conn,1,1
		if not RsWh.eof then	'目标位找到四项符合的物料
			'response.write "<script>alert('执行到1了');</script>"
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
			if month(lastInDate1) <> month(Now()) then	'判断入库的月份
				'response.write "<script>alert('执行到2了');</script>"
				call sussLoctionHref("目标库位已经有非本月物料,请重新选择库位!","WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
				response.end()
			else	'是本月物料,则开始判断供应商是否相同
				'response.write "<script>alert('执行到3了');</script>"
				if vendorNum1 <> nVendroNum then	'供应商不相同,则写入新库位
					'response.write "<script>alert('执行到4了');</script>"
					BESTAND = nMENGEWH
					insertSql = "insert into WH_REC_ITEM(item_name,vendor_num,LAST_IN_DATE,menge,loctor_x,loctor_y,loctor_z,loctor_t,Remark) VALUES('"&ItemName&"','"&nVendroNum&"',sysdate,'"&BESTAND&"','"&nLoctorX&"','"&nLoctorY&"','"&nLoctorZ&"','"&nLoctorT&"','"&nRemark&"')"
					InsertOperationSql = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'"&OpCode&"','"&UserCode&"','"&nMENGEWH&"','"&BESTAND&"','"&nLoctorX&"','"&nLoctorY&"','"&nLoctorZ&"','"&nLoctorT&"','"&nVendroNum&"')"
					conn.execute(insertSql)'写入库
					conn.execute(InsertOperationSql)'写日志
					call sussLoctionHref("成功入库：　"&nMENGEWH,"WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
					response.end()
				else	'供应商相同,则更新
					'response.write "<script>alert('执行到5了');</script>"
					BESTAND = nMENGEWH + menge1
					insertSql = "update WH_REC_ITEM set MENGE = '"&BESTAND&"' where ITEM_ID = '"&itemID1&"'"
					'response.write insertSql
					'response.end
					InsertOperationSql = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'"&OpCode&"','"&UserCode&"','"&nMENGEWH&"','"&BESTAND&"','"&LoctorX1&"','"&LoctorY1&"','"&LoctorZ1&"','"&LoctorT1&"','"&vendorNum1&"')"
					conn.execute(insertSql)'写入库
					conn.execute(InsertOperationSql)'写日志
					call sussLoctionHref("已经入到现有库位中!","WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
					response.end()
				end if
			end if
		else	'目标位没有三项符合的物料
			'response.write "<script>alert('执行到6了');</script>"
			RsWh.close
			SqlWh = Sql + " WHERE LOCTOR_X = '"&nLoctorX&"' and LOCTOR_Y = '"&nLoctorY&"' and LOCTOR_Z = '"&nLoctorZ&"'"
			RsWh.open SqlWh,conn,1,1
			if not RsWh.eof then	'目标位有符合三项的物料
				'response.write "<script>alert('执行到7了');</script>"
				call sussLoctionHref("目标库位已经有其他类型物料,请重新选择库位!","WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
				response.end
			else
				'response.write "<script>alert('执行到8了');</script>"
				BESTAND = nMENGEWH
				insertSql = "insert into WH_REC_ITEM(item_name,vendor_num,LAST_IN_DATE,menge,loctor_x,loctor_y,loctor_z,loctor_t,Remark) VALUES('"&ItemName&"','"&nVendroNum&"',sysdate,'"&BESTAND&"','"&nLoctorX&"','"&nLoctorY&"','"&nLoctorZ&"','"&nLoctorT&"','"&nRemark&"')"
				InsertOperationSql = "insert into WH_REC_OPERATION(ITEM_NAME,ADATE,OP_CODE,EMPID,MENGE,BESTAND,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,VENDOR_NAME) VALUES('"&ItemName&"',SYSDATE,'"&OpCode&"','"&UserCode&"','"&nMENGEWH&"','"&BESTAND&"','"&nLoctorX&"','"&nLoctorY&"','"&nLoctorZ&"','"&nLoctorT&"','"&nVendroNum&"')"
				conn.execute(insertSql)'写入库
				conn.execute(InsertOperationSql)'写日志
				call sussLoctionHref("成功入库：　"&nMENGEWH,"WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
				response.end()
			end if
			RsWh.close
		end if
	end if

end if

if BtnRemark = "注" then
	upRemarkSql = "update WH_REC_ITEM set Remark = '"&nRemark&"' where ITEM_ID = '"&itemID&"'"
	conn.execute(upRemarkSql)
	call sussLoctionHref("备注修改成功!","WarehouseIn.asp?SubmitSearch=ok&ItemName="&ItemName)
end if
%>
<table width="1300" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr><td height="20" colspan="2" class="t-t-DarkBlue" align="center">入库操作</td></tr>	
	<tr>
		<td colspan="2">
			<div align="center">
				<form name="form1" method="post" action="WarehouseIn.asp" >
				料号：
				<input name="ItemName" type="text" id="ItemName" value="<%=ItemName%>" onkeyup="lookup(this.value);" onblur="fill();" >
				&nbsp;
				<div class="suggestionsBox" id="suggestions" style="display: none;">
				<div class="suggestionList" id="autoSuggestionsList">
				</div>
			   </div>
				<input type="submit" name="SubmitSearch" value="查询" class="t-b-midautumn" />
			  </td>
				</form>
			</div>
		</td>
	</tr>
</table>
<%if SubmitSearch = "查询" or SubmitSearch = "ok" then

	'生成供应商列表
	ItemInfoSql = "SELECT pm.item_name,pm.description,sm.vendor_num,s.vendor_name,s.vendor_name_alt,unit FROM PRODUCT_MODEL PM LEFT JOIN supplier_material sm on pm.item_name = sm.item_name left join suppliers s on sm.vendor_num = s.vendor_num where pm.item_name = '"&ItemName&"'"
	VendorSelectStr = ""
		rs.open ItemInfoSql,conn,1,1
		if not rs.eof then
			ItemName = rs("item_name")
			ItemDescription = rs("description")
			unit = rs("unit")
			
			VendorSelectStr = ""
			while not rs.eof
				VendorNum = rs("vendor_num")
				VendorNameEn = rs("vendor_name")
				VendorNameCn = rs("vendor_name_alt")
				
				if isnull(VendorNameCn) then
					VendorName = VendorNameEN
				else
					VendorName = VendorNameCn
				end if
					
				VendorSelectStr = VendorSelectStr & "<option value='"&VendorNum&"'>"&VendorName&"</option>"
				rs.movenext
			wend
		end if
		rs.close
	'生成供应商列表
%>
<br>
<table width="1300" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<th width="140">料号</td>
		<th width="50">单位</td>
		<th width="1138">描述</td>
	</tr>
	  <tr class="t-t-List">
		<td width="156"><%=ItemName%></td>
		<td align="center"><%=unit%></td>
		<td align="left"><%=ItemDescription%></td>
	</tr>
</table>
<br>
<table width="1300" border="1" cellpadding="0" cellspacing="0">
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
		SQL = "select wri.item_id,wri.item_name,pm.description,wri.vendor_num,s.vendor_name,s.vendor_name_alt,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name left join suppliers s on wri.vendor_num = s.vendor_num where wri.item_name = '"&ItemName&"' order by LAST_IN_DATE asc"
		
		rs.open SQL,conn,1,1
		'response.write SQL
		if not rs.eof then
			while not rs.eof
			itemName = rs("item_name")
			itemid = rs("item_id")
			vendorNum = rs("vendor_num")
			MENGE = rs("MENGE")
			VendorNameEn = rs("vendor_name")
			VendorNameCn = rs("vendor_name_alt")
			if isnull(VendorNameCn) then
				VendorName = VendorNameEN
			else
				VendorName = VendorNameCn
			end if
			
			DayIn = rs("LAST_IN_DATE")
			InTextStr = "&nbsp;"
			InSubStr = "&nbsp;"
			if month(DayIn) = month(Now()) then
				InTextStr = "<input type='text' name='nMENGEWH' id='nMENGEWH' value='0' size='3'/>"
				InSubStr = "<input type='Submit' name='iowh' value='入库' onClick='return beforeSubmit()'/>"
			end if

			DayOut = rs("LAST_Out_DATE")
				if isnull(DayOut) then DayOut = "&nbsp;"
			LoctorX = rs("LOCTOR_X")
			LoctorY = rs("LOCTOR_Y")
			LoctorZ = rs("LOCTOR_Z")
			LoctorT = rs("LOCTOR_T")
			Remark = rs("Remark")
	%>
	  <form name="form2" method="post" action="WarehouseIn.asp">
		<input type="hidden" name="ItemName" value="<%=ItemName%>"/>
		<tr class="change">
			<td><%=VendorName%>&nbsp;<input type="hidden" name="itemid" value="<%=itemid%>"/></td>
			<td><%=DayIn%></td>
			<td><%=DayOut%></td>
			<td><%=LoctorX%></td>
			<td><%=LoctorY%></td>
			<td><%=LoctorZ%></td>
			<td><%=LoctorT%></td>
			<td><%=Formatnumber(MENGE,-1,-1)%></td>
			<td><%=InTextStr%></td>
			<td><%=InSubStr%></td>
			<td align="left"><input type="text" name="nRemark" value="<%=Remark%>"/><input type="Submit" name="BtnRemark" value="注" /></td>
		</tr>
	  </form>
	  <%
			rs.movenext
			wend
		end if
		rs.close
		Sql = "select item_name,description from product_model WHERE item_name = '"&ItemName&"'"
		rs.open Sql,conn,1,1
		if not rs.eof then
	%>
	</table>
	<br>
	<table width="1300" border="1" cellpadding="0" cellspacing="0" style="background-color:#cccccc;">
	<tr>
	  <td colspan="10" align="center">最新入库</td></tr>
	<tr style="color:#000;">
		<th>供应商</th>
		<th>排</th>
		<th>列</th>
		<th>高</th>
		<th>类型</th>
		<th width='50'>数量</th>
		<th width='40'>操作</th>
		<th width='180'>备注</th>
	  </tr>
	<form name="form2" method="post" action="WarehouseIn.asp">
		<tr>
			<td><input type="hidden" name="ItemName" value="<%=ItemName%>"/><select name="nVendroNum"><%=VendorSelectStr%></select></td>
			<td>
				<select name="nLoctorX">
					<%
						Remark = ""
						set RsLoctor = Server.CreateObject("adodb.recordset")
						SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'X' order by loctor_value"
						RsLoctor.open SqlX,conn,1,1
						LoctorStr = ""
						while not RsLoctor.eof
							LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
							if nLoctorX = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
							LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
							RsLoctor.movenext
						wend
						RsLoctor.close
						set RsLoctor = nothing
						response.write LoctorStr
					%>
				</select>
			</td>
			<td>
				<select name="nLoctorY">
					<%
						set RsLoctor = Server.CreateObject("adodb.recordset")
						SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Y' order by loctor_value"
						RsLoctor.open SqlX,conn,1,1
						LoctorStr = ""
						while not RsLoctor.eof
							LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
							if nLoctorY = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
							LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
							RsLoctor.movenext
						wend
						RsLoctor.close
						set RsLoctor = nothing
						response.write LoctorStr
					%>
				</select>
			</td>
			<td>
				<select name="nLoctorZ">
					<%
						set RsLoctor = Server.CreateObject("adodb.recordset")
						SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Z' order by loctor_value"
						RsLoctor.open SqlX,conn,1,1
						LoctorStr = ""
						while not RsLoctor.eof
							LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
							if nLoctorZ = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
							LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
							RsLoctor.movenext
						wend
						RsLoctor.close
						set RsLoctor = nothing
						response.write LoctorStr
					%>
				</select>
			</td>
			<td>
				<select name="nLoctorT">
					<option value="BULK">BULK</option>
					<%
						set RsLoctor = Server.CreateObject("adodb.recordset")
						SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'T' and LOCTOR_TYPE = '1' order by loctor_value"
						RsLoctor.open SqlX,conn,1,1
						LoctorStr = ""
						while not RsLoctor.eof
							LoctorStr = LoctorStr & "<option value='"&RsLoctor("loctor_value")&"'" 
							if nLoctorT = RsLoctor("loctor_value") then LoctorStr = LoctorStr & "selected"
							LoctorStr = LoctorStr & " >"&RsLoctor("loctor_value")&"</option>"
							RsLoctor.movenext
						wend
						RsLoctor.close
						set RsLoctor = nothing
						response.write LoctorStr
					%>
				</select>
			</td>
			<td><input type="text" name="nMENGEWH" id="nMENGEWH" value="0" size="3"/></td>
			<td><input type="Submit" name="iowh" value="新入" /></td>
			<td align="left"><input type="text" name="nRemark" value="<%=Remark%>"/></td>
		</tr>
	</form>
	<%end if%>
	</table>
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