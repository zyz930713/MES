<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_IO_List"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>库存记录查询</title>
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
 function mouseOver(obj){
  if(obj.className="css1")
	 obj.className="css2";
}
function mouseOut(obj){
  if(obj.className="css2")
	 obj.className="css1";
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
<!--
 .css1 {background-color: #cccccc;}
 .css2 {background-color: #0aa099;}
-->
</style>
</head>
<body style="margin:20px;background-color:#339966;">
<center>
<%
SubmitSearch = request("SubmitSearch")
ItemName = UCASE(trim(request("ItemName")))
LoctorX = request("LoctorX")
LoctorY = request("LoctorY")
LoctorZ = request("LoctorZ")
LoctorT = request("LoctorT")
'response.write LoctorT
%>
<table width="1300" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">库存查询</div>
		</td>
	</tr>	
	<tr>
		<td>
			<div align="center">
				<form name="form1" method="post" action="WarehouseItem.asp" >
				料号：<input type="text" name="ItemName" value="<%=ItemName%>" id="PART_NUMBER" onkeyup="lookup(this.value);" onblur="fill();"/>&nbsp;
				排：<select name="LoctorX">
						<%
							Remark = ""
							set RsLoctor = Server.CreateObject("adodb.recordset")
							SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'X' order by loctor_value"
							RsLoctor.open SqlX,conn,1,1
							response.write "<option value='%' >all</option>"
							while not RsLoctor.eof
								loctor_value = RsLoctor("loctor_value")
								if loctor_value = LoctorX then
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
					</select>
				列：<select name="LoctorY">
						<%
							set RsLoctor = Server.CreateObject("adodb.recordset")
							SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Y' order by loctor_value"
							RsLoctor.open SqlX,conn,1,1
							response.write "<option value='%' >all</option>"
							while not RsLoctor.eof
								loctor_value = RsLoctor("loctor_value")
								if loctor_value = LoctorY then
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
					</select>
				高：<select name="LoctorZ">
						<%
							set RsLoctor = Server.CreateObject("adodb.recordset")
							SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'Z' order by loctor_value"
							RsLoctor.open SqlX,conn,1,1
							response.write "<option value='%' >all</option>"
							while not RsLoctor.eof
								loctor_value = RsLoctor("loctor_value")
								if loctor_value = LoctorZ then
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
					</select>
				类型：<select name="LoctorT">
							<%
								set RsLoctor = Server.CreateObject("adodb.recordset")
								SqlX = "select loctor_value from wh_cfg_loctor where loctor_name = 'T' order by loctor_value"
								RsLoctor.open SqlX,conn,1,1
								dim SelectStr
								response.write "<option value='%' >所有类型</option>"
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
					<input type="submit" name="SubmitSearch" value="ToExcel" onClick="javascript:window.open('../../Functions/Export_Excel.asp')" />
					<div class="suggestionsBox" id="suggestions" style="display: none;">
					<div class="suggestionList" id="autoSuggestionsList">&nbsp;</div>
				</form>
			</div>
		</td>
	</tr>
</table>
<br>
<%
if SubmitSearch = "查询" or SubmitSearch = "ok" or SubmitSearch = "ToExcel" then
	SQL = "select wri.item_name,pm.description,wri.vendor_num,s.vendor_name,s.vendor_name_alt,unit,to_char(LAST_OUT_DATE,'yyyy-MM-dd') LAST_OUT_DATE,to_char(LAST_IN_DATE,'yyyy-MM-dd') LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name left join suppliers s on wri.vendor_num = s.vendor_num where wri.item_name like '%"&ItemName&"%' and LOCTOR_X like '"&LoctorX&"' and LOCTOR_Y like '"&LoctorY&"' and LOCTOR_Z like '"&LoctorZ&"' and LOCTOR_T like '"&LoctorT&"'"
	'SQL1 = "select wri.item_name,pm.description,s.vendor_name,s.vendor_name_alt,LAST_OUT_DATE,LAST_IN_DATE,MENGE QTY,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name left join suppliers s on wri.vendor_num = s.vendor_num where wri.item_name like '%"&ItemName&"%' and LOCTOR_X like '"&LoctorX&"' and LOCTOR_Y like '"&LoctorY&"' and LOCTOR_Z like '"&LoctorZ&"' and LOCTOR_T like '"&LoctorT&"'"
	'response.write SQL
	session("SQL") = SQL
	session("ExcelName") = "Item"
	rs.open SQL,conn,1,1
	if not rs.eof then
%>
	<table width="1300" border="1" cellpadding="0" cellspacing="0">
	<tr style="background:#cccccc;"><th width="110">料号</th><th>描述</th><th width="50">单位</td><th>供应商</th><th>数量</th><th>最后入库日期</th><th>最后出库日期</th><th>排</th><th>列</th><th>高</th><th>类型</th></tr>
<%
	while not rs.eof
		ItemName = rs("Item_name")
		unit = rs("unit")
		description = rs("description")
		VendorName = rs("VENDOR_Name")
		MENGE = rs("MENGE")
		DayIn = rs("LAST_IN_DATE")
			if isnull(DayIn) then DayIn = "&nbsp;"
		DayOut = rs("LAST_Out_DATE")
			if isnull(DayOut) then DayOut = "&nbsp;"
		LoctorX = rs("LOCTOR_X")
		LoctorY = rs("LOCTOR_Y")
		LoctorZ = rs("LOCTOR_Z")
		LoctorT = rs("LOCTOR_T")
	%>
		<tr onMouseOver="mouseOver(this)" onMouseOut="mouseOut(this)">
			<td><%=ItemName%></td>
			<td align="left"><%=description%></td>
			<td align="center"><%=unit%></td>
			<td align="left"><%=VendorName%>&nbsp;</td>
			<td><%=MENGE%></td>
			<td><%=DayIn%></td>
			<td><%=DayOut%></td>
			<td><%=LoctorX%></td>
			<td><%=LoctorY%></td>
			<td><%=LoctorZ%></td>
			<td><%=LoctorT%></td>
		</tr>
		<%rs.movenext
		wend
		response.write "</table>"
	end if
	rs.close
end if

%>
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