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
<body style="margin:20px;background-color:#339966;">
<center>
<%
SubmitSearch = request("SubmitSearch")
ItemName = UCASE(trim(request("ItemName")))
'MENGEWH = request("MENGEWH")
LoctorX = request("LoctorX")
LoctorY = request("LoctorY")
LoctorZ = request("LoctorZ")
LoctorT = request("LoctorT")
Remark = request("Remark")
BtnRemark = request("BtnRemark")

if BtnRemark = "注" then
	conn.execute("update WH_REC_ITEM set Remark = '"&Remark&"' where item_name = '"&ItemName&"' and loctor_x = '"&LoctorX&"' and loctor_y = '"&LoctorY&"' and loctor_z = '"&LoctorZ&"' and loctor_t = '"&LoctorT&"'")
	call sussLoctionHref("备注修改成功!","WarehouseMove.asp?SubmitSearch=ok&ItemName="&ItemName&"&LoctorT=ALL")
end if
%>
<table width="1202" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">移库操作</div>
		</td>
	</tr>	
	<tr>
		<td>
			<div align="center">
				<form name="form1" method="post" action="WarehouseMove.asp" >
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

<br>
<%
if SubmitSearch = "查询" or SubmitSearch = "ok" then

	' SQL = "select wri.item_name,pm.description,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name where wri.item_name = '"&ItemName&"'"
	SQL = "select wri.item_id,wri.item_name,pm.description,wri.vendor_num,s.vendor_name,s.vendor_name_alt,LAST_OUT_DATE,LAST_IN_DATE,MENGE,LOCTOR_X,LOCTOR_Y,LOCTOR_Z,LOCTOR_T,Remark from wh_rec_item wri left join product_model PM ON wri.item_name = pm.item_name left join suppliers s on wri.vendor_num = s.vendor_num where wri.item_name = '"&ItemName&"'"
	IF LoctorT <> "ALL" THEN SQL = SQL & " and LOCTOR_T = '"&LoctorT&"'"
	SQL = SQL & " order by LAST_IN_DATE asc"
	'response.write SQL
	rs.open SQL,conn,1,1
	
	if not rs.eof then
		ItemName = rs("Item_name")
		description = rs("description")
%>
	<table width="1202" border="1" cellpadding="0" cellspacing="0" style="background-color:#cccccc;">
		<tr style="color:#000;font-size:18px;">
			<td><%=ItemName%></td>
			<td colspan="10" align="left"><%=description%></td>
		</tr>
		<tr>
			<td colspan="11">&nbsp;</td>
		</tr>
	  <tr>
		<th>供应商</th>
		<th>入库日期</th>
		<th>出库日期</th>
		<th>排</th>
		<th>列</th>
		<th>高</th>
		<th>类型</th>
		<th>库存</th>
		<th width='60'>操作</th>
	  </tr>
<%
		while not rs.eof
		itemId = rs("item_id")
		MENGE = rs("MENGE")
		vendorNum = rs("vendor_num")
		VendorNameEn = rs("vendor_name")
		VendorNameCn = rs("vendor_name_alt")
		if isnull(VendorNameCn) then
			VendorName = VendorNameEN
		else
			VendorName = VendorNameCn
		end if
		DayIn = rs("LAST_IN_DATE")
		DayOut = rs("LAST_Out_DATE")
			if isnull(DayOut) then DayOut = "　"
		LoctorX = rs("LOCTOR_X")
		LoctorY = rs("LOCTOR_Y")
		LoctorZ = rs("LOCTOR_Z")
		LoctorT = rs("LOCTOR_T")
		Remark = rs("Remark")

		%>
		<form name="formS" method="post" action="WarehouseMoveTo.asp" target="main">
		<input type="hidden" name="ItemName" size="2" value="<%=ItemName%>"/>
			<tr class="change">
				<td><%=VendorName%>&nbsp;<input type="hidden" name="SitemId" size="2" value="<%=itemId%>"/><input type="hidden" name="SvendorNum" size="2" value="<%=vendorNum%>"/></td>
				<td><%=DayIn%><input type="hidden" name="SDayIn" size="2" value="<%=DayIn%>"/></td>
				<td><%=DayOut%></td>
				<td><%=LoctorX%><input type="hidden" name="SLoctorX" size="2" value="<%=LoctorX%>"/></td>
				<td><%=LoctorY%><input type="hidden" name="SLoctorY" size="2" value="<%=LoctorY%>"/></td>
				<td><%=LoctorZ%><input type="hidden" name="SLoctorZ" size="2" value="<%=LoctorZ%>"/></td>
				<td><%=LoctorT%><input type="hidden" name="SLoctorT" size="2" value="<%=LoctorT%>"/></td>
				<td><%=MENGE%><input type="hidden" name="SMENGEWH" size="2" value="<%=MENGE%>"/></td>
				<td><input type='Submit' name='IOWh' value='移库' onClick='return beforeSubmit()'/></td>
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
<div id="content">
<iframe src="" id="main" name="main" style="z-index:-1;background-color:#339966;" marginheight="0" frameborder="0" scrolling="no" vspace="0" hspace="0" marginwidth="0" width="1202" height="170">
</iframe>
<script type="text/javascript" language="javascript"> 
function iFrameHeight() {
var ifm= document.getElementById("main");
var subWeb = document.frames ? document.frames["main"].document : ifm.contentDocument;
if(ifm != null && subWeb != null) {
ifm.height = subWeb.body.scrollHeight;
}
}
</script>
</div>
<br>
  <table width="1202" border="1" cellspacing="0" cellpadding="0">
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