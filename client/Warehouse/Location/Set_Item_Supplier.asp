<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_Set"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�Ϻ��빩Ӧ�̰�</title>
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

</style>
</head>
<body style="margin:20px;background-color:#339966;">
<center>
<%
SubmitSearch = request("SubmitSearch")
AddVendor = request("AddVendor")
ItemName = UCASE(trim(request("ItemName")))
VendorNum = trim(request("VendorNum"))

if AddVendor = "��" then
	if not IsNumeric(VendorNum) then call sussLoctionHref("��Ӧ�̱�ű���Ϊ����!","Set_Item_Supplier.asp?SubmitSearch=ok&ItemName="&ItemName)
	sql = "SELECT vendor_num FROM suppliers where vendor_num = '"&VendorNum&"'"
	rs.open sql,conn,1,1
	if rs.eof then
		rs.close
		call sussLoctionHref("û���ҵ��˹�Ӧ��,��ȷ�Ϻ�����!","Set_Item_Supplier.asp?SubmitSearch=ok&ItemName="&ItemName)
		response.end
	else
		rs.close
		sql = "SELECT * FROM supplier_material where item_name = '"&ItemName&"' and vendor_num = '"&VendorNum&"'"
		rs.open sql,conn,1,1
		if not rs.eof then
			call sussLoctionHref("�Ϻ��Ѿ��빩Ӧ�̽��й���!","Set_Item_Supplier.asp?SubmitSearch=ok&ItemName="&ItemName)
			response.end
		else
			cfgVendroNum = right(VendorNum,5)
			insertSql = "insert into supplier_material(vendor_num,CFG_VENDOR_NUM,item_name,updateuser,updatedate) values('"&VendorNum&"','"&cfgVendroNum&"','"&ItemName&"','"&UserCode&"',sysdate)"
			'response.write insertSql
			'response.end()
			conn.execute(insertSql)
			call sussLoctionHref("�Ѿ��ɹ���!","Set_Item_Supplier.asp?SubmitSearch=ok&ItemName="&ItemName)
			response.end
		end if
		rs.close
	end if
end if

' if AddVendor = "ɾ��" then
	' delSql = "delete from supplier_material where item_name = '"&ItemName&"' and vendor_num = '"&VendorNum&"'"
	' conn.execute(delSql)
	' SubmitSearch = "ok"
' end if
%>
<form name="form1" method="post" action="Set_Item_Supplier.asp" >
<table width="1300" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue">
			<div align="center">�Ϻ��빩Ӧ�̰�</div>
		</td>
	</tr>	
	<tr style="background:#cccccc;">
		<td>
		�Ϻţ�<input type="text" name="ItemName" value="<%=ItemName%>" id="PART_NUMBER" onkeyup="lookup(this.value);" onblur="fill();"/>&nbsp;
			<input type="submit" name="SubmitSearch" value="��ѯ�Ϻ�" />&nbsp;
			<div class="suggestionsBox" id="suggestions" style="display: none;">
			<div class="suggestionList" id="autoSuggestionsList">&nbsp;</div>
		</td>
	</tr>
</table>
</form>
<%
if SubmitSearch = "��ѯ�Ϻ�" or SubmitSearch = "ok" then
	SQL = "select ITEM_NAME,DESCRIPTION from product_model pm where pm.item_name = '"&ItemName&"'"
	'response.write SQL
	rs.open SQL,conn,1,1
	if not rs.eof then
		ItemName = rs("Item_name")
		description = rs("description")
		rs.close
		SQL = "select sm.item_name,pm.description,sm.vendor_num,s.vendor_name,s.vendor_name_alt from supplier_material sm left join product_model pm on sm.item_name = pm.item_name left join suppliers s on sm.vendor_num = s.vendor_num where sm.item_name = '"&ItemName&"'"
		'response.write SQL
		rs.open SQL,conn,1,1
		Nr = 1
%>
<table width='1300' border='1' cellpadding='0' cellspacing='0' style="background:#ffffff;">
	<tr class="t-t-DarkGray"><th width='140'>�Ϻ�</th><th colspan="4">����</th></tr>
<tr class="t-t-List">
	<td><%=ItemName%></td>
	<td colspan="4"><%=description%></td>
</tr>
</table>
<br>
<table width='1300' border='1' cellpadding='0' cellspacing='0' style="background:#ffffff;">
<tr class="t-t-DarkGray"><th width='140'>��Ӧ�̱��</th><th>��Ӧ��(Ӣ��)</th><th>��Ӧ��(����)</th></tr>
<%
		while not rs.eof
		VendorNum = rs("vendor_num")
		VendorNameEn = rs("vendor_name")
		VendorNameCn = rs("vendor_name_alt")
%>
	<form name="formDel" method="post" action="Set_Item_Supplier.asp">
		<tr class="t-t-List">
			<td><%=VendorNum%><input type="hidden" name="ItemName" value="<%=ItemName%>" /><input type="hidden" name="VendorNum" value="<%=VendorNum%>" /></td>
			<td><%=VendorNameEn%></td>
			<td><%=VendorNameCn%>&nbsp;</td>
		</tr>
	</form>
<%
		rs.movenext
		wend
	end if
	rs.close
%>
</table>
<br>
<form name="form2" method="post" action="Set_Item_Supplier.asp">
<input type="hidden" name="ItemName" value="<%=ItemName%>" />
<table width='1300' border='1' cellpadding='0' cellspacing='0' style="background:#ffffff;">
<tr class="t-t-DarkGray"><th>��Ӧ�̱��</th><th>����</th></tr>
	<tr class="t-t-List">
		<td><input type="button" value="ǰȥ���ұ��" onClick="javascript:window.open('Set_SearchSupplierNum.asp')">&nbsp;<input type="text" name="VendorNum" /></td>
		<td><input type="Submit" name="AddVendor" value="��"/></td>
	</tr>
</table>
</form>
<%	
end if
%>
</table>
  <table width="1300" border="1" cellspacing="0" cellpadding="0">
    <tr style="background:#cccccc;">
      <td align="center">
      <input type="button" value="�޸�����" onClick="window.location.reload('../../Functions/UserPass_Modify.asp');" />
      ��������
      <input type="button" value="Close�ر�" onClick="window.location.reload('../../Functions/User_Logout.asp');" /></td>
    </tr>
  </table>
</center>

</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->