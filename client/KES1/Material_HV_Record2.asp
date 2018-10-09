<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<%
LabelId = request("LabelId")
LotNo = request("LotNo")
rtnValue=""

'get job number by tray id
if LabelId = "" or LotNo="" then 
	rtnValue = "Error LotNO  is Null.\n 物料信息不能为空。"

else
    LabelId=trim(request("labelId"))
	Job_number=session("JOB_NUMBER")
	LINE_NAME=session("LINE_NAME")
	 if Job_number="" then
	response.Redirect("/KES1/Station1.asp?errorstring=Idle time is too long.<br>空闲时间超出！")
	 end if
	Sheet_number=cstr(session("SHEET_NUMBER"))	
	CFG_Vendor_num=left(labelId,5)
	ITEM_N=mid(labelId,6,5)
	Lot_num=clng(right(labelId,8))
	LotNo=trim(request("LotNo"))
	EQUIPMENT=trim(request("EQUIPMENT"))
	Operator_user=session("code")
	updatetime=trim(request("Update_time"))
	update_time=now()
	sql="select * from supplier_material where CFG_VENDOR_NUM='"&CFG_Vendor_num&"' and substr(item_name,8,5)='"&ITEM_N&"' "
	rs.open sql,conn,1,3
	if not rs.eof then	
	Vendor_num=rs("Vendor_num")
	Item_name=rs("ITEM_NAME")	
	else
	rtnValue="没有查到此供应商，请联系主管"	
	end if
	rs.close
	
	
	if rtnValue="" then
    sql= "insert into material_HV_RECORD (job_number,sheet_number,VENDOR_NUM,ITEM_NAME,LOT_NUM,LOT_NO,LINE_NAME,OPERATOR_USER,UPDATE_TIME,EQUIPMENT) values ('"&job_number&"','"&Sheet_number&"','"&Vendor_num&"','"&Item_name&"','"&Lot_num&"','"&LotNo&"','"&LINE_NAME&"','"&Operator_user&"','"&update_time&"','"&EQUIPMENT&"')"
    conn.execute(sql)
	end if

'response.Redirect("Material_HV_Record.asp")
	
end if
%>
<script type="text/javascript">
	window.returnValue="<%=rtnValue%>";
	window.close();
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
