<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
pagename="DeleteStore.asp"
nid=trim(request.QueryString("nid"))
valid=false
diff=0
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
'delete store info in job_master_store
SQL="select * from JOB_MASTER_STORE_PRE where NID='"&nid&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	jobnumber=rs("JOB_NUMBER")
	input_quantity=cint(rs("INPUT_QUANTITY"))
	store_quantity=cint(rs("STORE_QUANTITY"))
	inspect_quantity=cint(rs("INSPECT_QUANTITY"))
	store_type=rs("STORE_TYPE")
	sub_job_numbers=rs("SUB_JOB_NUMBERS")
	part_number_tag=rs("PART_NUMBER_TAG")
	factory_id=rs("FACTORY_ID")
	line_name=rs("LINE_NAME")
	note=rs("NOTE")
	store_code=rs("STORE_CODE")
	store_time=rs("STORE_TIME")
	rs.delete
	valid=true
	word="删除成功，请重新入库！"
else
	valid=false
	word="记录不存在，无法删除！"
end if
rs.close

if valid=true then
	'save a change history record in history table
	SQL="select * from JOB_MASTER_STORE_CHANGE"
	rs.open SQL,conn,1,3
	rs.addnew
	rs("STORE_NID")=nid
	rs("CHANGE_TYPE")="2"'1=change;2=delete
	rs("CHANGE_CODE")=request.QueryString("changecode")
	rs("CHANGE_REASON")=request.QueryString("changereason")
	rs("CHANGE_TIME")=now()
	rs("JOB_NUMBER")=jobnumber
	rs("OLD_INPUT_QUANTITY")=input_quantity
	rs("OLD_STORE_QUANTITY")=store_quantity
	rs("SUB_JOB_NUMBERS")=sub_job_numbers
	rs("PART_NUMBER_TAG")=part_number_tag
	rs("FACTORY_ID")=factory_id
	rs("LINE_NAME")=line_name
	rs("OLD_STORE_TYPE")=store_type
	rs("OLD_NOTE")=note
	rs("OLD_STORE_CODE")=store_code
	rs("OLD_STORE_TIME")=store_time
	rs.update
	rs.close
	
	'Get Scrap info
	SQL="SELECT DECODE(SUM(SCRAP_QUANTITY),NULL,0,SUM(SCRAP_QUANTITY)) FROM JOB_MASTER_SCRAP_PRE WHERE STORE_NID='"+nid+"'"
	set rsScrap=server.createobject("adodb.recordset")
	preScrapQty=0
	rsScrap.open SQL,conn,1,3
	IF(rsScrap.RECORDCOUNT>0)then
		preScrapQty=cint(rsScrap(0))
	end if 
	
	SQL="DELETE JOB_MASTER_SCRAP_PRE WHERE STORE_NID='"+nid+"'"
	set rsScrap2=server.createobject("adodb.recordset")
	rsScrap2.OPEN SQL,conn,1,3
	'update store info in job_master
	SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then
		final_good_quantity=cint(rs("FINAL_GOOD_QUANTITY"))-store_quantity
		CONFIRM_good_quantity=cint(rs("CONFIRM_GOOD_QUANTITY"))-inspect_quantity
		final_scrap_quantity=cint(rs("FINAL_SCRAP_QUANTITY"))-preScrapQty		
		rs("CONFIRM_GOOD_QUANTITY")=CONFIRM_good_quantity
		rs("FINAL_SCRAP_QUANTITY")=final_scrap_quantity
		rs("FINAL_GOOD_QUANTITY")=final_good_quantity		
		rs("FINAL_INPUT_QUANTITY")=final_good_quantity + final_good_quantity
		if final_good_quantity = 0 then
			rs("FINAL_YIELD")=0
		else
			rs("FINAL_YIELD")=final_good_quantity/(final_good_quantity + final_good_quantity)
		end if
		rs("STORE_STATUS")="0"
		rs.update
	end if
	rs.close

	'Update job Inventory
	SQL="select job_number,part_number,good_qty,scrap_qty,lm_time,lm_user from job_inventory where job_number='"&jobnumber&"'"
	rs.open SQL,conn,1,3
	if not rs.eof then		
		rs("good_qty")=cint(rs("good_qty"))-store_quantity
		rs("scrap_qty")=cint(rs("scrap_qty"))-preScrapQty
		rs("lm_time")=now()
		rs("lm_user")=request.QueryString("changecode")
		rs.update
	end if	
	rs.close
	
end if
action="location.href='"&beforepath&"'"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
</head>

<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->