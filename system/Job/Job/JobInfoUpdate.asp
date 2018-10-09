<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
jobnumber=request.QueryString("jobnumber")
start_quantity=0
SQLPR="select M.LineCode,M.StartQuantity,U.CODE as CREATE_CODE,U.NAME as CREATE_NAME from tbl_MES_LotMaster M left join tbl_User_ID U on M.CREATEPERSON=U.ID where WipEntityName='"&jobnumber&"'"
rsPR.open SQLPR,connPR,1,3
if not rsPR.eof then
'start_quantity=cdbl(rsPR("StartQuantity"))
create_code=rsPR("CREATE_CODE")
create_name=rsPR("CREATE_NAME")
end if
rsPR.close

word=""
SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
'	if isnull(rs("START_QUANTITY")) or rs("START_QUANTITY")="" then
'		rs("START_QUANTITY")=start_quantity
'		word=word&"Start quantity <"&rs("START_QUANTITY")&" - "&start_quantity&"> is updated!\n"
'	else
'		if cdbl(rs("START_QUANTITY"))<>cdbl(start_quantity) then
'		rs("START_QUANTITY")=start_quantity
'		word=word&"Start quantity <"&rs("START_QUANTITY")&" - "&start_quantity&"> is updated!\n"
'		else
'		word=word&"Start quantity <"&rs("START_QUANTITY")&" - "&start_quantity&"> is not updated!\n"
'		end if
'	end if
'	if cdbl(rs("START_QUANTITY"))-cdbl(rs("FINAL_GOOD_QUANTITY"))<>cdbl(rs("FINAL_SCRAP_QUANTITY")) then
'		rs("FINAL_SCRAP_QUANTITY")=cdbl(rs("START_QUANTITY"))-cdbl(rs("FINAL_GOOD_QUANTITY"))
'	end if
	if cdbl(rs("FINAL_GOOD_QUANTITY"))+cdbl(rs("FINAL_SCRAP_QUANTITY"))=cdbl(rs("START_QUANTITY")) then
		rs("STORE_STATUS")=1
		
		word=word&"Job is finished!\n"
	else
		rs("STORE_STATUS")=0
		word=word&"Job is not finished!\n"
	end if
	rs("FINAL_YIELD")=cdbl(rs("FINAL_GOOD_QUANTITY"))/cdbl(rs("START_QUANTITY"))
	rs("ASSEMBLY_YIELD")=cdbl(rs("ASSEMBLY_GOOD_QUANTITY"))/cdbl(rs("START_QUANTITY"))
	if isnull(rs("CREATE_CODE")) or rs("CREATE_CODE")="" then
		rs("CREATE_CODE")=create_code
		word=word&"Create code <"&rs("CREATE_CODE")&" - "&create_code&"> is updated!\n"
	else
		if rs("CREATE_CODE")<>create_code then
		rs("CREATE_CODE")=create_code
		word=word&"Create code <"&rs("CREATE_CODE")&" - "&create_code&"> is updated!\n"
		else
		word=word&"Create code <"&rs("CREATE_CODE")&" - "&create_code&"> is not updated!\n"
		end if
	end if
	if isnull(rs("CREATE_NAME")) or rs("CREATE_NAME")="" then
		rs("CREATE_NAME")=create_name
		word=word&"Create name <"&rs("CREATE_NAME")&" - "&create_name&"> is updated!\n"
	else
		if rs("CREATE_NAME")<>create_name then
		rs("CREATE_NAME")=create_name
		word=word&"Create name <"&rs("CREATE_NAME")&" - "&create_name&"> is updated!\n"
		else
		word=word&"Create name <"&rs("CREATE_NAME")&" - "&create_name&"> is not updated!\n"
		end if
	end if
	'if isnull(rs("INPUT_TIME")) or rs("INPUT_TIME")="" then
		set rsJ=server.CreateObject("adodb.recordset")
		SQLJ="select START_TIME from JOB where ROWNUM=1 and JOB_NUMBER='"&jobnumber&"' order by START_TIME"
		rsJ.open SQLJ,conn,1,3
		if not rsJ.eof then
			input_time=rsJ("START_TIME")
			rs("INPUT_TIME")=input_time
			word=word&"Input time <"&input_time&"> is updated!\n"
		end if
		rsJ.close
		set rsJ=nothing
	'end if
	 if isnull(rs("FIRST_STORE_TIME")) or rs("FIRST_STORE_TIME")="" then
		set rsJ=server.CreateObject("adodb.recordset")
		SQLJ="select STORE_TIME from JOB_MASTER_STORE where ROWNUM=1 and JOB_NUMBER='"&jobnumber&"' order by STORE_TIME"
		rsJ.open SQLJ,conn,1,3
		if not rsJ.eof then
			create_time=rsJ("STORE_TIME")
			rs("FIRST_STORE_TIME")=create_time
			word=word&"Create time <"&create_time&"> is updated!\n"
		end if
		rsJ.close
		set rsJ=nothing
	end if
	'update assembly_input and yield
		set rsJ=server.CreateObject("adodb.recordset")
		SQLJ="select sum(JOB_START_QUANTITY) as JOB_START_QUANTITY,sum(JOB_GOOD_QUANTITY) as JOB_GOOD_QUANTITY from JOB where JOB_NUMBER='"&jobnumber&"'"
		rsJ.open SQLJ,conn,1,3
		if not rsJ.eof then
			rs("ASSEMBLY_INPUT_QUANTITY")=rsJ("JOB_START_QUANTITY")
			rs("ASSEMBLY_GOOD_QUANTITY")=rsJ("JOB_GOOD_QUANTITY")
			rs("ASSEMBLY_YIELD")=csng(rsJ("JOB_GOOD_QUANTITY"))/csng(rsJ("JOB_START_QUANTITY"))
			word=word&"Assembly Quantity <"&create_time&"> is updated!\n"
		end if
		rsJ.close
		set rsJ=nothing
	rs.update
end if
rs.close
action="location.href='"&beforepath&"'"
'response.Write(action)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->