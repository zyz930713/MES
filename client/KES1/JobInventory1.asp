<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<%
'On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
Response.Charset="gb2312"
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request.Form("asynid")
end if
if len(asynid)>0 then
	set rsAsyn=server.CreateObject("adodb.recordset")
	set rsAsyn1=server.CreateObject("adodb.recordset")
	dim member
	set member=jsObject()
	select case asynid
		case "1"'JOB_INVENTORY		 根据jobnumber模糊查询
			jobno=request.QueryString("txt_jobno")
			partno=request.QueryString("txt_partno")
			where=""
			if len(jobno)>0 then
				where=where&" and job_number like '"&jobno&"%'"
			end if
			if len(partno)>0 then
				where=where&" and part_number like '"&partno&"%'"
			end if
			'只显示good qty大于0的记录
			sql="select * from JOB_INVENTORY where good_qty>0 "&where
			response.Write QueryToJSON(conn, sql).Flush
		case "2"''根据工单获取包装记录
			jobno=request.QueryString("txt_jobno")
			sql="select * from JOB_PACK_DETAIL where job_number ='"&jobno&"' order by box_id"
			response.Write QueryToJSON(conn, sql).Flush			
	end select 	
end if
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->