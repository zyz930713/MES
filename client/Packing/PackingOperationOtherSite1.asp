<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->

<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
'Response.Charset="utf-8"
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request.Form("asynid")
end if
if len(asynid)>0 then
	set rsAsyn=server.CreateObject("adodb.recordset")
	
	dim member
	set member=jsObject()
	select case asynid
	
		case "2"'Package Save			
			opcode=trim(request.Form("txt_opcode"))
			boxid=trim(request.Form("txt_boxid"))
			customer=trim(request.Form("txt_customer"))
			remarks=trim(request.Form("txt_remarks"))
			packType=trim(request("txt_pack_type"))
			packLine=trim(request.Form("txt_pack_line"))
			partnumber=trim(request.Form("txt_part"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			custId=trim(request.Form("txt_custid"))		
			inspectionpqc=trim(request.Form("txt_INSPECTIONPQC"))	
			Qty=trim(request.Form("txt_boxsize"))
			job_number=trim(request.Form("txt_Job_Number"))
			bind_box_id=trim(request.Form("txt_Bind_Box_Id"))
			conn.beginTrans
			
				
			SQLH="SELECT WipEntityName,PRIMARYITEMDESC FROM tbl_MES_LotMaster  where WipEntityName='"&job_number&"'"			
			rsTicket.open SQLH,connTicket,1,3
			if  rsTicket.bof and rsTicket.eof then
			  member("error")="This Job ("&job_number&") does not  in system.|该工单("&job_number&")在系统中不存在，请确认工单号。"		
			else
			PRIMARYITEMDESC=rsTicket("PRIMARYITEMDESC")
			WipEntityName=rsTicket("WipEntityName")
			end if	
			rsTicket.close				
			packLine=mid(job_number,3,4)
			boxidLine=mid(boxid,3,4)	
			'check partial packing
			if member("error")="" then
			   if PRIMARYITEMDESC<>partnumber then
			    member("error")="工单号与成品料号不匹配，请不要乱输入工单号与料号。"				   
			   end if
			   if 	packLine<>	boxidLine then			   
			    member("error")="工单线别与箱号线别不一致。"				
			   end if	
			end if	
				
			if member("error")="" then
					
			sql="select * from job_pack_detail where box_id='"&boxid&"' and job_number='"&job_number&"'"		
			
				rs.open sql,conn,1,3
					if  rs.eof then
					
					rs.addnew
					rs("box_id")=boxid
					rs("job_number")=job_number	
					end if		
				
					rs("packed_qty")=cint(Qty)
					rs("part_number")=partnumber
					rs("customer")=customer
					rs("packed_user")=opcode
					rs("packed_time")=now()
					rs("remarks")=remarks
					rs("packed_line")=packLine
					rs("prod_shift")=prodShift
					rs("supplier")=supplier
					rs("customer_label_id")=custId
					rs("inspectionpqc")=inspectionpqc
					rs("BIND_BOX_ID")=bind_box_id
					rs.update
					rs.close				
			
			end if
		
			
			if err.number <> 0 then 
				conn.rollbackTrans   '对已执行的操作回滚
			
				member("error")="Save failed.保存失败.|"&err.description
		
				response.Write toJSON(member)
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="Save succesful.|保存成功."
				response.Write toJSON(member)
			end if	
		case "1"'label print
	
		case "3" 'check 2D code
		
			
	end select
	set rsAsyn=nothing 	
end if


%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->