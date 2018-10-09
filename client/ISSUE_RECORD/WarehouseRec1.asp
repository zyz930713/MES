<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request("asynid")
end if
errMsg=""
dim member
set member=jsObject()
if len(asynid)>0 then	
	select case asynid
		case "2" 'Save data and print pallet label
			opcode=request("txt_opcode")
						
			boxIdList=request("hidBoxIdList")
			newUnitQty=request("hidNewUnitQty")
			WHlocation=request("WHlocation")
			conn.beginTrans
			
			'generate pallet id
			
			if boxIdList="" then'reprint pallet label
				sql="select user_code from users where user_code='"&opcode&"' and instr(roles_id,'REPRINT_PALLET_LABEL')>0 "
				rs.open sql,conn,1,3
				if rs.eof then
					errMsg="You are not authorized to reprint pallet.|您没有权限进行此操作。"
				end if
				rs.close				
			end if
			if boxIdList<>"" then
				sql="update job_pack_detail set WHREC_user='"&opcode&"',WHREC_time=sysdate,GET_USER='',GET_TIME='',WHREC_JSQ=decode(WHREC_jsq,null,1,WHREC_jSQ+1),BOXIDSTATUS='"&WHlocation&"' where box_id in ('"&replace(boxIdList,",","','")&"')"
				conn.execute(sql)
				
					
			end if
			if err.number <> 0 then 				
				errMsg="确认入库失败|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="确认入库完成."
			end if			
			'response.Write toJSON(member)		
		case "1" 'input box id
			boxId=request("txt_inputboxid")
			planPart=request("plan_part")
			palletId=request("txt_pallet") 
			boxId1= request("txt_boxid1")
			planId=request("slt_plan")
			newUnitQty=request("hidNewUnitQty")
			WHlocation=request("WHlocation")
			WH=left(WHlocation,2)
            if WH<>"MR" then
				if WH<> left(boxId,2) then
				errMsg="箱号与库位不匹配！"
				end if
			end if 
			unitQty=0
			boxQty=1
			sql="select part_number,sum(packed_qty) as packed_qty,WHREC_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,WHREC_TIME"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This box id("&boxId&") dose not exist.|该箱号("&boxId&")不存在."
			else
				
				if  not isnull(rs("WHREC_TIME")) then
				WHREC_TIME=rs("WHREC_TIME")
				errMsg="该箱号("&boxId&")在("&WHREC_TIME&")已经扫描入库!"
				else
				unitQty=csng(rs("packed_qty"))
				newUnitQty=unitQty
				end if
			end if
			rs.close						
			
			'use exist pallet
			
			'update packing plan status			
			
			
			if errMsg <> "" then			
				member("error")=errMsg
			else
				member("qty")=unitQty
				member("box_qty")=boxQty
				member("new_qty")=newUnitQty
				sql="update job_pack_detail set WHREC_JSQ=decode(WHREC_jsq,null,1,WHREC_jSQ+1) where box_id='"&boxId&"'" 
			conn.execute(sql)
			end if
			'response.Write toJSON(member)
					
		
		
	end select
end if
%>
<script type="text/javascript">
	window.returnValue='<%=toJSON(member)%>';
	window.close();
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->