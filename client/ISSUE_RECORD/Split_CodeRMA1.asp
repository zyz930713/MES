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
	
	   case "3"
	   
	   boxId=trim(request("txt_inputboxid"))
	         
			 sql="select * from Job_pack_DETAIL where box_Id='"&boxId&"'"
			 rs.open sql,conn,1,3
			 
			 if rs.eof then
				 errMsg="箱号不存在！."
			 else
			 
				 if left (rs("box_ID"),6)<>"REJECT" then
				  errMsg="此箱号不是客退箱号！."
				 end if
			 
			 end if
	       if errMsg <> "" then 
				
				member("error")=errMsg				
			
		   end if		
	
	
		case "2" 'Save data and print pallet label
		
		    opcode=trim(request("txt_opcode"))
			boxId=trim(request("txt_inputboxid"))	
			boxIdList=request("hidBoxIdList")
			newUnitQty=request("hidNewUnitQty")
			txt_boxqty=request("txt_boxqty")
			conn.beginTrans
			
			'generate pallet id
			
			
			if boxIdList<>"" then
			
				if ISSUE_ID="" then
								'生成box id						
								countType="LOG"
								countCondition=countType&formatdate(Now,"ymmdd")
								sql="select count_value,lm_time,count_type,count_condition from serial_count "
								sql=sql+"where count_type='"&countType&"' and count_condition = '"&countCondition&"'"
								rs.open sql,conn,1,3
								if rs.eof then
									ISSUE_ID=countCondition&"001"
									rs.addNew
									rs("count_type")=countType
									rs("count_condition")=countCondition
									rs("count_value")=1
									rs("lm_time")=now()
								else
									rs("count_value")=clng(rs("count_value"))+1
									ISSUE_ID=countCondition&repeatstring(rs("count_value"),"0",3)
									
									rs("lm_time")=now()
								end if
								rs.update
								rs.close													
				end if
			    
				txtcomments="从箱号删除二维码操作"&txt_boxqty&"个"
				
				conn.execute("insert into issue_record (LM_USER,box_id,ISSUE_ID，txtcomments,open_time) values ('"&opcode&"','"&boxId&"','"&ISSUE_ID&"','"&txtcomments&"',sysdate)")
				
				sql="INSERT INTO JOB_2D_CODE_HIST(CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER)"
				sql=sql+" SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER FROM JOB_2D_CODE "
				sql=sql+" WHERE code IN ('"&replace(boxIdList,",","','")&"')"
				rs.open sql,conn,1,1
				
				
					
				
				
				conn.execute("update JOB_2D_CODE set BOX_ID='"&boxId&"',PACK_USER='"&opcode&"' ,PACK_TIME=sysdate where   code in ('"&replace(boxIdList,",","','")&"')")
				
				
				sql="INSERT INTO JOB_2D_CODE_HIST(CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER)"
				sql=sql+" SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER FROM JOB_2D_CODE "
				sql=sql+" WHERE code IN ('"&replace(boxIdList,",","','")&"')"
				rs.open sql,conn,1,1
				
				
				conn.execute("update JOB_2D_CODE set BOX_ID='',IS_PACKED='',PACK_USER='' ,PACK_TIME='' where   code in ('"&replace(boxIdList,",","','")&"')")

				
			'	sql="select * from job_2d_code  where   code in ('"&replace(boxIdList,",","','")&"')"
			
			'	rs.open sql,conn,1,3
			'	while not rs.eof 
				
					
				'conn.execute("update Job_pack_DETAIL set PACKED_QTY=PACKED_QTY-1  where BOX_ID='"&boxId&"'")
				' set rsP=server.createobject("adodb.recordset")
				'  sqlp="select PACKED_QTY from  Job_pack_DETAIL where BOX_ID='"&boxId&"'"
				 '   rsP.open sqlp,conn,1,3
					'if not rs.eof then
					'PACKED_QTY= rsP("PACKED_QTY")
					'if cint(PACKED_QTY)=0 then 
					'conn.execute("delete Job_pack_DETAIL  where BOX_ID='"&boxId&"'")
					'end if
					
					'end if
				  'rsp.close
		   
				'rs.movenext
                'wend
			    'rs.close
			end if
			if err.number <> 0 then 				
				errMsg="解包失败|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="解包完成."
			end if			
			'response.Write toJSON(member)		
		case "1" 'input box id
		    code=trim(request("txt_inputcode"))
			boxId=request("txt_inputboxid")
			planPart=request("plan_part")
			palletId=request("txt_pallet") 
			boxId1= request("txt_boxid1")
			planId=request("slt_plan")
			newUnitQty=request("hidNewUnitQty")
			
			unitQty=0
			boxQty=1
			sql="select a.code,a.box_id,b.delivery_id from job_2d_code a,job_pack_detail b  where a.box_id=b.box_id and code='"&code& "'  group by  a.code,a.box_id,b.delivery_id "
			'sql="select part_number,pallet_id,sum(packed_qty) as packed_qty,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			'sql=sql+" group by part_number,pallet_id,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This code ("&code&") dose not exist.|该二维码("&code&")不存在或没有打包."
			else
				if isnull(rs("DELIVERY_id")) then
				errMsg="该二维码("&code&")没有出货记录，此二维码不是客退产品不能解箱!" 
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