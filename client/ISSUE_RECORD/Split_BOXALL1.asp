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
		
		
					
		
		case "4" 'input box id to RMA
		
			boxId=trim(request("txt_inputboxid"))
			'member("error")=palletId
			unitQty=0
			
			
			sql="select * from job_pack where box_id='"&boxId&"'"
			rs.open sql,conn,1,3
			if not rs.eof then
			   errMsg="该箱号("&boxId&")存在临时库中，不能被解箱"
			end if
			rs.close
			
			
		    sql="select part_number,count(pallet_id)  as pallet_id,sum(packed_qty) as uniQty,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,pallet_id,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME"	
			
			rs.open sql,conn,1,3
			if rs.eof or isnull(rs("uniQty")) then
				errMsg="This box id("&boxId&") does not exist.|此箱号("&boxId&")不存在."
			elseif clng(rs("pallet_id"))>0 then
				errMsg="该箱号("&boxId&")已堆拍，请先解拍号再解箱!" 
			elseif rs("WHREC_TIME")<>"" then
				errMsg="该箱号("&boxId&")没有出库，所以不能解箱!" 			
			else
			
				unitQty=csng(rs("uniQty"))
				plan_id=rs("plan_id")
				
				
			end if
			rs.close
			if errMsg <> "" then			
				member("error")=errMsg
			else
				member("qty")=unitQty
				member("box_qty")=1				
			end if
			'response.Write toJSON(member) 
		case "5" 'RMA
			
			boxIdList=request("hidBoxIdList")
			boxIDQTY=request("hidBoxIdQTY")
			txt_boxqty=request("txt_boxqty")
			opCode=trim(request("txt_opcode"))
			txt_unitqty=trim(request("txt_unitqty"))
			'errMsg=boxIdList
			
			'errMsg=opCode
		
			
			
			conn.beginTrans
			
			set rsB=server.createobject("adodb.recordset")
			
			if left (boxIdList,1)="," then
			boxIdList=right(boxIdList,len(boxIdList)-1)
			end if
			
          'errMsg=boxIdList
  			tArr = Split(boxIdList,",") 
'
			For i = 0 To UBound(tArr)  ' UBound(tArr) 遍历数组有多少个
			box_id=(tArr(i))  
			
			
			sql="select * from job_pack_detail where box_id ='"&box_id&"'"
          
			rs.open sql,conn,1,3
			
			if  rs.eof then
			errMsg="箱号不存在"	
		    else
			    do while not rs.eof 
				   PACKED_QTY=rs("PACKED_QTY")
				   JOB_NUMBER=rs("JOB_NUMBER")
'				   errmsg=JOB_NUMBER
				   sql="select * from job_inventory where job_number='"&JOB_NUMBER&"'"
				   rsB.open sql,conn,1,3
					   if rsB.eof then
					 '  errMsg="工单不存在"	
					   else
					   rsB("good_qty")=csng(rsB("good_qty"))+csng(PACKED_QTY)
					   PACKED_QTYA=csng(rsB("PACKED_QTY"))
					   rsB("PACKED_QTY")=PACKED_QTYA-csng(PACKED_QTY)
					   rsB.update					   
					   end if
				   rsB.close
			    rs.movenext
			    loop
			'errMsg=cstr(PACKED_QTYA)
			end if
			rs.close
			
			Next
			
			
			
			if errMsg <> "" then			
				member("error")=errMsg
			else
			
			
			
			
		    if boxIdList <> "" then
			
			strWhere= " WHERE BOX_ID IN ('"&replace(boxIdList,",","','")&"')"
			'conn.beginTrans
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
				txtcomments="共删除"&txt_boxqty&"箱，共"&txt_unitqty&"个"
				
				conn.execute("insert into issue_record (LM_USER,box_id,ISSUE_ID，txtcomments,open_time) values ('"&opcode&"','"&boxId&"','"&ISSUE_ID&"','"&txtcomments&"',sysdate)")
			
			
			
			    'move job_2d_code to history
				sql="INSERT INTO JOB_2D_CODE_HIST(CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER)"
				sql=sql+" SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER FROM JOB_2D_CODE "
				sql=sql+strWhere
				rs.open sql,conn,1,1
				
							
				
				sql="update job_2d_code set box_id='',is_packed='',pack_time='',pack_user=''"
				sql=sql+strWhere
				rs.open sql,conn,1,1				
				
				'move job_pack_detail to history
				'sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
				''sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ)"
				'sql=sql+" SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,"
				'sql=sql+"SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ FROM JOB_PACK_DETAIL "
				'sql=sql+strWhere
sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME"
sql=sql+",DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID) SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,"
sql=sql+"REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,"
sql=sql+"DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID FROM JOB_PACK_DETAIL "
sql=sql+strWhere
				rs.open sql,conn,1,1
				
				'delete job_pack_detail
				sql="delete from job_pack_detail"
				sql=sql+strWhere
				rs.open sql,conn,1,1								
			
             end if
'			
		
			if err.number <> 0 then 				
				errMsg="Save failed.保存失败.|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			   	member("message")="Save succesful.|解箱成功."
			end if			
			response.Write toJSON(member)
			
	 end if
	end select
end if
%>
<script type="text/javascript">
	window.returnValue='<%=toJSON(member)%>';
	window.close();
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->