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
			GET_user=request("GET_opcode")		
			part_number=trim(request("txt_part_number"))
			WHlocation=request("WHlocation")
			conn.beginTrans
			
			'generate pallet id
	        if part_number<>"" then
			strWhere=" where part_number='"&part_number&"' and Is_Shipped is null and pallet_id is null and Boxidstatus='"&WHlocation&"' "
			else
			   errMsg="料号为空不能操作！"
			end if
		  
			
	        sql=" INSERT INTO Box_Id_Detail (Box_Id,Whrec_User,GET_USER,BOXIDSTATUS,EDIT_TIME) "
			sql=sql+"select Box_Id,'"&opCode&"' as whrec_user,'"&GET_user&"' as get_user,'out' as boxidstatus,sysdate as edit_time from Job_Pack_Detail"
			sql=sql+strWhere 
			conn.execute(sql)		
	
	
sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME"
sql=sql+",DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID) SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,"
sql=sql+"REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,"
sql=sql+"DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID FROM JOB_PACK_DETAIL "
sql=sql+strWhere
conn.execute(sql)	
	
	
			sql="update job_pack_detail set WHREC_user='"&opcode&"',WHREC_time='',GET_user='"&GET_user&"',GET_time=sysdate,get_JSQ=decode(get_jsq,null,1,get_jSQ+1),BOXIDSTATUS='' "	 
			sql=sql+strWhere 
	
			conn.execute(sql)
			
		
			if err.number <> 0 then 				
				errMsg="确认出库失败|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="确认出库完成."
			end if	
			
			'response.Write toJSON(member)		
		case "1" 'input box id
			part_number=request("txt_part_number")
		
			WHlocation=request("WHlocation") 
			
			newUnitQty=request("hidNewUnitQty")
			
			unitQty=0
			boxQty=1
			
			
			
			
			sql="select part_number,sum(packed_qty) as packed_qty from job_pack_detail where part_number='"&part_number&"' and Is_Shipped is null and pallet_id is null and Boxidstatus='"&WHlocation&"'" 			 	
			sql=sql+" group by part_number"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This Part Number"&part_number&") dose not exist.|该料号("&part_number&")不存在."
						
			else
			unitQty=csng(rs("packed_qty"))
			newUnitQty=csng(newUnitQty)+unitQty
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