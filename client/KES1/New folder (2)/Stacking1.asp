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
			planId=request("slt_plan")
			palletId=request("txt_pallet")				
			boxIdList=request("hidBoxIdList")
			newUnitQty=request("hidNewUnitQty")
			conn.beginTrans
			
			'generate pallet id
			if palletId="" then
				sql="select to_char(sysdate,'Y')as year,to_char(sysdate,'DD') as day,to_char(sysdate,'MM') as mon from dual"
				rs.open sql,conn,1,3
				if not rs.eof then
					strYear=rs("year")
					strMon=rs("mon")
					strDay=rs("day")
				end if
				rs.close
				countCondition="KEBPL"&strYear&strMon&strDay				
				sql="select count_value,lm_time,count_type,count_condition from serial_count "
				sql=sql+"where count_type='KEBPL' and count_condition = '"&countCondition&"'"
				rs.open sql,conn,1,3
				if rs.eof then
					palletId=countCondition&"001"
					rs.addNew
					rs("count_type")="KEBPL"
					rs("count_condition")=countCondition
					rs("count_value")=1
					rs("lm_time")=now()
				else
				    rs("count_value")=clng(rs("count_value"))+1
					palletId=countCondition&repeatstring(rs("count_value"),"0",3)
				
					rs("lm_time")=now()
				end if
				rs.update
				rs.close
			elseif boxIdList="" then'reprint pallet label
				sql="select user_code from users where user_code='"&opcode&"' and instr(roles_id,'REPRINT_PALLET_LABEL')>0 "
				rs.open sql,conn,1,3
				if rs.eof then
					errMsg="You are not authorized to reprint pallet.|您没有权限进行重印拍号。"
				end if
				rs.close				
			end if
			if boxIdList<>"" then
				sql="update job_pack_detail set pallet_id='"&palletId&"', plan_id='"&planId&"',stack_user='"&opcode&"',stack_time=sysdate where box_id in ('"&replace(boxIdList,",","','")&"')"
				conn.execute(sql)
				
				sql="update packing_plan set status='Wait',pack_qty=pack_qty+"&newUnitQty&" where plan_id='"&planId&"'"	
				conn.execute(sql)	
			end if
			if err.number <> 0 then 				
				errMsg="Print failed.打印失败|"&err.description
			end if 
			if errMsg = "" then	
				'print pallet label
				sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='WHPrint'"
				rs.open sql,conn,1,3
				if not rs.eof then
					printName = rs("PRINTER_NAME")
				end if
				rs.close				
				if printName="" then
					member("error")="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
				else
					set PrintCtl=server.createobject("PrintClass.PrintCtl")                       
					rtnCode=PrintCtl.PrintKEBPalletLabel(printName,palletId)
					'rtnCode="OK"
					if rtnCode<>"OK" then
						'member("error")="Label print error.打印标签发生错误.|" & rtnCode		
						errMsg="Label print error.打印标签发生错误.|" & rtnCode
					end if
				end if
			end if
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="Print succesful.|打印成功."
			end if			
			'response.Write toJSON(member)		
		case "1" 'input box id
			boxId=request("txt_inputboxid")
			planPart=request("plan_part")
			palletId=request("txt_pallet")
			boxId1= request("txt_boxid1")
			planId=request("slt_plan")
			newUnitQty=request("hidNewUnitQty")
			
			unitQty=0
			boxQty=1
			sql="select part_number,pallet_id,sum(packed_qty) as packed_qty,sum( case when WHREC_Time is null then 0 else 1 end ) as WT  from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,pallet_id"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This box id("&boxId&") dose not exist.|该箱号("&boxId&")不存在."
			
			elseif rs("WT")="0" then
			errMsg="该箱号("&boxId&")没有入库扫描，请先进行入库扫描."
			else
				if not isnull(rs("pallet_id")) then
					errMsg="This box id("&boxId&") was stacked.|该箱号("&boxId&")已经堆栈."
				end if
				if instr(planPart,rs("part_number"))=0 then
					errMsg="Part number of This box id("&boxId&") is different with packing plan.|该箱号("&boxId&")的料号与包装计划不一致."
				end if
				unitQty=csng(rs("packed_qty"))
				newUnitQty=csng(newUnitQty)+unitQty
			end if
			rs.close						
			
			'use exist pallet
			if boxId1="" and palletId<>"" then
				sql="select count(distinct box_id) as boxqty,sum(packed_qty) as packed_qty,plan_id from job_pack_detail where pallet_id='"&palletId&"' group by plan_id"
				rs.open sql,conn,1,3
				if not rs.eof then
					if planId <> rs("plan_id") then
						errMsg="Plan Id of selected and pallet is different.|选择的计划与拍号的计划不一致."
					else
						unitQty=unitQty+csng(rs("packed_qty"))
						boxQty=boxQty+csng(rs("boxqty"))
					end if
				end if
				rs.close
			end if
			'update packing plan status			
			if errMsg="" then
				sql="select quantity,status,pack_qty from packing_plan where plan_id='"&planId&"'"
				rs.open sql,conn,1,3
				if not rs.eof then
					if csng(rs("quantity"))<newUnitQty+csng(rs("pack_qty")) then
						errMsg="Packing qty exceeds plan qty.|已包装的数量超过计划设定的数量."
					end if
					if errMsg="" and boxId1=""  then
						if rs("status")<>"Initial" and rs("status")<>"Wait" then
							errMsg="Status("&rs("status")&") of packing plan is not Initial or Wait, cannot do Stack.|此计划的状态("&rs("status")&")不是Initial或者Wait,不能做堆栈。"
						'else
						'	rs("status")="Run"
						'	rs.update
						end if
					end if										
				end if
				rs.close
			end if
			
			if errMsg <> "" then			
				member("error")=errMsg
			else
				member("qty")=unitQty
				member("box_qty")=boxQty
				member("new_qty")=newUnitQty
			end if
			'response.Write toJSON(member)
					
		case "3" 'Close
			planId=request("slt_plan")
			sql="update packing_plan set status='Wait' where plan_id='"&planId&"'"
			conn.execute(sql)			
			'response.Write toJSON(member)
		case "4" 'input box id to split pallet
			palletId=request("txt_pallet")
			boxId=request("txt_inputboxid")
			'member("error")=palletId
			unitQty=0
			sql="select pallet_id,sum(packed_qty) as uniQty from job_pack_detail where box_id='"&boxId&"' group by pallet_id"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This box id("&boxId&") does not exist.|此箱号("&boxId&")不存在."
			elseif isnull(rs("pallet_id")) or rs("pallet_id") <> palletId then
				errMsg="Pallet id("&rs("pallet_id")&") of this box is different with input pallet id("&palletId&").|此箱的拍号("&rs("pallet_id")&")与输入的("&palletId&")不一致。"
			else
				unitQty=csng(rs("uniQty"))
			end if
			rs.close
			if errMsg <> "" then			
				member("error")=errMsg
			else
				member("qty")=unitQty
				member("box_qty")=1				
			end if
			'response.Write toJSON(member) 
		case "5" 'split pallet
			palletId=request("txt_pallet")
			boxIdList=request("hidBoxIdList")
			opCode=request("txt_opcode")
			txt_unitqty=trim(request("txt_unitqty"))
			
			strWhere=" WHERE PALLET_ID='"&palletId&"'"
			if boxIdList <> "" then
				strWhere=strWhere+" AND BOX_ID IN ('"&replace(boxIdList,",","','")&"')"
			end if
			   
			    sql="select plan_id from job_pack_detail where PALLET_ID='"&palletId&"' GROUP by plan_id "
				rs.open sql,conn,1,1
				if not rs.eof then
				plan_id=rs("plan_id")		
				end if			
				rs.close
		  
		   conn.beginTrans
			
			'move job_pack_detail to history
	sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME"
sql=sql+",DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID) SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,"
sql=sql+"REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,"
sql=sql+"DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID FROM JOB_PACK_DETAIL "
			sql=sql+strWhere
			conn.execute(sql)
			
			'spilt box
			sql="UPDATE JOB_PACK_DETAIL SET IS_SHIPPED='',PALLET_ID='',SHIPPED_USER='',SHIPPED_TIME='',PLAN_ID='',STACK_USER='',STACK_TIME='' "
			sql=sql+",DELIVERY_ID='' "+strWhere
			conn.execute(sql)
			
'			
			sql="select * from PACKING_PLAN where plan_id='"&plan_id&"'"
			 
			rs.open sql,conn,1,3
			if not  rs.eof then
			PACK_QTYA=CLng(trim(rs("PACK_QTY")))-CLng(txt_unitqty)
			end if
			rs("PACK_QTY")=PACK_QTYA
			rs.update
		   
			rs.close
		  
			if err.number <> 0 then 				
				errMsg="Save failed.保存失败.|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			   	member("message")="Save succesful.|保存成功."
			end if			
			response.Write toJSON(member)
	end select
end if
%>
<script type="text/javascript">
	window.returnValue='<%=toJSON(member)%>';
	window.close();
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->