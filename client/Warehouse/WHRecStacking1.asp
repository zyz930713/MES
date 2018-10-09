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
			WHlocation=request("WHlocation")
			txt_unitqty=request("txt_unitqty")
			conn.beginTrans
			
	       boxIdList=right(boxIdList,len(boxIdList)-1)
			tArr = Split(boxIdList,",")  '以逗号为分隔符，转换成数组tArr
			For i = 0 To UBound(tArr)  ' UBound(tArr) 遍历数组有多少个
					box_id=(tArr(i)) 					
				conn.Execute("INSERT INTO Box_id_Detail (box_id,WHREC_USER,BOXIDSTATUS,EDIT_TIME) values ('" & box_id & "','" & opcode & "','In',sysdate)")
			
			Next
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
						
				sql="update job_pack_detail set  pallet_id='"&palletId&"',stack_user='"&opcode&"',stack_time=sysdate , WHREC_user='"&opcode&"',WHREC_time=sysdate,"
				sql=sql+" GET_USER='',GET_TIME='',WHREC_JSQ=decode(WHREC_jsq,null,1,WHREC_jSQ+1),BOXIDSTATUS='"&WHlocation&"' where box_id in ('"&replace(boxIdList,",","','")&"')"
				conn.execute(sql)
				
				
				sql="delete job_pack   where box_id in ('"&replace(boxIdList,",","','")&"')"
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
					rtnCode=PrintCtl.PrintKEBPalletLabel(printName,palletId,txt_unitqty,CUSTOMER_NAME,QUANTITY)
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
			opcode=trim(request("txt_opcode"))
			WHlocation=request("WHlocation")
			WH=left(WHlocation,2)
            if WH<>"MR" then
				if WH<> left(boxId,2) then
				errMsg="箱号与库位不匹配！"
				end if
			elseif left(boxId,2)<>"FG" then				
				errMsg="报库箱号不可以入RMB"			
			end if 
			unitQty=0
			boxQty=1
			
			
			
			
			
			
			if errMsg="" then
			sql="select * from job_pack where box_id='"&boxId&"'"
			rs.open sql,conn,1,3
			
			if  rs.eof then
			 ' errMsg="该箱号("&boxId&")没有存在临时库中.请拉回扫描确认"
			end if
			rs.close
			end if
			
			if errMsg="" then
			
			sql="select part_number,sum(packed_qty) as packed_qty,WHREC_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,WHREC_TIME"
			rs.open sql,conn,1,3
				if rs.eof then
					errMsg="This box id("&boxId&") dose not exist.|该箱号("&boxId&")不存在.!!!!!!!!!!!!"
				else
					
					if  not isnull(rs("WHREC_TIME")) then
					WHREC_TIME=rs("WHREC_TIME")
					'WHREC_USER=rs("WHREC_USER")
					errMsg="该箱号("&boxId&")由工号("&WHREC_USER&")|在("&WHREC_TIME&")已经扫描入库!"
					else
					unitQty=csng(rs("packed_qty"))
					newUnitQty=csng(newUnitQty)+unitQty
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
			palletId=trim(request("txt_pallet"))		
			opcode=trim(request("txt_opcode"))
			'member("error")=palletId
			unitQty=0
			
			
			if errMsg="" then
				sql="select sum(packed_qty) as uniQty,PLAN_ID from job_pack_detail where PALLET_ID='"&palletId&"' group by pallet_id,PLAN_ID"
				rs.open sql,conn,1,3
				if rs.eof then
					errMsg="This pallet id("&palletId&") does not exist.|此拍号("&palletId&")不存在."
				elseif  not isnull(rs("PLAN_ID")) then
						errMsg="This pallet id("&palletId&") does not out.|此拍号("&palletId&")已绑定计划号，不能用此页面出库."
				else				
					unitQty=rs("uniQty")					
				end if
				rs.close
			end if
			
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
			getCode=trim(request("Get_opcode"))
			WHlocation=request("WHlocation")
		'	strWhere=" WHERE PALLET_ID='"&palletId&"'"
			if boxIdList <> "" then
				strWhere=" where PALLET_ID IN ('"&replace(boxIdList,",","','")&"')"
			end if
			   
			   
		  
		   conn.beginTrans
	        sql=" INSERT INTO Box_Id_Detail (Box_Id,Whrec_User,GET_USER,BOXIDSTATUS,EDIT_TIME) "
			sql=sql+"select Box_Id,'"&opCode&"' as whrec_user,'"&getCode&"' as get_user,'out' as boxidstatus,sysdate as edit_time from Job_Pack_Detail"
			sql=sql+strWhere
			conn.execute(sql)
			
			'move job_pack_detail to history
	sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME"
sql=sql+",DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID)"
sql=sql+" SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,"
sql=sql+"REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,"
sql=sql+"DELIVERY_ID,VENDOR_RECEIVE_DATE,CONFIRM_USER,CONFIRM_REMARKS,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ,INSPECTIONPQC,BIND_BOX_ID FROM JOB_PACK_DETAIL "
			sql=sql+strWhere
			conn.execute(sql)
			
			'spilt box
			'sql="UPDATE JOB_PACK_DETAIL SET IS_SHIPPED='',PALLET_ID='',SHIPPED_USER='',SHIPPED_TIME='',PLAN_ID='',STACK_USER='',STACK_TIME='' "
			'sql=sql+",DELIVERY_ID='' "+strWhere
			'conn.execute(sql)
			
			sql="UPDATE JOB_PACK_DETAIL SET PALLET_ID='',STACK_USER='',STACK_TIME='',WHREC_USER='"&opCode&"',WHREC_TIME='',GET_USER='"&getCode&"',GET_TIME=SYSDATE,get_JSQ=decode(get_jsq,null,1,get_jSQ+1),BOXIDSTATUS='"&WHlocation&"' "
			sql=sql+strWhere
			conn.execute(sql)
			
		  
			if err.number <> 0 then 				
				errMsg="Save failed.保存失败.|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			    conn.commitTrans     '执行事务提交
			   	member("message")="Save succesful.|解拍出库成功."
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