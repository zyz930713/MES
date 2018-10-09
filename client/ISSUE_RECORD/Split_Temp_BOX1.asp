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
			if  rs.eof then
			   errMsg="该箱号("&boxId&")不存在临时库中，不能被解箱"
			
		   
			else
			
				unitQty=csng(rs("PACKED_QTY"))
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
			
			
			conn.beginTrans
			
			
			
			if left (boxIdList,1)="," then
			boxIdList=right(boxIdList,len(boxIdList)-1)
			end if
			
         
  			tArr = Split(boxIdList,",") 
'
			For i = 0 To UBound(tArr)  ' UBound(tArr) 遍历数组有多少个
			box_id=(tArr(i))  
		'	if ISSUE_ID="" then
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
		'	end if
			txtcomments="从临时箱删除箱号"
			conn.execute("insert into issue_record (LM_USER,box_id,ISSUE_ID，txtcomments,open_time) values ('"&opcode&"','"&boxId&"','"&ISSUE_ID&"','"&txtcomments&"',sysdate)")
			
			Next
			
			
			
			if errMsg <> "" then			
				member("error")=errMsg
			else
			
			
			
			
				if boxIdList <> "" then
				
				strWhere= " WHERE BOX_ID IN ('"&replace(boxIdList,",","','")&"')"
					sql="delete from job_pack"
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