<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
On Error Resume Next
'ajax��ȡ����
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
			boxIdList=request("hidBoxIdList")
			newUnitQty=request("hidNewUnitQty")
			WHlocation=request("WHlocation")
			conn.beginTrans
			
			'generate pallet id
			boxIdList=right(boxIdList,len(boxIdList)-1)
			tArr = Split(boxIdList,",")  '�Զ���Ϊ�ָ�����ת��������tArr
			For i = 0 To UBound(tArr)  ' UBound(tArr) ���������ж��ٸ�
					box_id=(tArr(i)) 					
				conn.Execute("INSERT INTO Box_id_Detail (box_id,WHREC_USER,GET_USER,BOXIDSTATUS,EDIT_TIME) values ('" & box_id & "','" & opcode & "','" & GET_user & "','Out',sysdate)")
			
			Next
		
			if boxIdList<>"" then
	sql="update job_pack_detail set WHREC_user='"&opcode&"',WHREC_time='',GET_user='"&GET_user&"',GET_time=sysdate,get_JSQ=decode(get_jsq,null,1,get_jSQ+1),BOXIDSTATUS='"&WHlocation&"'  where box_id in ('"&replace(boxIdList,",","','")&"')"
				conn.execute(sql)
				
					
			end if
			if err.number <> 0 then 				
				errMsg="ȷ�ϳ���ʧ��|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '����ִ�еĲ����ع�
				member("error")=errMsg				
			else 
			    conn.commitTrans     'ִ�������ύ
			    member("message")="ȷ�ϳ������."
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
			unitQty=0
			boxQty=1
		
			sql="select part_number,count(pallet_id)  as pallet_id,sum(packed_qty) as packed_qty,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,pallet_id,WHREC_USER,WHREC_TIME,GET_USER,GET_TIME"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This box id("&boxId&") dose not exist.|�����("&boxId&")������."
			else
				if isnull(rs("WHREC_TIME")) then
				errMsg="�����("&boxId&")û����⣬���Բ��ܳ���!" 
				elseif clng(rs("pallet_id"))>0 then
				errMsg="�����("&boxId&")�Ѷ��ģ����Ƚ��ĺ��ٳ���!" 
				else
				
					if  not isnull(rs("GET_USER")) then				
					GET_USER=rs("GET_USER")
					GET_TIME=rs("GET_TIME")
					errMsg="�����("&boxId&")�ɹ���("&GET_USER&")|��("&GET_TIME&")�Ѿ�ɨ�����!"
					else
					unitQty=csng(rs("packed_qty"))
					newUnitQty=csng(newUnitQty)+unitQty
				    end if
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