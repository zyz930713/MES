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
		
		case "1" 'input box id to RMA
		
		boxId=trim(request("txt_boxId"))
			if boxId="" then
			
			conn.beginTrans
							'����box id						
							countType="BoxId"
							packType="REJECT"
							countCondition=packType&packLine&formatdate(Now,"ymmddhh")
							sql="select count_value,lm_time,count_type,count_condition from serial_count "
							sql=sql+"where count_type='"&countType&"' and count_condition = '"&countCondition&"'"
							rs.open sql,conn,1,3
							if rs.eof then
								boxId=countCondition&"001"
								rs.addNew
								rs("count_type")=countType
								rs("count_condition")=countCondition
								rs("count_value")=1
								rs("lm_time")=now()
							else
							    rs("count_value")=clng(rs("count_value"))+1
								boxId=countCondition&repeatstring(rs("count_value"),"0",3)
								
								rs("lm_time")=now()
							end if
							rs.update
							rs.close	
						member("boxId")=boxId	
				else			
					errMsg="����Ѿ����ɣ������ظ�����"														
				end if
				
				if err.number <> 0 then 				
				errMsg="Save failed.����ʧ��.|"&err.description
				end if 
				
				if errMsg <> "" then 
				conn.rollbackTrans   '����ִ�еĲ����ع�
				member("error")=errMsg				
				else 
				conn.commitTrans     'ִ�������ύ
				member("message")="Save succesful.|����ɹ�."
				end if			
				response.Write toJSON(member)
						
		case "2" 'RMAPrint
		txt_opcode=trim(request("txt_opcode"))
		boxid=trim(request("txt_boxId"))
		partNo=trim(request("txt_Part"))
		qty=trim(request("txt_QTY"))
		jobno=trim(request("txt_RMANO"))
		remarks=trim(request("txt_Remarks"))
		customer=""
		
		
		
			if errMsg = "" then	
				'print pallet label
				sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='WHPrint'"
				rs.open sql,conn,1,3
				if not rs.eof then
					printName = rs("PRINTER_NAME")
				end if
				rs.close				
				if printName="" then
					member("error")="Please contact engiener to set the printer for this machine.|����ϵ����ʦΪ�˻����趨��ȷ�ı�ǩ��ӡ����"
				else
					set PrintCtl=server.createobject("PrintClass.PrintCtl")                       
					rtnCode=PrintCtl.PrintBoxLabel(printName,boxid,partNo,jobno,qty,remarks,customer)
					'rtnCode="OK"
					if rtnCode<>"OK" then
						'member("error")="Label print error.��ӡ��ǩ��������.|" & rtnCode		
						errMsg="Label print error.��ӡ��ǩ��������.|" & rtnCode
					end if
				end if
			end if
			
			if errMsg="" then
			conn.beginTrans
			sql="insert into job_pack_detail (BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,PACKED_USER,PACKED_TIME,WHREC_USER,WHREC_TIME,BOXIDSTATUS) VALUES ('"&boxid&"','"&jobno&"','"&partNo&"','"&qty&"','"&txt_opcode&"',sysdate,'"&txt_opcode&"',sysdate,'MRB')"
		    conn.execute(sql)
			
			sql="insert into RMA_LIST (RMA_ID,BOX_ID,PACKED_QTY,RMA_USER,RMA_TIME ) VALUES ('"&jobno&"','"&boxid&"','"&qty&"','"&txt_opcode&"',sysdate)"
			conn.execute(sql)
			end if
			
			
			if errMsg <> "" then 
				conn.rollbackTrans   '����ִ�еĲ����ع�
				member("error")=errMsg				
			else 
			    conn.commitTrans     'ִ�������ύ
			    member("message")="Print succesful.|��ӡ�ɹ�."
			end if	
			
		case "3" 'Part  DESCRIPTION
		
		    ITEM_NAME=trim(request("txt_Part"))
			'errMsg=ITEM_NAME
			
			sql="select DESCRIPTION from PRODUCT_MODEL where ITEM_NAME='"&ITEM_NAME&"'"
			rs.open sql ,conn,1,3	
				if not rs.eof then
			    member("PARTDESCRIPTION")=rs("DESCRIPTION")	
				
				else
				errMsg="12NC�Ų����ڣ����������룡"		
				end if
			    rs.close	
			
		    if errMsg <> "" then 
			
				member("error")=errMsg				
			
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