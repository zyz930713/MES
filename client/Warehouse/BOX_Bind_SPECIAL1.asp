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
			opcode=trim(request("txt_opcode"))						
			txt_box=trim(request("txt_box"))
			txt_unitqty=trim(request("txt_unitqty"))
			Bind_boxid=trim(request("Bind_boxid"))
			Bind_boxid_Qty=trim(request("Bind_boxid_Qty"))
			PART_NUMBER=trim(request("PART_NUMBER"))
			Bind_Part=trim(request("Bind_Part"))
			conn.beginTrans
			
			sql="select * from BOX_ID_SPECIAL where BIND_BOX_ID='"&Bind_boxid&"'"
			rs.open sql,conn,1,3
			if not rs.eof then
			   qty=rs("PACKED_QTY")
			   errMsg="该箱号("&Bind_boxid&")已绑定，数量("&qty&")，请确认后，先删除再重新绑定"
			end if
			rs.close
			
			
			sql="select * from CUSTOMER_MATERIAL where CUSTOMER_LABEL='HUAWEI' and  ITEM_NAME='"&part_number&"'"
			rs.open sql,conn,1,3
			if  not rs.eof then			 
			   CUSTOMER_PN=rs("CUSTOMER_PN")
			end if
			rs.close
			
			if CUSTOMER_PN<>Bind_Part then
			errMsg="料号不匹配！"
			end if 
			
			if txt_unitqty<> Bind_boxid_Qty then
			errMsg="数量不对"
			end if
			'generate pallet id
			'errMsg=part_number
			
			if errMsg= "" then 
			
		sql="insert into BOX_ID_SPECIAL (BOX_ID,BIND_BOX_ID,PART_NUMBER,CUSTOMER,PACKED_QTY,WHREC_USER,WHREC_TIME)"
	    sql=sql+"  VALUES  ('"+txt_box+"','"+Bind_boxid+"','"+part_number+"','"+Bind_Part+"',"&txt_unitqty&",'"+opcode+"',sysdate)"
		'sql="insert into JOB_PACK_DETAIL_SPECIAL (BOX_ID,BIND_BOX_ID)"
	  ' sql=sql+"  VALUES  ('"+txt_box+"','"+Bind_boxid+"')"
		    'sql="insert into BOX_ID_SPECIAL (BOX_ID,BIND_BOX_ID) VALUES  ('"+txt_box+"','"+Bind_boxid+"')"
	
			 conn.execute(sql)   
				'conn.execute(sql)
				
					
			end if
			if err.number <> 0 then 				
				errMsg="绑定失败|"&err.description
			end if 
			
			if errMsg <> "" then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")=errMsg				
			else 
			   conn.commitTrans     '执行事务提交
			    member("message")="确认绑定完成."
			end if			
			'response.Write toJSON(member)		
		case "1" 'input box id
			boxId=request("txt_inputboxid")
			planPart=request("plan_part")
			palletId=request("txt_pallet") 
			boxId1= request("txt_boxid1")
			planId=request("slt_plan")
			opcode=request("txt_opcode")
			newUnitQty=request("hidNewUnitQty")
			
           
			unitQty=0
			boxQty=1
			sql="select part_number,sum(packed_qty) as packed_qty,PACKED_LINE,PROD_SHIFT,WHREC_TIME from job_pack_detail where box_id='"&boxId&"'" 			 	
			sql=sql+" group by part_number,PACKED_LINE,PROD_SHIFT,WHREC_TIME"
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This box id("&boxId&") dose not exist.|该箱号("&boxId&")不存在."
			else
				
				
				unitQty=csng(rs("packed_qty"))
				part_number=rs("part_number")
				newUnitQty=unitQty
				PACKED_LINE=rs("PACKED_LINE")
				PROD_SHIFT=rs("PROD_SHIFT")
				
			end if
			rs.close						
			
				
				if part_number="990321100028"  then
				
				part_number="240326000001"
				
				
				end if
			sql="select * from BOX_ID_SPECIAL where box_id='"&boxId&"'"
			rs.open sql,conn,1,3
			if not rs.eof then
			   qty=rs("PACKED_QTY")
			   errMsg="该箱号("&boxId&")已绑定，数量("&qty&")，请确认后，先删除再重新绑定"
			end if
			rs.close
			
			sql="select * from CUSTOMER_MATERIAL where CUSTOMER_LABEL='HUAWEI' and  ITEM_NAME='"&part_number&"'"
			rs.open sql,conn,1,3
			if  rs.eof then			 
			   errMsg="此产品不是华为产品！"
			end if
			rs.close
		
			'use exist pallet
		
			if errMsg= "" then 
			

			
				member("qty")=unitQty
				member("box_qty")=boxQty
				member("new_qty")=newUnitQty
				member("boxID")=boxId	
				member("part_number")=part_number			
			  else			
			   ' 
				member("error")=errMsg		
			  end if		
			'end if
			
		
			'response.Write toJSON(member)
					
		
		
	end select
end if
%>
<script type="text/javascript">
	window.returnValue='<%=toJSON(member)%>';
	window.close();
</script>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->