<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%selectType=request("selectType")
txtSubJobList=request("txtSubJobList")
Action=request("Action")
box_id=request("box_id")
opCode=request("txtOperatorCode")
if Action="DEL" and box_id <>"" then
     set rsA=server.createobject("adodb.recordset")
	 set rsB=server.createobject("adodb.recordset")
     sqlA="select * from job_pack_detail where box_id ='"&box_id&"'"
     rsA.open sqlA,conn,1,3
if rsA.eof then
	 word="<span align='center' style='color:red;'>The "&box_id&" is not Print, please Print.<br>"&"</span>"
	 response.Redirect("Packup1.asp?word="&word)
	 else
	 do while not rsA.eof 
	      PACKED_QTY=rsA("PACKED_QTY")
		  JOB_NUMBER=rsA("JOB_NUMBER")
		  sql="select * from job_inventory where job_number='"&JOB_NUMBER&"'"
		  rsB.open sql,conn,1,3
		  if rsB.eof then
		  else
		  rsB("good_qty")=csng(rsB("good_qty"))+csng(PACKED_QTY)
		  PACKED_QTYA=csng(rsB("PACKED_QTY"))
		  rsB("PACKED_QTY")=PACKED_QTYA-csng(PACKED_QTY)
		  rsb.update
		  end if
		  rsB.close
     rsA.movenext
	 loop
end if
     rsA.close
sql="INSERT INTO JOB_2D_CODE_HIST(CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER)"
				sql=sql+" SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER FROM JOB_2D_CODE "
				sql=sql+" WHERE BOX_ID IN ('"&box_id&"')"
				rs.open sql,conn,1,1
conn.Execute("update job_2d_code set box_id='', is_packed='', pack_time='', pack_user=''    where box_id='"&box_id&"'") 

				
sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
				sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ)"
				sql=sql+" SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS||' Split by "&opCode&" at '||sysdate as REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,"
				sql=sql+"SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME,GET_JSQ,BOXIDSTATUS,WHREC_JSQ FROM JOB_PACK_DETAIL "
				sql=sql+" WHERE BOX_ID IN ('"&box_id&"')"
				rs.open sql,conn,1,1			
					 
    
conn.Execute("DELETE job_pack_detail where box_id in ('"&box_id&"')")
response.Redirect("Split_Box_index.asp")
else
response.Redirect ("Split_Box_index.asp")
end if
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->