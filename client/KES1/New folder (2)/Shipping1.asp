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

			
if len(asynid)>0 then
	dim member
	set member=jsObject()
	select case asynid
		case "3" 'Save			
			
			
			opcode=request("txt_opcode")
			delivery=request("txt_delivery")
			palletList=	request("hidPalletIdList")
			conn.beginTrans	
	
			sql="update job_pack_detail set is_shipped=1,shipped_user='"&opcode&"',shipped_time=sysdate,delivery_id='"&delivery&"' where pallet_id in ('"&replace(palletList,",","','")&"')"
			'sql="select * from job_pack_detail where pallet_id in  ('','KEBPL30718001')"
			conn.execute(sql)
			'errMsg=sql
			'member("error")=errMsg
			if err.number <> 0 then 
				conn.rollbackTrans   '对已执行的操作回滚
			   member("error")="Save failed.保存失败.|"&err.description
			' errMsg="Save failed.保存失败.|"&err.description
'				
			else 
			    conn.commitTrans     '执行事务提交
			   member("message")="Save succesful.|保存成功."
'			 
'				'errMsg="Save succesful.|保存成功."
			end if	
		'	response.Write toJSON(member)		
		case "2" 'input pallet
		
			pallet=request("txt_pallet")
			delivery=request.Form("txt_delivery")
			unitQty=0
			palletQty=1
			
			'check pallet no
			sql="select delivery_id,sum(packed_qty) qty from job_pack_detail where pallet_id='"&pallet&"' group by delivery_id "
			rs.open sql,conn,1,3
			if rs.eof then
				errMsg="This pallet id("&pallet&") does not exists.|此拍号("&pallet&")不存在。"
			elseif not isnull(rs("delivery_id")) then
				errMsg="This pallet id("&pallet&") was shipped.|此拍号("&pallet&")已出货。"
			else
				unitQty=csng(rs("qty"))
			end if
			rs.close
			
			'get delivery info
			sql="select count(distinct pallet_id) as palletQty,sum(packed_qty) as uniQty from job_pack_detail where delivery_id='"&delivery&"' "
			rs.open sql,conn,1,3
			if not rs.eof then
				unitQty=unitQty+csng(rs("uniQty"))
				palletQty=palletQty+csng(rs("palletQty"))
			end if
			
			if errMsg="" then
				member("unit_qty")=unitQty
				member("pallet_qty")=palletQty
			else
				member("error")=errMsg
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