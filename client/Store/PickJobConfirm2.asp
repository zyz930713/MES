<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/TrackJobMaster.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/GetFactoryID.asp" -->
<%

code=trim(request.Form("txtUserCode"))
rework="0"
factory=request.Form("factory")
jobnumber=request("txtSelectJobNumber")
sheetnumber=request("txtSelectSheetNumber")
pagename="PickJobConfirm.asp?action=2&factory="&factory&"&txtSelectJobNumber="+jobnumber+"&txtSelectSheetNumber="+sheetnumber

NO_YIELD=request.Form("NO_YIELD")
subjobnumber=ucase(trim(request.Form("jobnumber")))
subjobnumber=left(subjobnumber,len(subjobnumber)-1)
subjobnumberdisplay=subjobnumber
session("JOB_NUMBER")=subjobnumber
asubjobnumber=split(subjobnumber,",")
store_quantity=cint(request.Form("quantity"))
inspect_quantity=cint(request.Form("inspect"))
rejectqty=cint(request.Form("rejectqty"))
INSTORE_GUID_ID=request.Form("txtINSTORE_GUID_ID")

erp_account=request.Form("ERP_Account")
erp_reason=request.Form("ERP_Reason")
'erp_refer_str=request.Form("txtRefer")
'arrerp_refer_str=split(erp_refer_str,",")


	
	
'check whether user is recorded operator.
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
	rs.open SQL,conn,1,3
	if rs.eof then
		response.Redirect(pagename+"&errorstring="&code&" 不存在，请联系工程师。")
	else
		if rs("LOCKED")="1" then
			response.Redirect(pagename+"&errorstring="&code&" 被锁定，请联系工程师。")
		else
			if rs("PRACTISED")="1" and (date()<rs("PRACTISE_START_TIME") or date()>rs("PRACTISE_END_TIME")) then
				response.Redirect(pagename+"&errorstring="&code&" 的实习期限已过期，请联系工程师。")
			end if
		end if
	end if
rs.close

'check whether sub job number is stored, it this sub job is stored, return to store1.asp
'if not, get INPUT_QUANITY of total sub job
'if rework<>"1" then
if rework="0" then
	input_quantity=0
		for i=0 to ubound(asubjobnumber)
			'check have stored or not
			SQL="select * from JOB_MASTER_STORE_PRE where SUB_JOB_NUMBERS like '%"&trim(asubjobnumber(i))&"%' and store_type='N' and STORE_TIME>=to_date('2008-1-1','yyyy-mm-dd hh24:mi:ss')"
			rs.open SQL,conn,1,3
			if not rs.eof then
				response.Redirect(pagename+"&errorstring="&asubjobnumber(i)&"已经由"&rs("STORE_CODE")&"在"&rs("STORE_TIME")&"做过"&rs("JOB_NUMBER")&"入库！请重新选择工单。")
				exit for
			end if
			rs.close
			'get job number, sheet number and job type of each sub job
			if instr(trim(asubjobnumber(i)),"-")=0 then
				response.Redirect(pagename+"&errorstring="&asubjobnumber(i)&"工单格式错误，缺少-，请重新扫描子工单！")
			end if
				thisasubjobnumber=split(trim(asubjobnumber(i)),"-")
				thissubjobnumber=""
			for w=0 to ubound(thisasubjobnumber)
				if w<>ubound(thisasubjobnumber) then
					thissubjobnumber=thissubjobnumber&thisasubjobnumber(w)&"-"
				else
					if isnumeric(thisasubjobnumber(w)) then
						thissheetnumber=cint(thisasubjobnumber(w))
						jobtype="N"	
					else
						response.Redirect(pagename+"&errorstring="&asubjobnumber(i)&"工单格式错误，子工单编号不是数字，请重新扫描子工单！")
					end if
				end if
			next
			thissubjobnumber=left(thissubjobnumber,len(thissubjobnumber)-1)
			
			'check is sub sheet finsihed prodution in production lines.
			SQL="select STATUS from JOB where JOB_NUMBER='"&thissubjobnumber&"' and SHEET_NUMBER='"&thissheetnumber&"' and JOB_TYPE='N'"
			rs.open SQL,conn,1,3
			if not rs.eof then
				if rs("STATUS")<>"1" then
					response.Redirect(pagename+"&errorstring="&asubjobnumber(i)&"没有完成产线生产，不能入库！请重新选择工单。")
				exit for
				end if
			end if
			rs.close
			'get correct start quantity of each sub job
			SQL="select JOB_START_QUANTITY,job_good_quantity from JOB where JOB_NUMBER='"&thissubjobnumber&"' and SHEET_NUMBER='"&thissheetnumber&"' and JOB_TYPE='N'"
			rs.open SQL,conn,1,3
			if not rs.eof then
				if rs("JOB_START_QUANTITY")<>"" then
					input_quantity=input_quantity+cint(rs("JOB_START_QUANTITY"))
				else
				response.Redirect(pagename+"&errorstring=该"&asubjobnumber(i)&"工单开始数量错误，请联系工程师！")
				end if
				
				job_start_qty= rs("JOB_START_QUANTITY")

				store_quantity= rs("job_good_quantity")
				inspect_quantity= rs("job_good_quantity")
				rejectqty= cdbl(rs("JOB_START_QUANTITY"))-cdbl(rs("job_good_quantity"))
		
				
			else
				response.Redirect(pagename+"&errorstring=该"&asubjobnumber(i)&"工单没有在线上生产，请重新扫描子工单！")
			end if
			rs.close
			
				 
			if cdbl(store_quantity)>cint(input_quantity) or cint(inspect_quantity)>cint(input_quantity) then
			response.Redirect(pagename&"&errorstring=入库数量大于投入数量或复检数量大于投入数量，请检查并重新扫描子工单（缺少工作单）！复检="&inspect_quantity&"；入库="&store_quantity&"；投入="&input_quantity)
			end if
		
			jobnumber=thissubjobnumber
			subjobnumber=thissheetnumber
			'check whether store is finished!
			SQL="select STORE_STATUS from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
			rs.open SQL,conn,1,3
			if not rs.eof then
				if rs("STORE_STATUS")="1" then
				response.Redirect(pagename&"&errorstring="+jobnumber+"该工单入库已完成，请重新扫描子工单！")
				end if
			end if
			rs.close
			
			'save store info in job_master
			SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
			rs.open SQL,conn,1,3
				if rs.eof then
					response.Redirect(pagename&"&errorstring="+jobnumber+"该工单信息有错误，请联系IT！")
				else
					part_number_tag=rs("PART_NUMBER_TAG")
					factory_id=rs("FACTORY_ID")
					line_name=rs("LINE_NAME")
					start_quanity=csng(rs("START_QUANTITY"))
					final_scrap_quantity=csng(rs("FINAL_SCRAP_QUANTITY"))
					rework_good_quantity=csng(rs("REWORK_GOOD_QUANTITY"))
					final_input_quantity=csng(rs("FINAL_INPUT_QUANTITY"))
					final_good_quantity=csng(rs("FINAL_GOOD_QUANTITY"))
					final_inspect_quantity=csng(rs("FINAL_INSPECT_QUANTITY"))
				end if
				if isnull(rs("FIRST_STORE_TIME")) or rs("FIRST_STORE_TIME")="" then
					rs("FIRST_STORE_TIME")=now()
				end if
			
				
				assembly_input_quantity=csng(rs("ASSEMBLY_INPUT_QUANTITY"))
				assembly_good_quantity=csng(rs("ASSEMBLY_GOOD_QUANTITY"))
				assembly_scrap_quantity=csng(input_quantity)-csng(store_quantity)
				final_input_quantity=csng(final_input_quantity)+csng(input_quantity)
				final_good_quantity=csng(final_good_quantity)+csng(store_quantity)
				final_inspect_quantity=csng(final_inspect_quantity)+csng(inspect_quantity)
				final_scrap_quantity=csng(rs("final_scrap_quantity"))+csng(rejectqty)
				
				rs("ASSEMBLY_SCRAP_QUANTITY")=csng(rs("ASSEMBLY_SCRAP_QUANTITY"))+csng(assembly_scrap_quantity)
				rs("FINAL_INPUT_QUANTITY")=final_input_quantity
				rs("FINAL_GOOD_QUANTITY")=final_good_quantity
				rs("FINAL_YIELD")=csng(final_good_quantity)/csng(final_input_quantity)
				rs("final_scrap_quantity")=csng(rs("final_scrap_quantity"))+rejectqty
				
				
				rs("LAST_UPDATE_TIME")=now()
					
				if csng(final_good_quantity)>csng(start_quanity) then'
					response.Redirect(pagename&"&errorstring=工单:"+jobnumber+"累计入库数量大于工单开始数量，请确认入库的数量！累计="&final_good_quantity&"；开始="&start_quanity)
				elseif csng(final_scrap_quantity)+csng(final_good_quantity)>csng(start_quanity) then
					response.Redirect(pagename&"&errorstring=工单:"+jobnumber+"累计入库数量加上累计报废数量大于工单开始数量，请确认报废的数量！开始数量="&start_quanity&"；累计入库="&final_good_quantity&"；累计报废="&final_scrap_quantity)
				else
					if csng(final_good_quantity)+csng(final_scrap_quantity)=csng(start_quanity) then ' store is finished
						rs("STORE_STATUS")=1
						rs("COMPLETE_DATE")=now()
					end if
					rs.update
					'trackJobMaster jobnumber,"Store","Manual",store_quantity,"0","0","/goodsstore/store2.asp","228"
				end if
			rs.close
			
		
			InStoreTime=now()
			streo_id=""
			'save store info in job_master_store
			SQL="select * from JOB_MASTER_STORE_PRE WHERE JOB_NUMBER='"+jobnumber+"' AND INSTORE_GUID_ID='"+INSTORE_GUID_ID+"'"
			rs.open SQL,conn,1,3
			if(rs.recordcount=0) then
				rs.addnew
					NID="WH"&NID_SEQ("PROD_STORE")
					rs("NID")=NID
					store_id=NID
					rs("JOB_NUMBER")=jobnumber
					rs("PART_NUMBER_TAG")=part_number_tag
					rs("FACTORY_ID")=factory_id
					rs("LINE_NAME")=line_name
					rs("STORE_CODE")=code
					rs("STORE_TIME")=InStoreTime
					rs("STORE_QUANTITY")=store_quantity
					rs("INSPECT_QUANTITY")=inspect_quantity
					rs("MRB")=0
					rs("MRB_RFP")=0
					if NO_YIELD="1" then
						rs("NO_YIELD")=1
					else
						rs("NO_YIELD")=0
					end if
					rs("INPUT_QUANTITY")=input_quantity
					ontime_yiled=csng(store_quantity)/csng(input_quantity)
					rs("YIELD")=ontime_yiled
					rs("SUB_JOB_NUMBERS")=jobnumber&"-"&repeatstring(subjobnumber,"0",3)
					rs("STORE_TYPE")="N"
					if request.Form("note")<>"" then
						rs("NOTE")=trim(request.Form("note"))
					end if
					rs("INSTORE_GUID_ID")=INSTORE_GUID_ID
					 
				rs.update
			else
					rs("STORE_TIME")=InStoreTime

					rs("INPUT_QUANTITY")=csng(rs("INPUT_QUANTITY"))+csng(job_start_qty)
					rs("STORE_QUANTITY")=csng(rs("STORE_QUANTITY"))+csng(store_quantity)
					rs("INSPECT_QUANTITY")=csng(rs("INSPECT_QUANTITY"))+csng(inspect_quantity)
					rs("INSPECT_QUANTITY")=csng(rs("INSPECT_QUANTITY"))+csng(input_quantity)
					rs("SUB_JOB_NUMBERS")=rs("SUB_JOB_NUMBERS")+","+jobnumber&"-"&repeatstring(subjobnumber,"0",3)
				rs.update
			end if 
			rs.close
			
			'add scrap info 2010-12-17 jack zhang 
			'if (rejectqty<>0) then
			'SQL="select * from JOB_MASTER_SCRAP  WHERE JOB_NUMBER='"+jobnumber+"' AND INSTORE_GUID_ID='"+INSTORE_GUID_ID+"'"
			SQL="select * from JOB_MASTER_SCRAP  WHERE 1=2"
			set rsScrap=server.createobject("adodb.recordset")
			rsScrap.open SQL,conn,1,3
			'if(rsScrap.recordcount=0) then
				rsScrap.addnew
					NID="CH"&NID_SEQ("PROD_SCRAP")
					rsScrap("NID")=NID
					rsScrap("JOB_NUMBER")=jobnumber
					rsScrap("PART_NUMBER_TAG")=part_number_tag
					rsScrap("FACTORY_ID")=factory_id
					rsScrap("LINE_NAME")=line_name
					rsScrap("SCRAP_CODE")=code & "-0000"
					rsScrap("SCRAP_TIME")=InStoreTime
					rsScrap("SCRAP_QUANTITY")=rejectqty
					rsScrap("NOTE")="System"
					rsScrap("REASON")="System"
					rsScrap("STORE_NID")=store_id
					rsScrap("INSTORE_GUID_ID")=INSTORE_GUID_ID
					
					rsScrap("scrap_account")=erp_account
					rsScrap("SCRAP_REASON_ID")=erp_reason
					'add by jack zhang 2012-10-29 for Scrap Series
						set rsScrapSeries=server.createobject("adodb.recordset")
						SQL="select * from SUB_SCRAP_MODEL_SERIES_MAPPING  WHERE PART_NUMBER='"+part_number_tag+"'"
						rsScrapSeries.open SQL,conn,1,3
						refer_series="SP-A03"
						if(rsScrapSeries.recordcount>0) then
								refer_series=refer_series+"-"+rsScrapSeries("SREIES_NAME")+"-00"
						end if 					
						rsScrap("SCRAP_REFERENCE")=refer_series
					'end add
					rsScrap.update
'			else
'					rsScrap("SCRAP_TIME")=InStoreTime
'					rsScrap("SCRAP_QUANTITY")=csng(rsScrap("SCRAP_QUANTITY"))+csng(rejectqty)
'					rsScrap("scrap_account")=erp_account
'					rsScrap("SCRAP_REASON_ID")=erp_reason
					'rsScrap("SCRAP_REFERENCE")=arrerp_refer_str(i)
					
				'rsScrap.update
			'end if 
			
			SQL="select * from JOB_IN_STORE_HISTORY  WHERE JOB_NUMBER='"+jobnumber&"-"&repeatstring(subjobnumber,"0",3)+"'"
			set rsJOB_IN_STORE_HISTORY=server.createobject("adodb.recordset")
			rsJOB_IN_STORE_HISTORY.open SQL,conn,1,3
		 
			if(rsJOB_IN_STORE_HISTORY.recordcount=0) then
				rsJOB_IN_STORE_HISTORY.addnew
				rsJOB_IN_STORE_HISTORY("JOB_NUMBER")=jobnumber&"-"&repeatstring(subjobnumber,"0",3)
				rsJOB_IN_STORE_HISTORY.update
			end if 
			
			
			
			'trackJobMaster jobnumber,"Scrap","Reject","0","0",rejectqty,"/store/store2.asp",batchno&"-417"
			rsScrap.close
			'end if 

		next
end if

 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KES1 Barcode System</title>
<link href="/CSS/Store.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=3;
adCount=0;
function showPopup()
{
	adCount+=1
	if(adCount<timePopup)
	{
	setTimeout("showPopup()",3000);
	document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
	else
	{
	closePopup()
	}
}
function closePopup()
{
	location.href="PickJob.asp";
}
</script>
</head>

<body onLoad="showPopup();"  bgcolor="#FF9933">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2"><div align="center" class="strongred">System will close window in 3 seconds.&nbsp;<br>
      系统将在3秒钟后关闭窗口。 <span id="countinsert"></span></div></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-c-green">Summary of storing 入库成功摘要</td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Operator 操作员:<% =code %></td>
  </tr>
  <tr>
    <td width="37%" height="20">Job Number 工单号</td>
    <td width="63%" height="20"><%if rework="0" then%><% =subjobnumberdisplay %><%else%><% =subjobnumberdisplay %><%end if%></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->