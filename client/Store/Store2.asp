<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/TrackJobMaster.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/Functions/GetFactoryID.asp" -->
<%
pagename="Store2.asp"
code=trim(request.Form("code"))
factory=request.Form("factory")
MRB=request.Form("MRB")
MRB_RFP=request.Form("MRB_RFP")
NO_YIELD=request.Form("NO_YIELD")
subjobnumber=ucase(trim(request.Form("jobnumber")))
subjobnumber=left(subjobnumber,len(subjobnumber)-1)
session("JOB_NUMBER")=subjobnumber
asubjobnumber=split(subjobnumber,",")
store_quantity=cdbl(request.Form("quantity"))
'inspect_quantity=cdbl(request.Form("inspect"))
inspect_quantity=store_quantity

rejectqty=cdbl(request.Form("rejectqty"))
erp_account=request.Form("ERP_Account")
erp_reason=request.Form("ERP_Reason")
erp_refer=request.Form("ERP_Refer")

jobnumber=request("mainjobnumber")

rework=0

if rework="0" then
	input_quantity=0
	for i=0 to ubound(asubjobnumber)
		'check have stored or not
		SQL="select * from JOB_MASTER_STORE_PRE where SUB_JOB_NUMBERS = '"&trim(asubjobnumber(i))&"'"
		rs.open SQL,conn,1,3
		if not rs.eof then
			response.Redirect("Store1.asp?factory="&factory&"&errorstring="&asubjobnumber(i)&"�Ѿ���"&rs("STORE_CODE")&"��"&rs("STORE_TIME")&"����"&rs("JOB_NUMBER")&"��⣡������ѡ�񹤵���")
			exit for
		end if
		rs.close
		
		'get job number, sheet number and job type of each sub job
		if instr(trim(asubjobnumber(i)),"-")=0 then
			response.Redirect("Store1.asp?factory="&factory&"&errorstring="&asubjobnumber(i)&"������ʽ����ȱ��-��������ɨ���ӹ�����")
		end if
		
		sheetnumber=cint(right(asubjobnumber(i),len(asubjobnumber(i))-len(jobnumber)-1))
		if sheetnumber >=500 then
			rework="1"
		end if
				
		SQL="select STATUS,JOB_START_QUANTITY from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"'"
		
		rs.open SQL,conn,1,3
		if not rs.eof then
			'check is sub sheet finsihed prodution in production lines.
			if rs("STATUS")<>"1" then
				response.Redirect("Store1.asp?factory="&factory&"&errorstring="&asubjobnumber(i)&"û����ɲ���������������⣡������ѡ�񹤵���")
				exit for
			end if
			'get correct start quantity of each sub job
			if rs("JOB_START_QUANTITY")<>"" then
				input_quantity=input_quantity+cdbl(rs("JOB_START_QUANTITY"))
			else
				response.Redirect("Store1.asp?factory="&factory&"&errorstring=��"&asubjobnumber(i)&"������ʼ������������ϵ����ʦ��")
			end if			
		else
			response.Redirect("Store1.asp?factory="&factory&"&errorstring=��"&asubjobnumber(i)&"����û��������������������ɨ���ӹ�����")			
			response.End()
		end if
		rs.close						
	next

	if cdbl(store_quantity)>cdbl(input_quantity) or cdbl(inspect_quantity)>cdbl(input_quantity) then
		response.Redirect("Store1.asp?factory="&factory&"&errorstring=�����������Ͷ�������򸴼���������Ͷ�����������鲢����ɨ���ӹ�����ȱ�ٹ�������������="&inspect_quantity&"�����="&store_quantity&"��Ͷ��="&input_quantity)
	end if
end if

'save store info in job_master
SQL="select * from JOB_MASTER where JOB_NUMBER='"&jobnumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
	response.Redirect("Store1.asp?factory="&factory&"&errorstring=�������������ڣ�����ϵIT����ʦ��")
else
	'check whether store is finished!
	if rs("STORE_STATUS")="1" then
		response.Redirect("Store1.asp?factory="&factory&"&errorstring=�ù����������ɣ�������ɨ���ӹ�����")
	end if
	part_number_tag=rs("PART_NUMBER_TAG")
	factory_id=rs("FACTORY_ID")
	line_name=rs("LINE_NAME")
	start_quanity=cdbl(rs("START_QUANTITY"))
	final_scrap_quantity=cdbl(rs("FINAL_SCRAP_QUANTITY"))
	final_good_quantity=cdbl(rs("FINAL_GOOD_QUANTITY"))
	final_inspect_quantity=cdbl(rs("CONFIRM_GOOD_QUANTITY"))
	
	if isnull(rs("FIRST_STORE_TIME")) or rs("FIRST_STORE_TIME")="" then
		rs("FIRST_STORE_TIME")=now()
	end if
	
	final_scrap_quantity=final_scrap_quantity+rejectqty+(store_quantity-inspect_quantity)
	final_inspect_quantity=final_inspect_quantity+inspect_quantity
	final_good_quantity=final_good_quantity+store_quantity
	rs("FINAL_SCRAP_QUANTITY")=final_scrap_quantity
	rs("FINAL_GOOD_QUANTITY")=final_good_quantity
	rs("FINAL_INPUT_QUANTITY")=final_good_quantity+final_scrap_quantity
	rs("CONFIRM_GOOD_QUANTITY")=final_inspect_quantity
	rs("FINAL_YIELD")=final_good_quantity/(final_good_quantity+final_scrap_quantity)
	rs("LAST_UPDATE_TIME")=now()
	
	
	if final_scrap_quantity+final_inspect_quantity>start_quanity then
		response.Redirect("Store1.asp?factory="&factory&"&errorstring=�ۼ�������������ۼƱ����������ڹ�����ʼ��������ȷ�ϱ��ϵ���������ʼ����="&start_quanity&"���ۼ����="&final_good_quantity&"���ۼƱ���="&final_scrap_quantity)
	end if
	'Eng Borrow
	EngBorrow=getEngBorrow(jobnumber)					
	SQL="select * from JOB_MASTER_SCRAP_PRE WHERE job_number='"+jobnumber+"' and  trim(NOTE)='Eng Borrow'"
	set rsEngBorrowCheckEng=server.createobject("adodb.recordset")
	rsEngBorrowCheckEng.open SQL,conn,1,3
	if(rsEngBorrowCheckEng.recordcount=0) then
		summary_quantity=final_inspect_quantity+final_scrap_quantity+EngBorrow
	else
		summary_quantity=final_inspect_quantity+final_scrap_quantity
	end if
	rsEngBorrowCheckEng.close

	'��һ��ֻҪϵͳƽ��Ϳ��Թع�����������ϵͳ���good��Ϊ�Ƚ�����
	if summary_quantity=start_quanity then ' store is finished
		rs("STORE_STATUS")=1
		rs("COMPLETE_DATE")=now()
		rs("AVERAGE_CYCLE_TIME")=cstr(datediff("n",rs("INPUT_TIME"),now()))
	end if
	rs.update
end if
rs.close

InStoreTime=now()
streo_id=""
'save store info in job_master_store
SQL="SELECT * FROM JOB_MASTER_STORE_PRE where 1=2"
rs.open SQL,conn,1,3
rs.addnew
	NID="WH"&NID_SEQ("PROD_STORE")
	rs("NID")=NID
	streo_id=NID
	rs("JOB_NUMBER")=jobnumber
	rs("PART_NUMBER_TAG")=part_number_tag
	rs("FACTORY_ID")=factory_id
	rs("LINE_NAME")=line_name
	rs("STORE_CODE")=code+ "-0000"
	rs("STORE_TIME")=InStoreTime
	rs("STORE_QUANTITY")=store_quantity
	rs("INSPECT_QUANTITY")=inspect_quantity
	if MRB="1" then
		rs("MRB")=1
	else
		rs("MRB")=0
	end if
	if MRB_RFP="1" then
		rs("MRB_RFP")=1
	else
		rs("MRB_RFP")=0
	end if
	if NO_YIELD="1" then
		rs("NO_YIELD")=1
	else
		rs("NO_YIELD")=0
	end if

	if rework="0" then
		rs("INPUT_QUANTITY")=input_quantity
		ontime_yiled=store_quantity/input_quantity
		rs("YIELD")=ontime_yiled
		rs("SUB_JOB_NUMBERS")=subjobnumber
		rs("STORE_TYPE")="N"
	else
		rs("YIELD")=0
		if(rework=1)then
			rs("STORE_TYPE")="R"
		end if
		if(rework=2)then
			rs("STORE_TYPE")="S"
		end if
		if(rework=3)then
			rs("STORE_TYPE")="C"
		end if
		rs("SUB_JOB_NUMBERS")=subjobnumber
	end if
	if request.Form("note")<>"" then
		rs("NOTE")=trim(request.Form("note"))
	end if
	rs("PRINT_TIMES")=1
rs.update
rs.close

SQL="select * from JOB_MASTER_SCRAP_PRE where 1=2"
set rsScrap=server.createobject("adodb.recordset")
rsScrap.open SQL,conn,1,3
rsScrap.addnew
	NID="CH"&NID_SEQ("PROD_SCRAP")
	rsScrap("NID")=NID
	rsScrap("JOB_NUMBER")=jobnumber
	rsScrap("PART_NUMBER_TAG")=part_number_tag
	rsScrap("FACTORY_ID")=factory
	rsScrap("LINE_NAME")=line_name
	rsScrap("SCRAP_CODE")=code & "-0000"
	rsScrap("SCRAP_TIME")=InStoreTime
	rsScrap("SCRAP_QUANTITY")=rejectqty
	rsScrap("NOTE")="System"
	rsScrap("REASON")="System"
	rsScrap("scrap_account")=erp_account
	rsScrap("SCRAP_REASON_ID")=erp_reason
	rsScrap("SCRAP_REFERENCE")=erp_refer	
	rsScrap("STORE_NID")=streo_id
rsScrap.update
rsScrap.close

'���ϵͳ��ʵ�ﲻ����
if store_quantity<>inspect_quantity then
	SQL="select * from JOB_MASTER_SCRAP_PRE where 1=2"
	set rsScrap1=server.createobject("adodb.recordset")
	rsScrap1.open SQL,conn,1,3
	rsScrap1.addnew
	NID="CH"&NID_SEQ("PROD_SCRAP")
	rsScrap1("NID")=NID
	rsScrap1("JOB_NUMBER")=jobnumber
	rsScrap1("PART_NUMBER_TAG")=part_number_tag
	rsScrap1("FACTORY_ID")=factory
	rsScrap1("LINE_NAME")=line_name
	rsScrap1("SCRAP_CODE")=code & "-0000"
	rsScrap1("SCRAP_TIME")=InStoreTime
	rsScrap1("SCRAP_QUANTITY")=store_quantity-inspect_quantity
	rsScrap1("NOTE")="Gap"
	rsScrap1("REASON")="Gap"
	rsScrap1("scrap_account")=erp_account
	rsScrap1("SCRAP_REASON_ID")=erp_reason
	rsScrap1("SCRAP_REFERENCE")=erp_refer
	rsScrap1("STORE_NID")=streo_id
	rsScrap1.update
	rsScrap1.close
	
end if 

'���Eng Borrowû����������
SQL="select * from JOB_MASTER_SCRAP_PRE WHERE job_number='"+jobnumber+"' and  NOTE='Eng Borrow'"
set rsScrapCheckEng=server.createobject("adodb.recordset")
rsScrapCheckEng.open SQL,conn,1,3
if(rsScrapCheckEng.recordcount=0) then
	if cdbl(EngBorrow)>0 then
		SQL="select * from JOB_MASTER_SCRAP_PRE"
		set rsScrap2=server.createobject("adodb.recordset")
		rsScrap2.open SQL,conn,1,3
		rsScrap2.addnew
		NID="CH"&NID_SEQ("PROD_SCRAP")
		rsScrap2("NID")=NID
		rsScrap2("JOB_NUMBER")=jobnumber
		rsScrap2("PART_NUMBER_TAG")=part_number_tag
		rsScrap2("FACTORY_ID")=factory
		rsScrap2("LINE_NAME")=line_name
		rsScrap2("SCRAP_CODE")=code & "-0000"
		rsScrap2("SCRAP_TIME")=InStoreTime
		rsScrap2("SCRAP_QUANTITY")=EngBorrow
		rsScrap2("NOTE")="Eng Borrow"
		rsScrap2("REASON")="Eng Borrow"
		rsScrap2("scrap_account")=erp_account
		rsScrap2("SCRAP_REASON_ID")=erp_reason
		rsScrap2("SCRAP_REFERENCE")=erp_refer
		rsScrap2("STORE_NID")=streo_id
		rsScrap2.update
		rsScrap2.close
		
		SQL="update JOB_MASTER set FINAL_SCRAP_QUANTITY=FINAL_SCRAP_QUANTITY+"+cstr(EngBorrow)+" where job_number='"+jobnumber+"'"
		set rsMaster2=server.createobject("adodb.recordset")
		rsMaster2.open SQL,conn,1,3
	end if 
end if 

'Add job Inventory
scrap_quantity = rejectqty + (store_quantity-inspect_quantity) + EngBorrow
SQL="select job_number,part_number,good_qty,scrap_qty,lm_time,lm_user from job_inventory where job_number='"&jobnumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
	rs.addnew
	rs("job_number")=jobnumber
	rs("part_number")=part_number_tag
	rs("good_qty")=store_quantity
	rs("scrap_qty")=scrap_quantity
else
	rs("good_qty")=cdbl(rs("good_qty"))+store_quantity
	rs("scrap_qty")=cdbl(rs("scrap_qty"))+scrap_quantity
end if
rs("lm_time")=now()
rs("lm_user")=code
rs.update
rs.close

'update print history
SQL="update label_print_history set is_store=1 where job_number='"&jobnumber&"'"
rs.open SQL,conn,1,3

'add by jack zhang 2010-12-23
function getEngBorrow(JobNumber)
	set rsSub=server.CreateObject("adodb.recordset")
	SQL="SELECT decode(SUM(UNITQTY),null,0,SUM(UNITQTY)) FROM ENG_BORROW WHERE JOBNUMBER='"&JobNumber&"'"
	rsSub.open SQL,conn,1,3
	qty=cdbl(rsSub(0))
	rsSub.close
	
	getEngBorrow=qty
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript">
timePopup=5;
adCount=0;
function showPopup()
{
	adCount+=1
	if(adCount<timePopup)
	{
	setTimeout("showPopup()",1000);
	document.all.countinsert.innerText="("+(timePopup-adCount)+")";
	}
	else
	{
	closePopup()
	}
}
function closePopup()
{
	<%if request("store_type") = "subStore" then %>
		location.href="/Store/Pickjob.asp";
	<%else%>
		location.href="/Store/Store1.asp?factory=<%=factory%>";
	<%end if%>
}
</script>
</head>

<body onLoad="showPopup();" bgcolor="#339966">
<table width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="2"><div align="center" class="strongred">System will close window in 5 seconds.&nbsp;<br>
      ϵͳ����5���Ӻ�رմ��ڡ� <span id="countinsert"></span></div></td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-c-green">Summary of storing ���ɹ�ժҪ</td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-t-Borrow">Operator ����Ա:<% =code %></td>
  </tr>
  <tr>
    <td width="37%" height="20">Job Number ������</td>
    <td width="63%" height="20"><%if rework="0" then%><% =subjobnumber %><%else%><% =jobnumber %><%end if%></td>
  </tr>
  <tr>
    <td height="20">Input Quantity  Ͷ������ </td>
    <td height="20"><% =input_quantity%>
    &nbsp;</td>
  </tr>
  <tr>
    <td height="20">Inspect Quantity �������� </td>
    <td height="20"><% =inspect_quantity%></td>
  </tr>
  <tr>
    <td height="20">Good Quantity �ϸ����� </td>
    <td height="20"><% =store_quantity%></td>
  </tr>
  <tr>
    <td height="20">Scrap Quantity ��������</td>
    <td height="20"><% =scrap_quantity %>
    &nbsp;</td>
  </tr>
  <tr>
    <td height="20">Ontime Yiled ��ʱ����</td>
    <td height="20"><% =formatpercent(ontime_yiled,2)%></td>
  </tr>
  <tr>
    <td height="20">Scrapping Account</td>
    <td height="20"><% =erp_account %></td>
  </tr>
  <tr>
    <td height="20">Scrapping Reason </td>
    <td height="20"><% =erp_reason %></td>
  </tr>
  <tr>
    <td height="20">Scrapping Reference</td>
    <td height="20"><% =erp_refer %></td>
  </tr>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->