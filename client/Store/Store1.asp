<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Functions/GetSection.asp" -->

<!--#include virtual="/Functions/GetERPAccount.asp" -->
<!--#include virtual="/Functions/GetERPReason.asp" -->

<%response.Expires=0
response.CacheControl="no-cache"
pagename="Store1.asp"
factory=request("factory")
factory_name=session("factory_name")
if factory="" then
	'response.Redirect("SelectFactory.asp")
	SQL="select NID||'-'||FACTORY_NAME AS FACTORY,FACTORY_NAME from FACTORY order by FACTORY_NAME"
	rs.open SQL,conn,1,3
	if not rs.eof then
		factory=rs("FACTORY")
	end if
	rs.close
end if

if instr(factory,"-") > 0 then
	factory_name=right(factory,len(factory)-instr(factory,"-"))
	session("factory_name")=factory_name
	factory=left(factory,instr(factory,"-")-1)
end if
errorstring=request.QueryString("errorstring")
%>
<%

TotalQty=0
TotalRejectQty=0
if request.querystring("batchno")<>"" then
	batchno=request.querystring("batchno")
	batchnoArr=split(batchno,",")
	for i=0 to ubound(batchnoArr)
		if batchnoArr(i)<>"" then
			Flag=""
			SQL="select job_number,subjoblist,actualqty,transactiontype from label_print_history where batchno='"&batchnoArr(i)&"'"
			set rsSubJobList=server.createobject("adodb.recordset")
			rsSubJobList.open SQL,conn,1,3
			if rsSubJobList.recordcount>0 then
				if i=0 then
					MainJob=rsSubJobList("Job_Number")
				elseif MainJob<> rsSubJobList("Job_Number") then
					errorstring="This batch no("&batchnoArr(i)&") is mix job number.该批号("&batchnoArr(i)&")的工单不一致."
					exit for					
				end if
				
				SubJobListArr=split(rsSubJobList("SubJobList"),"-")
				for j=0 to ubound(SubJobListArr)
					subJob = MainJob&"-"&string(3-len(SubJobListArr(j)),"0")&SubJobListArr(j)&","
					if instr(SubJobList,subJob)=0 then
						SubJobList=SubJobList&subJob
					end if					
					if(cint(SubJobListArr(j))>100)then
						Flag="R"
					else
						Flag="N"
					end if 
				next 
				
				TotalQty=TotalQty+cint(rsSubJobList("actualqty"))
				'check OQC
				set rsOQC=server.createobject("adodb.recordset")
				SQL="select 1 from job_oqc_detail where batchno='"&batchnoArr(i)&"' and ispassoqc=1"
				rsOQC.open SQL,conn,1,3
				if rsOQC.eof then
					errorstring="This batch no("&batchnoArr(i)&") is not pass OQC.该批号("&batchnoArr(i)&")没有经过OQC."
					exit for
				end if												
			else
				errorstring="This batch no("&batchnoArr(i)&") is not printed.该批号("&batchnoArr(i)&")没有打印."
				exit for
			end if
			rsSubJobList.close
			set rsSubJobList=nothing						
							
		end if	
	next
	'get scrap qty
	SQL="select nvl(sum(actualqty),0) as scrap_qty from label_print_history where transactiontype=2 and job_number='"&MainJob&"' and is_store=0"
	rs.open SQL,conn,1,3
	if not rs.eof then
		TotalRejectQty=TotalRejectQty+cint(rs("scrap_qty"))
	end if	
	rs.close 
end if
storeType=request("store_type")
'for sub line store
if storeType="subStore" then
	arrJobNumber=split(request("txtSelectJobNumber"),",")
	arrSheetNumber=split(request("txtSelectSheetNumber"),",")
	for i=0 to ubound(arrJobNumber)
		if arrJobNumber(i)<>"" then
			if MainJob="" then
				MainJob=arrJobNumber(i)
			elseif MainJob<>arrJobNumber(i) then
				errorstring="Cannot mix job number.不允许混工单."
				exit for			
			end if
			SubJobList=SubJobList&arrJobNumber(i)&"-"&string(3-len(arrSheetNumber(i)),"0")&arrSheetNumber(i)&","
			strSql="select job_good_quantity,job_defectcode_quantity,job_type from job where job_number='"&arrJobNumber(i)&"' and sheet_number='"&arrSheetNumber(i)&"' "
			rs.open strSql,conn,1,3
			if not rs.eof then
				TotalRejectQty=TotalRejectQty+cint(rs("job_defectcode_quantity"))
				TotalQty=TotalQty+cint(rs("job_good_quantity"))
				if rs("job_type")="N" then
					Flag="N"
				else
					Flag="R"
				end if
			end if
			rs.close
		end if
	next
end if

if MainJob <> "" then
	'Get part number and serial name for job			
	SQL="SELECT SERIES_NAME FROM SERIES_NEW WHERE NID=(SELECT SERIES_GROUP_ID FROM PRODUCT_MODEL WHERE ITEM_NAME=(SELECT PART_NUMBER_TAG FROM JOB_MASTER WHERE JOB_NUMBER='"&MainJob&"'))"
	set rsSerial=server.createobject("adodb.recordset")
	rsSerial.open SQL,conn,1,3
	if rsSerial.recordcount=1 then
		IF (Flag="N")THEN
			Scrap_Ref="SP-A01-"&rsSerial("SERIES_NAME")&"-00"
		else
			Scrap_Ref="SP-A04-"&rsSerial("SERIES_NAME")&"-00"
		end if 
	else
		Scrap_Ref="No Serial Name"
	end if
	rsSerial.close
	set rsSerial=nothing
	
	'get Inventory
	SQL="select completionsubinventory from TBL_MES_LOTMASTER where wipentityname='"+MainJob+"'"		
	set rsInventory=server.createobject("adodb.recordset")		
	rsInventory.open SQL,connTicket,1,3
	if (rsInventory.recordcount>0)then
		Inventory=rsInventory(0)
	end if 
	rsInventory.close
	set rsInventory=nothing	
end if

if errorstring <>"" then
	TotalQty=0
	TotalRejectQty=0
	MainJob=""
	SubJobList=""
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="javascript">
function pageload()
{
	<%if storeType<>"subStore" then%>
		document.form1.batchno.focus();
	<%end if%>
}
function newgetJobStock()
{
	<%if MainJob<>"" then%>	
		getJobStock("<%=MainJob%>")
	<%end if%>
}
function getJobStock(strjobnumber)
{
	if (strjobnumber.length>0)
	{
		if (strjobnumber.search(",")!=-1&&strjobnumber.search("-")!=-1)
		{
			var thisjobnumber=strjobnumber.substring(0,strjobnumber.length-1);
		}
		else
		{
			var thisjobnumber=strjobnumber
		}
		
		var mainjobnumber=thisjobnumber.split(",");

		document.form1.mainjobnumber.value=mainjobnumber[0];
		 
		document.all.JobStockFrame.src="JobStock.asp?mainjobnumber="+mainjobnumber[0];
	}
}
function getSubJobList(strbatchno)
{
	if (strbatchno.length>0)
	{
		form1.action="Store1.asp?batchno="+strbatchno;
		form1.submit();
	}
}

function jobreject()
{
	if(document.all.errorjobnumber.value!=""&&document.all.reason.selectedIndex!=0)
	{
		if(confirm("确认要将"+document.all.errorjobnumber.value+"退回生产线吗？"))
		{
		window.open("JobReject.asp?jobnumber="+document.all.errorjobnumber.value+"&reason="+document.all.reason.options[document.all.reason.selectedIndex].value);
		}
	}
	else
	{
	alert("错误工单和退回原因不得为空！");
	}
}
function SubmitForm()
{
	var jobNumer = document.form1.jobnumber.value;	
	if(!jobNumer){		
		alert("Job Number cannot be blank! \n工单号不得为空!");
		return ;
	}else if(document.form1.code.value=="")
	{
		window.alert("Operator Code cannot be blank!\n操作员工号不能为空！");
		document.form1.code.focus()
		return ;
	}
	
	if(document.form1.quantity.value=="")
	{
		window.alert("Quantity cannot be blank!\n入库数量不能为空！");
		document.form1.quantity.focus()
		return ;
	}
	if(document.form1.rejectqty.value!="0")
	{
		if(document.form1.ERP_Account.selectedIndex==0)
		{
			window.alert("Scrapping Account cannot be blank!\n报废帐号不能为空！");
			document.form1.ERP_Account.focus()
			return;
		}
		if(document.form1.ERP_Reason.selectedIndex==0)
		{
			window.alert("Scrapping Reason cannot be blank!\n报废理由不能为空！");
			document.form1.ERP_Reason.focus();
			return ;
		}
		if(document.form1.ERP_Refer.value=="")
		{
			window.alert("Scrapping Reference cannot be blank!\n报废注释不能为空！");
			document.form1.ERP_Refer.focus();
			return;
		}
	}
	
	//document.form1.action="Store2.asp";
	document.form1.Next.disabled=true;
	document.form1.submit();	
}
function section_change()
{
	with (document.form1)
	{
		if (section.selectedIndex!=0)
		{
			Print.disabled=false;
		}
		else
		{
			Print.disabled=true;
		}
	} 
}
</script>
</head>

<body onLoad="pageload();newgetJobStock()" onKeyPress="keyhandler();"  bgcolor="#339966">
<span id="erroralarm"></span>
<form  method="post" name="form1" target="_self" action="Store2.asp" >
<table width="98%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
   <tr>
    <td height="20" colspan="2" class="t-t-DarkBlue"  align="center"><%=factory_name%>&nbsp; Job Good/Scrap PreStore 工单良品/报废预入库 </td>
  </tr>
  <tr>
    <td height="20" colspan="2" class="t-c-green">Scan BatchNo info 
      扫描批号信息</td>
    </tr>
	<%if errorstring<>"" then%>
  <tr>
    <td height="20" colspan="2" class="strongred"><div align="center"><%=errorstring%></div></td>
    </tr>
	<%end if%>
<%if storeType<>"subStore" then%>	
  <tr>
    <td>Batch No 批号<span class="red">*</span></td>
    <td><input name="batchno" type="text" id="batchno" tabindex="1" onKeyPress="tabhandler(1,this,document.form1.code)" size="50" autocomplete="off" onChange="getSubJobList(this.value);" onFocus="focushandler(this)" onBlur="blurhandler(this)" value="<%if errorstring="" then%><%=batchno%><%else%><%response.Write("")%><%end if%>"><br><span id="errorinsertjobnumber" class="red"></span></td>
  </tr>
<%end if%>   
  <tr>
    <td>Job Number 工单号<span class="red">*</span></td>
    <td><input name="jobnumber" type="text" id="jobnumber" readonly value="<%=SubJobList%>" size="50" ></td>
  </tr>
  <tr>
    <td width="17%">Operator Code 工号 <span class="red">*</span> </td>
    <td>
      <input name="code" type="text" id="code" accesskey="o" tabindex="2" onKeyPress="tabhandler(1,this,document.form1.inspect)" size="10"  autocomplete="off" onFocus="focushandler(this)" onBlur="blurhandler(this)"><br><span id="errorinsertcode" class="red"></span></td>
  </tr>
   <tr>
    <td height="20">Stored Quantity <br>入库数量 <span class="red">*</span></td>
    <td height="20">
		<input name="quantity" type="text" id="quantity" readonly value="<%=TotalQty%>" > 
	</td>
  </tr> 
   <tr>
    <td height="20">Scrap Quantity <br>报废数量 <span class="red">*</span></td>
    <td height="20">
		<input name="rejectqty" type="text" id="rejectqty" readonly value="<%=TotalRejectQty%>" > </td>
  </tr>
  
  <tr>
    <td height="20">Scrap Account <br>报废账号 <span class="red">*</span></td>
    <td height="20" colspan="6">
      <select name="ERP_Account" id="ERP_Account">
	  <option value="0"></option>
	  <%=getERPAccount()%>
      </select>
	  <br><span id="errorinsertERPaccount" class="red"></span>
    </td>
  </tr>
  <tr>
    <td height="20">Scrap Reason<br>报废原因<span class="red">*</span></td>
    <td height="20" colspan="6">
      <select name="ERP_Reason" id="ERP_Reason">
	   <option value="0"></option>
	  <%=getERPReason()%>
      </select>
	  <br><span id="errorinsertERPreason" class="red"></span>
    </td>
  </tr>
  <tr>
    <td height="20">Scrap Reference<br>报废参照 <span class="red">*</span></td>
    <td height="20" colspan="6"><input name="ERP_Refer" type="text" id="ERP_Refer" onFocus="focushandler(this)" onBlur="blurhandler(this)" value="<%=Scrap_Ref%>"><br><span id="errorinsertERPrefer" class="red"></span></td>
  </tr>
  <tr>
    <td height="20">Remarks 注释</td>
    <td height="20" ><input name="note" type="text" size="80" id="note" onFocus="focushandler(this)" onBlur="blurhandler(this)"></td>
  </tr>
  
  <tr>
    <td height="20">Inventory 入库库位</td>
    <td height="20"><input name="inventory" type="text" id="inventory" value="<%=Inventory%>"></td>
  </tr>
  
  <tr>
    <td height="20" colspan="2" >Stored Info 工单入库信息</td>
    </tr>
  <tr>
    <td height="20" colspan="2" class="red"><iframe src="" width="100%" height="220" scrolling="auto" frameborder="0" id="JobStockFrame"></iframe></td>
    </tr>
  
  <tr>
    <td height="20" colspan="2" valign="middle" align="center">
		<input type="hidden" name="store_type" id="store_type" value="<%=storeType%>" >
        <input name="factory" type="hidden" id="factory" value="<%=factory%>">
        <input name="mainjobnumber" type="hidden" id="mainjobnumber" value="<%=MainJob%>">
        <input name="Next" type="button" accesskey="n" tabindex="5" value="Next 下一步" onclick="SubmitForm()">       
		&nbsp;
		<select name="section" id="section" onChange="section_change()">
		<option value=""></option>
		<%=getSection("OPTION",null," where S.FACTORY_ID='"&factory&"'",null,null)%>
        </select>
        <input name="Print" type="button" disabled id="Print" onClick="javascript:window.open('PrintStoreList.asp?factory=<%=factory%>&section='+document.form1.section.options[document.form1.section.selectedIndex].value)"
		 value="Print 打印">
		&nbsp;               
        <input name="Search" type="button" id="Search" value="Query 查询" onClick="javascript:window.open('StoreRecords.asp?factory=<%=factory%>')">
        &nbsp;
		<input name="btnSubStore" type="button" id="btnSubStore" onClick="javascript:location.href='PickJob.asp'" value="HFL Store 入库">
		</td>
   </tr> 
   <tr>
    <td height="20" colspan="2" valign="middle" align="center" >
		Job Number 工单号 &nbsp;<input name="searchjobnumber" type="text" id="searchjobnumber" size="10" >
      <input name="SearchPrintList" type="button" id="SearchPrintList" onClick="javascript:if(document.all.searchjobnumber.value){window.open('SearchPrintList.asp?jobnumber='+document.all.searchjobnumber.value)}else{alert('Please input job number.\n请输入工单号。');document.all.searchjobnumber.focus();}" 
	  value="Print Query 打印查询">
       &nbsp;  
	  <input name="Close" type="button" id="Close" tabindex="7" onClick="javascript:window.close()" value="Close 关闭">          
    </td>
  </tr> 
</table>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td colspan="2" class="darkbule">支持电话：18910736955 (24小时);</td>
    </tr>
  <tr>
    <td><div align="left"><span class="darkbule style1">IP Address: <%=request.ServerVariables("REMOTE_HOST")%></span></div>      </td>
    <td><div align="right"><span class="darkbule style1">System Name: <%=request.ServerVariables("HTTP_HOST")%></span></div></td>
  </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->