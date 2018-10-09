<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/TrackJobMaster.asp" -->
<!--#include virtual="/Functions/SendEmail.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/Functions/GetFactoryID.asp" -->
<!--#include virtual="/Functions/GetSeriesNew.asp" -->
<!--#include virtual="/Functions/GetERPAccount.asp" -->
<!--#include virtual="/Functions/GetERPReason.asp" -->

<%
jobnumber=request("txtSelectJobNumber")
sheetnumber=request("txtSelectSheetNumber")
errorstring=request("errorstring")	
factory=request("factory")
factory="FA00000002"

if(request.QueryString("action")="2")then
	sql="select * from job where 1=1"
	JobArr=split(jobnumber,",")
	SheetNumberArr=split(sheetnumber,",")
	JobList=" and("
	for i=0 to ubound(JobArr)
		if(JobArr(i)<>"") then
			JobList=JobList+"(job_number='"+JobArr(i)+"' and sheet_number='"+SheetNumberArr(i)+"') OR"
		end if 
	next 
	JobList=JobList+" 1=2)"
	sql=sql+JobList +" order by job_number, sheet_number"
	rs.open sql,conn,1,3
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
    INSTORE_GUID_ID = TypeLib.Guid
	
	
end if 
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KES1 Barcode System</title>
<script language="JavaScript" src="/Functions/JCMD.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<link href="/CSS/List.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script>
	function AutoTransaction()
	{
		if(document.getElementById("txtUserCode").value=="")
		{
			alert('请输入操作人工号!');
			return;
		}
		
		if(document.form1.ERP_Account.selectedIndex==0)
		{
			alert("报废帐号不能为空！");
			ERP_Account.focus()
			return ;
		}
		if(document.form1.ERP_Reason.selectedIndex==0)
		{
			alert("报废理由不能为空！");
			ERP_Reason.focus()
			return ;
		}
		
		
		var count=document.getElementById("count").value;

		if(count>20)
		{
			alert('一次数量不能超过20个工单!');
			return;
		}

		for(var i=0;i<=count;i++)
		{
			document.getElementById("jobnumber").value=document.getElementById("jobnumber").value+document.getElementById("celljobnumber"+i).value+",";
			//var OneRefer;
//			if(document.getElementById("Re_Scrap_Code"+i).selectedIndex==0)
//			{
//				alert("请选择Scrap code");
//				return;
//			}
//			
//			if(document.getElementById("Re_Series"+i).selectedIndex==0)
//			{
//				alert("请选择Seires");
//				return;
//			}
//			
//			if(document.getElementById("Re_DefectCode"+i).selectedIndex==0)
//			{
//				alert("请选择Defect Code");
//				return;
//			}
			
		//	OneRefer="SP-"+document.getElementById("Re_Scrap_Code"+i).options[document.getElementById("Re_Scrap_Code"+i).selectedIndex].value;
//			OneRefer=OneRefer+"-"+document.getElementById("Re_Series"+i).options[document.getElementById("Re_Series"+i).selectedIndex].text;
//			OneRefer=OneRefer+"-"+document.getElementById("Re_DefectCode"+i).options[document.getElementById("Re_DefectCode"+i).selectedIndex].value;
//			document.getElementById("txtRefer").value=document.getElementById("txtRefer").value+OneRefer+",";
		}
		
		
		document.getElementById("btnSubmit").disabled="true";
		form1.action="PickJobConfirm2.asp";
		form1.submit();
	}
	
	function Back()
	{
		location.href="pickjob.asp";
	}
</script>
</head>

<body>
 
<form  method="post" name="form1" target="_self" >
<table width="98%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6"  style="background-color: #339966;font-size: 16px;color: #FFFFFF;"><div align="left">工单自动入库确认</div></td>
  </tr>
   <%if errorstring<>"" then%>
  <tr>
    <td height="20" colspan="6" style="font-size: 20px;color: #FF0000;font-weight: bold;"><div align="center"><%=errorstring%>&nbsp;</div></td>
    </tr>
	<%end if %>
  <tr>
  	<Td colspan="6">
		<table  width="100%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr>
				<td class="t-t-Borrow">序号</td>
				<td class="t-t-Borrow">工单号</td>
				<td class="t-t-Borrow">子工单号</td>
				<td class="t-t-Borrow">线别</td>
				<td class="t-t-Borrow">型号</td>
				<td class="t-t-Borrow">开始时间</td>
				<td class="t-t-Borrow">结束时间</td>
				<td class="t-t-Borrow">开始数量</td>
				<td class="t-t-Borrow">良品数量</td>
				<td class="t-t-Borrow" >报废数量</td>
		<!--		<td class="t-t-Borrow" width="54%">报废Reference</td>-->
			</tr>
			<%
				if(request.QueryString("action")<>"")then 
					i=0
					isVisual=true
					while  not rs.eof
					
			%>
				<tr>
				  <td><%=i+1%></td>
				<td><%=rs("job_number")%><input type="hidden" name="celljobnumber<%=cstr(i)%>" value="<%=rs("job_number")&"-"&repeatstring(rs("sheet_number"),"0",3)%>"></td>
				<td><%=rs("sheet_number")%></td>
				<td><%=rs("line_name")%></td>
				<td><%=rs("part_number_tag")%></td>
				<%
						set rsScrapSeries=server.createobject("adodb.recordset")
						SQL="select * from SUB_SCRAP_MODEL_SERIES_MAPPING  WHERE PART_NUMBER='"+rs("part_number_tag")+"'"
						rsScrapSeries.open SQL,conn,1,3
						if(rsScrapSeries.recordcount<=0) then
							 isVisual=false
						end if 	
				%>
				<td><%=rs("start_time")%></td>
				<td><%=rs("close_time")%></td>
				<td><%=rs("JOB_START_QUANTITY")%></td>
				<td><%=rs("JOB_GOOD_QUANTITY")%></td>
				<td><%=cstr(cint(rs("JOB_START_QUANTITY"))-cint(rs("JOB_GOOD_QUANTITY")))%></td>
					<%
					TotalStartQty=TotalStartQty+cdbl(rs("JOB_START_QUANTITY"))
					TotalGoodQty=TotalGoodQty+cdbl(rs("JOB_GOOD_QUANTITY"))
					TotalScrap=TotalScrap+cdbl(rs("JOB_START_QUANTITY"))-cdbl(rs("JOB_GOOD_QUANTITY"))
				%>
				</tr>
			<%
					i=i+1
					rs.movenext
					wend 
				end if 
			%>

			
			<tr>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">&nbsp;</td>
				<td class="t-t-Borrow">总计</td>
				<td class="t-t-Borrow"><%=cstr(TotalStartQty)%></td>
				<td class="t-t-Borrow"><%=cstr(TotalGoodQty)%></td>
				<td class="t-t-Borrow" ><%=cstr(TotalScrap)%></td>
			</tr>
		</table>
	</Td>
  </tr>
  <Tr>
  	<td width="10%">
		操作人
	</td>
	<td colspan="5">
		 <input type="text" name="txtUserCode" id="txtUserCode" value="<%=UserCode%>" >
	</td>
  </Tr>
  
  <tr>
    <td height="20">Scrapping Account  <span class="red">*</span></td>
    <td height="20" colspan="4"><label>
      <select name="ERP_Account" id="ERP_Account">
	  <option value="">-- 选择帐号 --</option>
	  <%=getERPAccount()%>
      </select>
	  <br><span id="errorinsertERPaccount" class="red"></span>
    </label></td>
  </tr>
  <tr>
    <td height="20">Scrapping Reason   <span class="red">*</span></td>
    <td height="20" colspan="4"><label>
      <select name="ERP_Reason" id="ERP_Reason">
	   <option value="">-- 选择理由 --</option>
	  <%=getERPReason()%>
      </select>
	  <br><span id="errorinsertERPreason" class="red"></span>
    </label></td>
  </tr>
 
    <Tr>
		<td colspan="6">
			<%if isVisual=true then%>
			<input type="button" name="btnSubmit" id="btnSubmit" value="提交" onclick="AutoTransaction()">
		<%else
			response.write "有型号找不到归属的Series，不能入库"
		end if %>
			<input type="button" name="btnBack" id="btnBack" value="返回" onclick="Back()">
		</td>
 	 </Tr>
 	 <Tr>
		<td colspan="6">
			<input type="hidden" name="txtSelectJobNumber" id="txtSelectJobNumber" value="<%=jobnumber%>" >
			<input type="hidden" name="txtSelectSheetNumber" id="txtSelectSheetNumber" value="<%=sheetnumber%>">
			
			<input type="hidden" name="txtRefer" id="txtRefer" value="<%=refer%>" >
			
			<input type="hidden" name="jobnumber" id="jobnumber" value="" >
			<input type="hidden" name="count" value="<%=cstr(i-1)%>">
			<input type="hidden" name="txtINSTORE_GUID_ID" value="<%=INSTORE_GUID_ID%>">
			
			
		</td>
  </Tr>
 
 </table>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td colspan="2" class="darkbule">支持电话：6666(24小时);</td>
    </tr>
  <tr>
    <td><div align="left"><span class="darkbule style1">IP Address: <%=request.ServerVariables("REMOTE_HOST")%></span></div>      </td>
    <td><div align="right"><span class="darkbule style1">System Name: <%=request.ServerVariables("HTTP_HOST")%></span></div></td>
  </tr>
</table>
<input type="hidden" name="count" value="<%=cstr(i-1)%>">
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->