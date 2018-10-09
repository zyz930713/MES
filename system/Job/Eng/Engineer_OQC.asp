<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
batchNo=request("batchNo")
jobNumber = batchNo
if instr(batchNo,"-") >0 then
	jobNumber=left(batchNo,instr(batchNo,"-")-1)
end if

'
'SQL="Update  JOB_RETEST set ISRELEASETOOQC=0 where BATCHNO='1'"
'rsQ.open SQL,conn,1,3


if request.QueryString("Action")="2" then
	nid=request.Form("nid")
	action=request.Form("action")
	'update by alice on 2011-1-26
	SQL="Update JOB_RETEST set ENGINEERCODE='"&session("code")&"',ISLOWYIELDHOLD=0,NID='"&nid&"',Action='"&action&"' where BATCHNO='"&batchNo&"'"
	'response.Write sql
'	response.End()
	set rsE=server.createobject("adodb.recordset")
	rsE.open SQL,conn,1,3
	response.write "<script>window.alert('Release to Retest successfully!');location.href='OQC_List.asp';</script>"
end if

if request.QueryString("Action")="3" then
	SQL="Update JOB_RETEST set ENGINEERCODE='"&session("code")&"',ISRELEASETOOQC=1 where BATCHNO='"&batchNo&"'"
	set rsE=server.createobject("adodb.recordset")
	rsE.open SQL,conn,1,3
	
	'Add by jack zhang 2010-9-7
		PrintName="KES-TESTZEBRA"
		set PrintCtl=server.createobject("PrintClass.PrintCtl")
		SQL="select count(RejectQty),max(modelname) from Job_ReTest_Detail where batchno='"+batchNo+"'"
		set rsTotalRejectQty=server.createobject("adodb.recordset")
		rsTotalRejectQty.open SQL,conn,1,3
		if rsTotalRejectQty.recordcount>0 then
			TotalRejectQty=rsTotalRejectQty(0).value
			model=rsTotalRejectQty(1).value
		end if
		
		BatchNoArr=split(batchNo,"-")
		NewBatch=BatchnoArr(0)+"-"+BatchnoArr(1)+"-6"   '6 Retest Scrap
		if cint(TotalRejectQty)>0 then
			ReturnCode=PrintCtl.PrintScrap_ReTest(PrintName, 1,NewBatch, "", TotalRejectQty,Model)
			if(ReturnCode="OK") then
				SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"
				set rs_AddHistory=server.CreateObject("adodb.recordset")
				rs_AddHistory.open SQL,conn,1,3
				rs_AddHistory.addnew
					rs_AddHistory("JOB_NUMBER")=jobNumber
					rs_AddHistory("BATCHNO")=NewBatch
					rs_AddHistory("QTY")=TotalRejectQty
					rs_AddHistory("SUBJOBLIST")=batchNo
					rs_AddHistory("ACTUALQTY")=TotalRejectQty
					rs_AddHistory("OPERATOR")=session("code")
					rs_AddHistory("REMARK")=""
					rs_AddHistory("PRINTTIME")=NOW()
					rs_AddHistory("TRANSACTIONTYPE")=6
					Set TypeLib = CreateObject("Scriptlet.TypeLib")
   					Guid = TypeLib.Guid
					rs_AddHistory("id")=Guid
				rs_AddHistory.update
							
				IsPrintLabel="1"
			else
				response.write("<script>window.alert('Print ReTest Scrap label failed!')</script>")
			end if
	end if
	'End add
	if IsPrintLabel="1" then
		response.write "<script>window.alert('Release to OQC successfully,Please get Retest Scrap label!');location.href='Engineer_OQC.asp';</script>"
	else
		response.write "<script>window.alert('Release to OQC successfully!');location.href='Engineer_OQC.asp';</script>"
	end if
	
	
end if

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="../../CSS/General.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript">
	/*function FindData()
	{
		document.form1.action="Engineer_OQC.asp?Action=1";
		document.form1.submit();
	}*/
	function ToRetest()
	{
	 
		if(document.form1.nid.value!="" && document.form1.action.value!="")
		{
			 
			document.form1.submit();
		}
		else
		{
			alert("请输入NCMR编号和相应的处理措施！");
		}
		
	}
	/*function ToOQC()
	{
		document.form1.action="Engineer_OQC.asp?Action=3";
		document.form1.submit();
	}*/
</script>
</head>

<body>
<form method="post" name="form1" action="Engineer_OQC.asp?Action=2">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">

<tr class="t-t-Borrow" >
		<td  colspan="6" align="left" height="20">Retest</td>
	</tr>
	<%
		
	 	'if request.QueryString("Action")="1" then
			SQL="select JR.*,JD.*,detail.* from JOB_RETEST JR left join JOB_RETEST_DETAIL JD on JR.BATCHNO=JD.BATCHNO "
			SQL=SQL+" left join JOB_RETEST_DETAIL detail on detail.DETAILID=JD.DETAILID "
			SQL=SQL+" where JR.BATCHNO='"&batchNo&"' and ISLOWYIELDHOLD=1 and ISRELEASETOOQC=0"
			set rsQ=server.createobject("adodb.recordset")
			rsQ.open SQL,conn,1,3

			if rsQ.recordcount>0 then	
	%>
	<tr  height="20">
		<td class="t-c-GrayLight">批次号</td>
		<td><%=batchNo%></td>
		<td class="t-c-GrayLight">型号</td>
		<%
			set rsM=server.createobject("adodb.recordset")
			SQL="select * from job_master where job_number='"&jobNumber&"'"
			ModelName=""
			rsM.open SQL,conn,1,3
			if rsM.recordcount>0 then				 
			 ModelName=rsM("PART_NUMBER_TAG")
			end if 
			%>
			
		<td>
			<%=ModelName%>
			<input type="hidden" id="Model" name="Model" value="<%=ModelName%>" />		</td>
		<td class="t-c-GrayLight">开始总数量</td>
		<td><%=rsQ("STARTQTY")%></td>
		
			
		<%
			SQL="SELECT RetestRatio FROM GENERAL_SETTING WHERE ModelName='"&ModelName&"' "

			set rsRetestRatio=server.createobject("adodb.recordset")
			rsRetestRatio.open SQL,conn,1,3
			if rsRetestRatio.recordcount>0 then
				RetestRatio=rsRetestRatio(0).value
				SampleQty=cint(startQty*RetestRatio)
			end if 
			
		%>
	</tr>
	<tr  height="20">
		<td class="t-c-GrayLight">抽样数（测试数量）</td>
		<td><%=rsQ("SAMPLEQTY")%></td>
		<td class="t-c-GrayLight">抽样比例</td>
		<td><%=rsQ("SAMPLERATE")%></td>
		<td class="t-c-GrayLight">合格数量</td>
		<td><%=rsQ("GOODQTY")%></td>
	</tr>
	<tr  height="20">		
		<td class="t-c-GrayLight">不合格数量</td>
		<td><%=rsQ("REJECTQTY")%></td>
		<td class="t-c-GrayLight">Line Loss 数量</td>
		<td> <%=rsQ("LINELOSSQTY")%></td>
		<td>OP Code</td>
		<td><%=rsQ("STARTOPERATOR")%></td>
	</tr>
	<tr>	</tr>
	<tr class="t-t-Borrow"  height="20">
		<td colspan="6">不良信息</td>
	</tr>
	<%
	SQL="Select a.*,b.*,d.DEFECT_CHINESE_NAME,d.Defect_Name from job_retest_detail a, job_retest_defect b, defectcode d"
	SQL=SQL+" where a.detailid=b.detailid(+) and b.defectcodeid=d.nid(+) "
	SQL=SQL+" and a.detailid=(select max(detailid) from job_retest_detail c where c.batchno=a.batchno) "
	SQL=SQL+" and a.batchno='"&batchNo&"' order by b.defectshowsequence"
	
	rs.open SQL,conn,1,3
	if rs.recordcount >0 then
	for i=1 to rs.recordcount
		 
	%>
	<tr  height="20">
		<td class="t-c-GrayLight"><%=rs("Defect_Name")%><%=rs("Defect_Chinese_Name")%></td>
		<td><%=rs("DEFECTQTY")%>
		<%
			if i< rs.recordcount then
			i=i+1  
			
			rs.movenext	
			 
		%></td>
		<td class="t-c-GrayLight"><%=rs("Defect_Name")%><%=rs("Defect_Chinese_Name")%></td>		
		<td width="10%"><%=rs("DEFECTQTY")%></td>
		<% else %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<%end if%>
		<%
			if i< rs.recordcount then
			i=i+1  
			
			rs.movenext	
			
		%>
		
		<td class="t-c-GrayLight"><%=rs("Defect_Name")%><%=rs("Defect_Chinese_Name")%></td>
		<td width="10%"><%=rs("DEFECTQTY")%></td>
		<% else %>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<%end if%>
	</tr>
	<%	
	rs.movenext	 
	next
	end if
	%>
	<input type="hidden" id="count" name="count" value="<%=i-1%>" />
	<tr class="t-t-Borrow"  height="20">
		<td colspan="6">确认</td>
	</tr>
	<tr>
		<td>NCMR 编号</td>
		<td><input type="text" id="nid" name="nid" /></td>
		<td>处理措施</td>
		<td colspan="3"><textarea name="action" id="aciton" cols="50" rows="3"></textarea></td>
	</tr>
	<tr  height="20">
		
	  <td colspan="6" align="center">
	  <input type="hidden" id="batchNo" name="batchNo" value="<%=batchNo%>" />
		  <input name="Retest" type="button" class="t-c-greenCopy" id="Retest" value="Release " onclick="ToRetest(); " />
		
		 
	</tr>
</table>
<%
		end if
'	end if
%>
</form>

</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->