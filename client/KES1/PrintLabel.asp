<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">

	function SaveData()
	{
		//document.form1.computername.value="KES-IT-05";
	 	if(document.form1.computername.value=="")
		{
			window.alert("Can not get computer name of this machine,Please contact with IT. \n没有能获取到机器名，请联系IT！");
			return;
		}
		
		var jobString=document.form1.txtSubJobList.value;	
		if(jobString.substr(jobString.length-1,jobString.length)==",")
		{
			jobString=jobString.substr(0,jobString.length-1);
		}
		 
		var Jobs=""
		var check=0;
		if(jobString!="")
		{
			
			job=jobString.split(",");		
			//add by jack zhang 2011-1-4
			for(var i=0;i<job.length;i++)
			{
				for(var j=i+1;j<job.length;j++)
				{
					if(job[i]==job[j])
					{
						window.alert("请不要输入相同的子工单!");
						return;
					}
				}
			}		
			//end add	 
			if(job.length>8)
			{
				window.alert("请不要输入超过8个子工单!");
				return;
			}

			for(var i=0;i<job.length;i++)
			{		
				var newJob=document.form1.newJob.value;
				var jobA=job[i].split("-");
				if (jobA[1]=="E" || jobA[1]=="R")
				{
					job1=jobA[0]+"-"+jobA[1];
					newJob=newJob+parseFloat(jobA[2]).toString()+"-";
				}
				else
				{
					job1=jobA[0];
					newJob=newJob+parseFloat(jobA[1]).toString()+"-";
				}
				if(job1!="")
				{
					document.form1.batchNo.value=job1;
					 
				}
				for(var j=i+1;j<job.length;j++)
				{
					jobB=job[j].split("-");
					if (jobB[1]=="E" || jobB[1]=="R")
					{
						job2=jobB[0]+"-"+jobB[1];
					}
					else
					{
						job2=jobB[0];
					}
					if(job1!=job2 && job2!="")
					{
						check=1;	
						alert("请输入相同的主工单！");		 
						break;
					}
					 
				}
				if(check==1)
				{
					break;
				}
				else
				{
					document.form1.newJob.value=newJob;
				}
			}
			
			//Add by jack zhang 2010-9-16 slapping and retest can't be printted with normail.
			var SubJobArr=document.form1.newJob.value.split("-");
			var IsNormal=false;
			//var NotNormal=false;
			var IsRework=false;
			var IsChangeModel=false;
			
			for(var i=0;i<=SubJobArr.length-1;i++)
			{
				if (parseFloat(SubJobArr[i])<500)
				{
					IsNormal=true;
				}
				if (parseFloat(SubJobArr[i])>=500 && parseFloat(SubJobArr[i])<700)
				{
					IsRework=true;
				}
				if (parseFloat(SubJobArr[i])>=700)
				{
					IsChangeModel=true;
				}
				
			}
			if((IsNormal==true && IsRework==true) ||(IsNormal==true && IsChangeModel==true)||(IsRework==true && IsChangeModel==true))
			{
				window.alert("不通类型的工单不能一起打印Label!");
			}
			else
			{
			//End add jack zhang 2010-9-16 slapping and retest can't be printted with normail.
				Jobs=document.form1.newJob.value;
				Jobs=Jobs.substr(0,Jobs.length-1)
				document.form1.newJob.value=Jobs;			 
				if(check!=1)
				{ 
					document.form1.action="PrintLabel2.asp";
					document.form1.submit();
					
					
				}
			}
		}
		else
		{
			alert("请填写完整的打印信息！");
		}
	}
	 
</script>
</head>
<body onLoad="form1.txtSubJobList.focus();" bgcolor="#339966">

<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">Label Print(标签打印)</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<Tr>
		<td height="20" colspan="2"><%=request.QueryString("word")%></td>
	</Tr>
	<%end if%>
    <tr>
      <td >Sub Job Number 子工单号</td>
      <td ><input name="txtSubJobList" type="text" id="txtSubJobList" size="50" value=<%=SubJob%> ></td>
    </tr>
	<tr>
      <td >Operator Code 工号</td>
      <td ><input name="txtOperatorCode" type="text" id="txtOperatorCode" value=<%=OperatorCode%> ></td>
    </tr>
	<tr><td colspan="2">&nbsp;</td></tr>
    <tr>
      <td colspan="2" align="center">
	  <input type="hidden" id="newJob" name="newJob">
      <input type="hidden" id="batchNo" name="batchNo">
	  <input type="hidden" id="computername" name="computername">
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Next 下一步" onClick="SaveData();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close 关闭" onClick="javascript:window.close();"> 
          
	  </td>
    </tr>
  </form>
</table>
</body>
<script>
<%if Request.Cookies("computer_name") = "" then %>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
<%else%>
	document.form1.computername.value="<%=Request.Cookies("computer_name")%>"; 
<%end if%>
</script>
</html>