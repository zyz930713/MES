<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
code=request.Form("txtOperatorCode")
SNNO=request("SNNO")
Sendarea=request("Sendarea")


ComputerName=request("computername")

district=GetdistrictName(ComputerName)

SNNOA = GetdistrictSNNO(PTCSTATE,BOXNO,GETCODE,ACCEPTCODE,TXTCOMMENTS)


	function GetdistrictName(ComputerName)
		set rsdistrictName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(ComputerName)+"'"
	 
		rsdistrictName.open SQL,conn,1,3
	 	if(rsdistrictName.recordcount=0) then
			GetdistrictName=""
			exit function
		end if 
		    
		if (rsdistrictName(1)="") then
			GetdistrictName=""
			exit function
		end if 
		
	
		GetdistrictName=rsdistrictName(1)
	
	end function
	

	function GetdistrictSNNO(PTCSTATE,BOXNO,GETCODE,ACCEPTCODE,TXTCOMMENTS)
		set rsdistrictSNNO=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM PTC_SN WHERE SNNO='"+trim(SNNO)+"'"
 
		rsdistrictSNNO.open SQL,conn,1,3
	 	 if rsdistrictSNNO.eof and rsdistrictSNNO.bof then
			GetdistrictSNNO=""
			
		 else 	
		
		GetdistrictSNNO=rsdistrictSNNO("SNNO")
		PTCSTATE=rsdistrictSNNO("PTCSTATE")
		BOXNO=rsdistrictSNNO("BOXNO")
		GETCODE=rsdistrictSNNO("GETCODE")
		ACCEPTCODE=rsdistrictSNNO("ACCEPTCODE")
		TXTCOMMENTS=rsdistrictSNNO("TXTCOMMENTS")
		end if 
	   
	end function
'response.End()
if SNNOA<>"" then
     if trim(PTCSTATE)<>"等待确认借出" then
word="<span align='center' style='color:red;'>此单号已经使用过，请换一个新的单号</span>"
	response.Redirect("SendProd.asp?word="&word)
	end if
end if



if isnull(district) or instr(district,Sendarea)=0 then
	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师，此机器不能借料。</span>"
	response.Redirect("SendProd.asp?word="&word)
end if 
	
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("SendProd.asp?word="&word)
else		
OperatorCode=request.Form("txtOperatorCode")	
end if 
%>
<script language="javascript" type="text/javascript" src="/Components/My97DatePicker/WdatePicker.js"></script>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="My97DatePicker/WdatePicker.js"></script>
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
	if (document.form1.area.value=="")
   {
		alert("请选择区域! ");
		document.form1.area.focus();
		return false
   }else if (document.form1.BoxNO.value=="")
  {
    alert("箱号不能为空！");
	document.form1.BoxNO.focus();
	return false;
   }else   if (document.form1.getCode.value=="")
  {
    alert("借走人不能为空！");
	document.form1.getCode.focus();
	return false;
  }else   if (document.form1.acceptCode.value=="")
  {
    alert("接收人不能为空！");
	document.form1.acceptCode.focus();
	return false;
  }else if (document.form1.txtComments.value=="")
  {
    alert("备注不能为空！");
	document.form1.txtComments.focus();
	return false;
  }else if (document.form1.SelectType.value=="归还")
	if (document.form1.ExpectDate.value=="")
	  {
		alert("时间不能为空！");
		document.form1.ExpectDate.focus();
		return false;
	  }
  if (confirm("确定要提交吗？"))
   { 
			document.form1.action="SendProd_Save.asp";
			document.form1.submit();
	}
}
	 
	 function checkNumber() {    //判断输入字符的keyCode，数字在48到57之间，超出部分返回false   
	  if ((event.keyCode >= 48) && (event.keyCode <= 57))
	   {      
	     event.returnValue = true;    
		 } 
		 else 
		 {        event.returnValue = false;    }
		 
		 }
 
	 
	 
	 
	 
	   function Changearea()
        {
           
		   
		  
		    var data = { "LAB": ["不归还","归还"] };
			var data1 = { "CS": ["不归还","归还"] };
			var data2 = { "研发": ["不归还","归还"] };
		    var prov = document.getElementById("area");
	        var SelectType = document.getElementById("SelectType");
	        var provName = prov.value;
          
		  
          
		  
		  
         	if (provName=="LAB")
			
			{    for (var i = SelectType.childNodes.length - 1; i >= 0; i--) //从后往前删就行了
           			 {
                	var option = SelectType.childNodes[i];
                   SelectType.removeChild(option);
           			 }
			
				 
          
				 var SelectTypes = data[provName];
           		 for (var i = 0; i < SelectTypes.length; i++)
            		{
                  var option = document.createElement("option");
                  option.value = SelectTypes[i];
                  option.innerText = SelectTypes[i];
                  SelectType.appendChild(option);
                    }

			}
			else if (provName=="CS")
			
			{    for (var i = SelectType.childNodes.length - 1; i >= 0; i--) //从后往前删就行了
           			 {
                	var option = SelectType.childNodes[i];
                   SelectType.removeChild(option);
           			 }
			
				 
          
				 var SelectTypes = data1[provName];
           		 for (var i = 0; i < SelectTypes.length; i++)
            		{
                  var option = document.createElement("option");
                  option.value = SelectTypes[i];
                  option.innerText = SelectTypes[i];
                  SelectType.appendChild(option);
                    }

			}
			else if (provName=="研发")
			
			{    for (var i = SelectType.childNodes.length - 1; i >= 0; i--) //从后往前删就行了
           			 {
                	var option = SelectType.childNodes[i];
                   SelectType.removeChild(option);
           			 }
			
				 
          
				 var SelectTypes = data2[provName];
           		 for (var i = 0; i < SelectTypes.length; i++)
            		{
                  var option = document.createElement("option");
                  option.value = SelectTypes[i];
                  option.innerText = SelectTypes[i];
                  SelectType.appendChild(option);
                    }

			}
			else
			{ 
				 for (var i = SelectType.childNodes.length - 1; i >= 0; i--) //从后往前删就行了
           			 {
                   var option = SelectType.childNodes[i];
                   SelectType.removeChild(option);
           			 }
			
			 var option = document.createElement("option");
			      option.value = "归还";
                  option.innerText ="归还";
				 SelectType.appendChild(option);
				
				}
	 
           
        }
	 
	
	 
</script>
 

</head>
<body bgcolor="#339966">
<table width="95%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1" action="SendProd_Save.asp"  >
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">(物料借出操作)<%Response.Write(year(Now())&"-"&month(Now())&"-"&day(Now())&" "&time())%></td>
    </tr>
    
	<tr>
      <td colspan="2"><BR></td>
    </tr>
	<tr>
		<td>
		<table width="100%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr>
			    <td>区域</td>
				<td>单号</td>
				<td>借出方</td>
				<td>箱号</td>
				<td>是否归还</td>
				<td>归还日期</td>
				<td>借走人工号</td>
				<td>接收人工号</td>
				<td>Comments(备注)</td>
			</tr>
				<tr>
				<td width="8%"><select name="area" onChange="Changearea()" id="area">
                   <option></option>
                   <option value="LAB">实验室</option>
				   <option value="FA">FA</option>
				   <option value="研发">研发</option>
				   <option value="工程">工程</option>
				   <option value="IQC">IQC</option>
                   <option value="生产线">生产线</option>
                   <option value="X-Ray">X-Ray</option>
                   <option value="2D打码等级">2D打码等级</option>
				   <option value="CS">CS</option>
                </select></td>
				
				<td width="11%"><input name="SNNOD" type="text" id="SNNOD"  value="<%=request("SNNO")%>"  disabled="disabled" size="15"><input type="hidden" id="SNNO" name="SNNO"  value="<%=request("SNNO")%>" ></td>				
				<td width="11%"><input name="SendareaD" type="text" id="SendareaD"  value="<%=request("Sendarea")%>"  disabled="disabled" size="15">
				<input type="hidden" id="Sendarea" name="Sendarea"  value="<%=request("Sendarea")%>" ></td>
				<td width="10%"><input name="BoxNO" type="text" id="BoxNO" value="<%=BoxNO%>" size="15" ></td>
				<td width="15%"><select name="SelectType" id="SelectType"></select></td>
				<td width="15%"><input name="ExpectDate" id="ExpectDate" type="text" value="<%=request("ExpectDate")%>" size=20" onClick="WdatePicker({skin:'whyGreen',dateFmt:'yyyy-MM-dd HH:mm:ss'})"/></td>
				<td><input name="getCode" type="text" id="getCode" value="<%=getCode%>" size="10" ></td>
				<td><input name="acceptCode" type="text" id="acceptCode" value="<%=acceptCode%>" size="10" ></td>
				<td width="20%"><input type="text" id="txtComments" name="txtComments" value="<%=txtComments%>"  maxlength="45"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center"><label>
		  <input name="btnSubmit" type="button" id="btnSubmit"  value="提交扫描二维码" onClick="CheckForm();">
		</label>
		  &nbsp;
			<input name="btnClose2" type="button"  id="btnClose2"  value="返回上一页面" onClick="javascript:history.go(-1);">
			&nbsp;
			<input name="btnClose" type="button"  id="btnClose"  value="Close(关闭)" onClick="window.close();">
	  </Td>
	</tr>
		<input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
		<input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
		<input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->