<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.min.css" />
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script type="text/javascript" src="../Scripts/jquery.blockUI.js"></script>
<script type="text/javascript" src="../Scripts/jquery.form.js"></script>
<script type="text/javascript" src="../Scripts/common.js"></script>
<script type="text/javascript">
	function checkJobNo(obj){
		if(event.keyCode==13){
			var index=obj.value.lastIndexOf("-");
			if(index>0){
				obj.value=obj.value.substring(0,index);
			}
		}
	}
	$(function(){
		$(document).ajaxStart($.blockUI).ajaxStop($.unblockUI);		
		
		var wsh=new ActiveXObject("WScript.Network"); 
		$("#computername").val(wsh.ComputerName);
		//alert(wsh.ComputerName);
		form1.txt_opcode.focus();		
	})
	

	
	function FromCheck(){
		if($("#txt_opcode").val().length==0){
			alert("Please input the op code.\n请输入工号.");
			$("#txt_opcode").focus();
			return false;
		}
		if($("#txt_part").val().length==0){
			alert("Please input the 12NC.\n请输入成品料号.");
			$("#txt_part").focus();
			return false;
		}
		if($("#txt_Job_Number").val().length==0){
			alert("Please input the Job NunberC.\n请输入工单号.");
			$("#txt_Job_Number").focus();
			return false;
		}
		if($("#txt_custid").val().length==0){
			alert("请输入客户标签号");
			$("#txt_custid").focus();
			return false;
		}
		if($("#txt_boxsize").val().length==0){
			alert("Please input the Qty.\n请输入包装数量.");
			$("#txt_boxsize").focus();
			return false;
		}
		else 
		{  boxSize=$("#txt_boxsize").val()
			if(isNaN(boxSize))
			
			{
				alert("Input Quantity is not a number!\n输入的数量不是数字！");
				$("#txt_boxsize").select();
				return false;					
				}
			
			}
		
		//检查keyin数据
		var tempjobno,temppackqty,check;
		$("#table1 tbody tr").each(function(i,n){
			tempjobno=$.trim($(n).find("input[name='txt_jobno']").val());
			temppackqty=$.trim($(n).find("input[name='txt_packqty']").val());
			
		
			
			if((tempjobno.length>0&&temppackqty.length==0)||(tempjobno.length==0&&temppackqty.length>0)){
				alert("Please fill in the blanks at line "+(i+1)+"\n请将第"+(i+1)+"行数据填写完整.");
				check=false;
				return false;
			}
			if(temppackqty.length>0){
				temppackqty=parseInt(temppackqty);
				if(isNaN(temppackqty)){
					alert("The value in line "+(i+1)+" is not a effective number.\n第"+(i+1)+"行的包装数量不是有效数字.");
					check=false;
					return false;
				}			
			}
		})
		if(check==false){
			return check;
		}
		else{
			return true;			
		}		
	}
	
	function FormatChinese(parmString){
		var parmArray = parmString.split("&");
		var parmStringNew="";
		$.each(parmArray,function(index,data){
			var li_pos = data.indexOf("="); 
			if(li_pos >0){
				var name = data.substring(0,li_pos);
				var value = escape(decodeURIComponent(data.substr(li_pos+1)));
				var parm = name+"="+value;
				parmStringNew = parmStringNew=="" ? parm : parmStringNew + '&' + parm;
			}
		});
		return parmStringNew;
	}
	
	function Save(){
		if(!FromCheck()){
			return false;
		}
		var r=$("#form1").formSerialize();	
		r+="&asynid=1&txt_pack_type="+form1.txt_pack_type.value;
		r=FormatChinese(r);
		
		$.post("PackingOperationException1.asp",r,function(data,textStatus){
			//alert(data);
			var result=$.parseJSON(data);
			if(result.error==undefined){
				alert(result.message.replace("|","\n"));
				window.location="PackingOperationException.asp"; 
				$("#btnReset").click();	
				$("table tbody").html("");
				AddRows();
				document.getElementById("unitTestInfo").src="PVSTestData.asp";	
			}
			else{
				alert(result.error.replace("|","\n"));
			} 
		});
		document.getElementById("txt_opcode").blur();
		document.getElementById("txt_opcode").focus();
	}
	
	function Print(){
		if(!FromCheck()){
			return false;
		}
		
		var r=$("#form1").formSerialize();	
		r+="&asynid=2&txt_pack_type="+form1.txt_pack_type.value;			
		$.ajax({
			type: "post",
			url: "PackingOperationOtherSite1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);
			//	$("#txt_boxid").val(result.boxid);
				if(result.error==undefined){
				alert(result.message.replace("|","\n"));
			    location.reload(); 
								
				}
				else{
					alert(result.error.replace("|","\n"));
				}			
			}			
		});
	}
	
	

	
</script>
<body bgcolor="#339966">
<%
packType=request("txt_pack_type")
txt_customer=request("txt_customer")

if packType="" then
	response.Redirect("SelectPackType.asp")
elseif packType="NPI" then
packTypeDesc="NPI Final Good 良品"
packType="FG"
elseif packType="SCRAP" then
	packTypeDesc="Scrap 报废品"
end if 
txt_boxid=trim(request("txt_boxid"))


if txt_boxid<>"" then    
			
 			sql="select job_number, PART_NUMBER,PACKED_QTY,CUSTOMER, PACKED_LINE,CUSTOMER_LABEL_ID ,BIND_BOX_ID from job_pack_detail where  box_id='"&txt_boxid&"'"
	
			rs.open sql,conn,1,3
			
			if rs.eof then
			word="<span align='center' style='color:red;'>"&txt_boxid&"此箱号不存在，请确认箱号！</span>"
			response.Redirect("PackingOperationSelect.asp?word="&word)
			else
			
 			exists2DInfo=""
 			packedQty=0
		 ''   if  not rs.eof then
			while not rs.eof
			job_number=rs("job_number")
			if left(job_number,2)<>"KB"  then
			   if left(job_number,3)<>"EKB" then
			   word="<span align='center' style='color:red;'>此箱号中包函RWK工单不能和KB工单混包</span>"
	           response.Redirect("PackingOperationSelect.asp?word="&word)
			   end if
			end if
			job_number=rs("job_number")
			qty=rs("PACKED_QTY")
			part_number_tag=rs("PART_NUMBER")
			line_name=rs("PACKED_LINE")
			CUSTOMER_LABEL_ID=rs("CUSTOMER_LABEL_ID")
			CUSTOMER=rs("CUSTOMER")
			BIND_BOX_ID=rs("BIND_BOX_ID")
			packedQty=packedQty+cint(qty)
			rs.movenext
			wend
			rs.close
			
		  end if

    
end if

%>






<form id="form1" method="post">
<input type="hidden" id="txt_pack_type" value="<%=packType%>" />

  <table id="table1" width="800px" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Packing 包装
			&nbsp;&nbsp;&nbsp;<a style="color:#6495ED;" href="SelectPackType.asp"><%=packTypeDesc%></a>		</td>
      </tr>
      <tr >
        <td align="right">Customer 客户</td>
        <td><input name="txt_customer" type="text" id="txt_customer" style="background-color:#666666" value="<%=txt_customer%>" size="22" readonly="readonly"/></td>
		<td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" size="22" name="txt_opcode"  onkeydown="if(this.value&&event.keyCode==13){event.keyCode=9;}"  value=""  /></td>		        
      </tr>
      <tr >
        <td align="right">Box ID 箱号</td>
		<td ><input type="text" id="txt_boxid" name="txt_boxid"   value="<%=request("txt_boxid")%>"size="22" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}" ></td>
        <td align="right">Part Number 料号</td>
        <td><input name="txt_part" type="text" id="txt_part" onkeydown="if(this.value&&event.keyCode==13){event.keyCode=9;}"    value="<%=part_number_tag%>" size="22"  /></td>
      </tr>
	  <tr  >
	  	<td align="right">Job Numbera工单号</td>
        <td><input type="text" name="txt_Job_Number" id="txt_Job_Number" size="22" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}"  value="<%=job_number%>"/></td>
		<td align="right">Box Size 满箱数量</td>
        <td><input type="text" id="txt_boxsize" name="txt_boxsize"  size="22" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}" value="<%=qty%>"/></td>
	  </tr>
	  <tr >
	  	<td align="right">客户标签号</td>
        <td><input name="txt_custid" type="text" id="txt_custid"   value="<%=CUSTOMER_LABEL_ID%>" size="22"  onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}"/></td>
		<td align="right">Remarks 备注</td>
        <td><input name="txt_remarks" type="text" id="txt_remarks" size="22" onkeydown="if(this.value&amp;&amp;event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr>
	  <tr >
	  	<td align="right">&nbsp;</td>
        <td>&nbsp;</td>
		<td align="right">&nbsp;</td>
        <td>&nbsp;</td>
	  </tr>
	   <tr >
	  	<td align="right">&nbsp;</td>
        <td>&nbsp;</td>
		<td align="right">&nbsp;</td>
        <td>&nbsp;</td>
	   </tr>
	  
	  <tr><td colspan="4">&nbsp;</td></tr>
      <tr align="center">
        <td colspan="4">&nbsp;</td>
      </tr>	  
    </thead>
    <tbody align="center">
	
    </tbody>
    <tfoot>	
    <tr>
        <td colspan="4" align="center" >&nbsp;</td>		
      </tr>
      <tr>
        <td colspan="4" align="center" >
		
			<input type="button" id="btn_print" value="扫入BPS系统" onClick="Print()"/>  
			&nbsp;
			<input name="btnReset" type="reset" id="btnReset" value="Reset 重置"> 
			&nbsp;
			<input name="Close" type="button" id="Close" onClick="window.close()" value="Close 关闭">		</td>		
      </tr>
    </tfoot>
  </table>

  <input type="hidden" id="hidJobList" name="hidJobList" value="">
  <input type="hidden" id="computername" name="computername" />
</form>
<%'end if%>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>