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
		AddRows();
		var wsh=new ActiveXObject("WScript.Network"); 
		$("#computername").val(wsh.ComputerName);
		//alert(wsh.ComputerName);
		form1.txt_2DCode.focus();		
	})
	
	function AddRows(){
		html="";
		for(var i=0;i<2;i++){
			html+="<tr><td colspan='2'><input type='text' name='txt_jobno' style='background-color:#666666'  readonly='true' onKeyDown='checkJobNo(this);' /></td><td colspan='2'><input type='text' name='txt_packqty' style='background-color:#666666'  /></td></tr>"
		}
		$(html).appendTo($("#table1 tbody"));
	}
	
	function FromCheck(){
		if($("#txt_opcode").val().length==0){
			alert("Please input the op code.\n请输入工号.");
			$("#txt_opcode").focus();
			return false;
		}
	if($("#txt_remarks").val().length==0){
			alert("备注不能为空，请输入日期.");
			$("#txt_remarks").focus();
			return false;
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
		
		$.post("PackingOperationNbass1.asp",r,function(data,textStatus){
			//alert(data);
			var result=$.parseJSON(data);
			if(result.error==undefined){
				alert(result.message.replace("|","\n"));
				window.location="PackingOperationNbass.asp"; 
				$("#btnReset").click();	
				$("table tbody").html("");
				AddRows();
				document.getElementById("unitTestInfo").src="PVSTestData.asp";	
			}
			else{
				alert(result.error.replace("|","\n"));
			} 
		});
		document.getElementById("txt_2DCode").blur();
		document.getElementById("txt_2DCode").focus();
	}
	
	function Print(){
		if(!FromCheck()){
			return false;
		}
		 
           if($("#txt_boxid").val()==""){
				alert("do not Print,Box ID is Null.\n箱号不能为空。");
				return false;
			}
		
		 if($("#txt_customer").val()=="HuaWei"){
				if($("#txt_custid").val()==""){
				alert("do not Print,Customer ID is Null.\n客户标签号不能为空请输入。");
				return false;
			}
			}
		
		var r=$("#form1").formSerialize();	
		r+="&asynid=2&txt_pack_type="+form1.txt_pack_type.value;			
		$.ajax({
			type: "post",
			url: "PackingOperationNbass1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);
				$("#txt_boxid").val(result.boxid);				
				if(result.error==undefined){
				$("#txt_custid").val(result.cust_id);
					var iscontinued=true;
					while(iscontinued){
						var pro=prompt("Please confirm the boxid is "+result.boxid+"\n请确认扫描打印的箱号是否为 "+result.boxid,"");
						if (pro!=null && pro!=""){
							if(pro==result.boxid){
								Save();
								//alert("Save");
								iscontinued=false;
							}
							else{
								if(!confirm("The boxid is not agreed with the system, check again?\n输入箱号与系统不一致，重新确认？")){
									iscontinued=false;
								}
							}
						}
					}				
				}
				else{
					alert(result.error.replace("|","\n"));
				}			
			}			
		});
	}
	
	function input2DCode(objValue){
	
	  if($("#txt_opcode").val().length==0){
			alert("Please input the op code.\n请输入工号.");
			$("#txt_opcode").focus();
			
			return false;
		}
	
	
		if(event.keyCode==13){
			var name= 	$.trim(objValue).toUpperCase();



			if($("#hidJobList").val().indexOf(objValue)>=0){
				alert("Input duplicate Job Number.\n输入了重复的工单。");
		document.getElementById("txt_2DCode").select();
		document.getElementById("txt_2DCode").focus();
				return false;
			}
         if (!$("#txt_boxsize").val()=="" &&  !$("#txt_count").val()==""){
           if($("#txt_boxsize").val()==$("#txt_count").val()){
				alert("Box is Full.\n包装箱已满。");
				return false;
			}
		 }
		//document.getElementById("txt_2DCode").value=name;
		
		
		
		//document.getElementById("unitTestInfo").src="PVSTestData.asp?txt_2d_code="+name;	
	
			var r=$("#form1").formSerialize();	
			r+="&asynid=3&txt_pack_type="+form1.txt_pack_type.value;	
			$.ajax({
				type: "post",
				url: "PackingOperationNbass1.asp",
				async:false,
				data:r,
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){
						$("#txt_part").val(result.part_number);
						$("#txt_boxid").val(result.boxid);
						$("#txt_pack_line").val(result.pack_line);
						$("#txt_boxsize").val(result.boxsize);
						$("#txt_shift").val(result.prod_shift);
						$("#txt_supplier").val(result.supplier);
						$("#txt_count").val(result.PCount);
						$("#txt_Job_Number").val(result.job_number)
						$("#txt_JOB_GOOD_QTY").val(result.JOB_GOOD_QTY)
						$("#txt_remarks").val(result.remarks)
						$("#txt_customer").val(result.customer)
					
						$("#hidJobList").val($("#hidJobList").val()+","+objValue);	
						document.getElementById("txt_2DCode").value="";
						
						var tempjobno,temppackqty;
						var aryCount=$("#table1 tbody tr").length;
						$("#table1 tbody tr").each(function(i,n){
							tempjobno=$.trim($(n).find("input[name='txt_jobno']").val());
							temppackqty=$.trim($(n).find("input[name='txt_packqty']").val());
							//if(temppackqty==""){
							//	temppackqty=1;
						//	}else{
						//		temppackqty=parseInt(temppackqty)+1;
						//	}
							
							if(tempjobno==result.job_number){
								$(n).find("input[name='txt_packqty']").val(result.packqty);
								return false;
							}else if(tempjobno.length==0){
								$(n).find("input[name='txt_jobno']").val(result.job_number)
								$(n).find("input[name='txt_packqty']").val(result.packqty);
								if(i==aryCount-1){
									AddRows();
								}
								return false;
							}
						})
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});	
			document.getElementById("txt_2DCode").select();	
		}
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

         
          
			
 			sql="select job_number, PART_NUMBER,job_number,PACKED_QTY,CUSTOMER,PROD_SHIFT, PACKED_LINE,CUSTOMER_LABEL_ID,REMARKS from job_pack_detail where  box_id='"&txt_boxid&"'"
	
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
			subJob=rs("job_number")
			qty=rs("PACKED_QTY")
			part_number_tag=rs("PART_NUMBER")
			line_name=rs("PACKED_LINE")
			PROD_SHIFT=rs("PROD_SHIFT")
			CUSTOMER_LABEL_ID=rs("CUSTOMER_LABEL_ID")
			txt_customer=rs("CUSTOMER")
			REMARKS=rs("REMARKS")
			exists2DInfo=exists2DInfo+"<tr><td colspan='2'><input type='text' name='txt_jobno' style='background-color:#666666'  readonly='true' value="&subJob&" /></td><td colspan='2'><input type='text' name='txt_packqty' style='background-color:#666666'  value="&qty&" /></td></tr>"
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
        <td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" size="22" name="txt_opcode" /></td>
		<td align="right">Customer 客户</td>
        <td><input type="text" id="txt_customer" name="txt_customer" value="<%=txt_customer%>" style="background-color:#666666" readonly="readonly"/></td>		        
      </tr>
      <tr >
        <td align="right">Job Numbera工单号</td>
		<td ><input type="text" name="txt_2DCode"  id="txt_2DCode" size="22" onKeyDown="input2DCode(this.value)" /></td>
        <td align="right">Remarks 备注</td>
        <td><input type="text" id="txt_remarks" name="txt_remarks" value="<%=remarks%>" /></td>
      </tr>
	  <tr  >
	  	<td align="right">Box ID 箱号</td>
        <td><input type="text" id="txt_boxid" name="txt_boxid"   value="<%=request("txt_boxid")%>"size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Part Number 料号</td>
        <td><input type="text" id="txt_part" name="txt_part"  style="background-color:#666666" readonly="true" value="<%=part_number_tag%>"></td>
	  </tr>
	  <tr >
	  	<td align="right">Box Size 满箱数量</td>
        <td><input type="text" id="txt_boxsize" name="txt_boxsize"  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Pack Line 包装线别</td>
        <td><input type="text" id="txt_pack_line" name="txt_pack_line"  style="background-color:#666666" readonly="true" value="<%=line_name%>"/></td>
	  </tr>
	  <tr >
	  	<td align="right">Shift 班别</td>
        <td><input type="text" id="txt_shift" name="txt_shift" value="<%=PROD_SHIFT%>"  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Supplier 供应商</td>
        <td><input type="text" id="txt_supplier" name="txt_supplier"  style="background-color:#666666" readonly="true" /></td>
	  </tr>
	   <tr >
	  	<td align="right">已包装数量</td>
        <td><input type="text" id="txt_count" name="txt_count"value="<%=packedQty%>"  size="22" style="background-color:#666666" /></td>
		<td align="right">客户标签号</td>
        <td>
        <%if txt_customer<>"HuaWei" then %>
        <input name="txt_custid" type="text" id="txt_custid" style="background-color:#666666"  value="<%=CUSTOMER_LABEL_ID%>" size="22" readonly="true"/>
        <%else%>
        <input name="txt_custid" type="text" id="txt_custid"   value="<%=CUSTOMER_LABEL_ID%>" size="22" />
        <%end if%>
        </td>
	   </tr>
	   <tr >
	  	<td align="right">正在包装的工单</td>
        <td><input type="text" id="txt_Job_Number" name="txt_Job_Number"value=""  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">可包装的良品数量</td>
        <td><input name="txt_JOB_GOOD_QTY" type="text" id="txt_JOB_GOOD_QTY" style="background-color:#666666"  value="" size="22" readonly="true"/></td>
	   </tr>
       <tr >
	  	<td align="right">检查人工号</td>
        <td><input type="text" id="txt_INSPECTIONPQC" name="txt_INSPECTIONPQC"value=""  size="22"  /></td>
		<td align="right">Lot NO</td>
        <td><input type="text" id="txt_LotNO" name="txt_LotNO"value=""  size="22"  /></td>
	   </tr>
	  <tr><td colspan="4">&nbsp;</td></tr>
      <tr align="center">
        <td colspan="2">Job Number 工单号</td>
        <td colspan="2">Pack Qty 包装数量</td>
      </tr>	  
    </thead>
    <tbody align="center">
	<%=exists2DInfo%>
    </tbody>
    <tfoot>	
    <tr>
        <td colspan="4" align="center" ><%if packType="FG" then%>
			<input type="checkbox" value="Y" id="chkPartialPack" name="chkPartialPack">
			箱号 
             <input type="checkbox" value="Y" id="chkCustidPack" name="chkCustidPack">
            只打印客户标签
            <input type="checkbox" value="Y" id="chkTrayPack" name="chkTrayPack">打印单个小标签
            <br><br>
		<%end if%>	&nbsp;
		
		</td>		
      </tr>
      <tr>
        <td colspan="4" align="center" >
		
			<input type="button" id="btn_print" value="Print 打印" onClick="Print()"/>  
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