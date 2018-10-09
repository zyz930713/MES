<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
			html+="<tr><td colspan='2'><input type='text' name='txt_jobno' style='background-color:#666666'  readonly='true' onKeyDown='checkJobNo(this);' /></td><td colspan='2'><input type='text' name='txt_packqty' style='background-color:#666666' readonly='true' /></td></tr>"
		}
		$(html).appendTo($("#table1 tbody"));
	}
	
	function FromCheck(){
		if($("#txt_opcode").val().length==0){
			alert("Please input the op code.\n请输入工号.");
			$("#txt_opcode").focus();
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
		
		$.post("PackingOperation2.asp",r,function(data,textStatus){
			//alert(data);
			var result=$.parseJSON(data);
			if(result.error==undefined){
				alert(result.message.replace("|","\n"));
				//location.reload(); 
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
		
		var r=$("#form1").formSerialize();	
		r+="&asynid=2&txt_pack_type="+form1.txt_pack_type.value;			
		$.ajax({
			type: "post",
			url: "PackingOperation2.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);
				//$("#txt_boxid").val(result.boxid);
				if(result.error==undefined){
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
		if(event.keyCode==13){
			document.getElementById("unitTestInfo").src="PVSTestData.asp?txt_2d_code="+objValue;	
		
			var r=$("#form1").formSerialize();	
			r+="&asynid=3&txt_pack_type="+form1.txt_pack_type.value;	
			$.ajax({
				type: "post",
				url: "PackingOperation1.asp",
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
						document.getElementById("txt_2DCode").value="";
						
						var tempjobno,temppackqty;
						var aryCount=$("#table1 tbody tr").length;
						$("#table1 tbody tr").each(function(i,n){
							tempjobno=$.trim($(n).find("input[name='txt_jobno']").val());
							temppackqty=$.trim($(n).find("input[name='txt_packqty']").val());
							if(temppackqty==""){
								temppackqty=1;
							}else{
								temppackqty=parseInt(temppackqty)+1;
							}
							
							if(tempjobno==result.job_number){
								$(n).find("input[name='txt_packqty']").val(temppackqty);
								return false;
							}else if(tempjobno.length==0){
								$(n).find("input[name='txt_jobno']").val(result.job_number)
								$(n).find("input[name='txt_packqty']").val(1);
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
packType=request.Form("rad_pack_type")
if packType="" then
	response.Redirect("SelectPackType.asp")
elseif packType="NPI" then
response.Redirect "PackingOperationNPI.asp?rad_pack_type="&packType
else
packTypeDesc="Final Good 良品"
if packType="SCRAP" then
response.Redirect "PackingOperationNPI.asp?rad_pack_type="&packType
	packTypeDesc="Scrap 报废品"
end if
if  packType="RWK" then
packType="SCRAP"
response.Redirect "PackingOperationNPINew.asp?rad_pack_type="&packType
end if
if  packType="RWKFG" then
packType="FG"
response.Redirect "PackingOperationRWKFG.asp?rad_pack_type="&packType
end if
if  PackType="CheckABC" then
packType="FG"
response.Redirect "PackingOperationCheckABC.asp?rad_pack_type="&packType
end if

%>
<form id="form1" method="post">
<input type="hidden" id="txt_pack_type" value="<%=packType%>" />
  <table id="table1" width="800px" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Packing 包装
			&nbsp;&nbsp;&nbsp;<a style="color:#6495ED;" href="SelectPackType.asp"><%=packTypeDesc%></a>
		</td>
      </tr>
      <tr >
        <td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" size="22" name="txt_opcode" /></td>
		<td align="right">Customer 客户</td>
        <td><input type="text" id="txt_customer" name="txt_customer" /></td>		        
      </tr>
      <tr >
        <td align="right">2D Code 二维码</td>
		<td ><input type="text" name="txt_2DCode" size="22" onKeyDown="input2DCode(this.value)" /></td>
        <td align="right">Remarks 备注</td>
        <td><input type="text" id="txt_remarks" name="txt_remarks" /></td>
      </tr>
	  <tr  >
	  	<td align="right">Box ID 箱号</td>
        <td><input type="text" id="txt_boxid" name="txt_boxid"  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Part Number 料号</td>
        <td><input type="text" id="txt_part" name="txt_part"  style="background-color:#666666" readonly="true" /></td>
	  </tr>
	  <tr >
	  	<td align="right">Box Size 满箱数量</td>
        <td><input type="text" id="txt_boxsize" name="txt_boxsize"  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Pack Line 包装线别</td>
        <td><input type="text" id="txt_pack_line" name="txt_pack_line"  style="background-color:#666666" readonly="true" /></td>
	  </tr>
	  <tr >
	  	<td align="right">Shift 班别</td>
        <td><input type="text" id="txt_shift" name="txt_shift"  size="22" style="background-color:#666666" readonly="true" /></td>
		<td align="right">Supplier 供应商</td>
        <td><input type="text" id="txt_supplier" name="txt_supplier"  style="background-color:#666666" readonly="true" /></td>
	  </tr>
	  <tr><td colspan="4">&nbsp;</td></tr>
      <tr align="center">
        <td colspan="2">Job Number 工单号</td>
        <td colspan="2">Pack Qty 包装数量</td>
      </tr>	  
    </thead>
    <tbody align="center">
    </tbody>
    <tfoot>	
      <tr>
        <td colspan="4" align="center" >
		<%if packType="FG" then%>
			<input type="checkbox" value="Y" id="chkPartialPack" name="chkPartialPack">Partial Packing 部分包装<br><br>
		<%end if%>	
			<input type="button" id="btn_print" value="Print 打印" onClick="Print()"/>  
			&nbsp;
			<input name="btnReset" type="reset" id="btnReset" value="Reset 重置"> 
			&nbsp;
			<input name="Close" type="button" id="Close" onClick="window.close()" value="Close 关闭">        	
		</td>		
      </tr>
    </tfoot>
	
  </table>

  <center>
  <iframe src="PVSTestData.asp" width="95%" height="320" scrolling="auto" frameborder="0" id="unitTestInfo"></iframe>
  </center>
  <input type="hidden" id="computername" name="computername" />
</form>
<%end if%>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>