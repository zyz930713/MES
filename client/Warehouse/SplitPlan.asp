<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Split_Plan"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/Monitor.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.min.css" />
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script type="text/javascript" src="../Scripts/jquery.blockUI.js"></script>
<script type="text/javascript" src="../Scripts/jquery.form.js"></script>
<script type="text/javascript" src="../Scripts/common.js"></script>
<script type="text/javascript">
	$(function(){
		$(document).ajaxStart($.blockUI).ajaxStop($.unblockUI);		
		form1.txt_opcode.focus();		
	})

	function inputPLAN_ID(objValue){
		if(event.keyCode==13 && objValue!=""){
			if($("#txt_PLAN_ID").val()==""){
				alert("PLAN_ID Id cannot be blank.\n计划号不能为空。");
				$("#txt_PLAN_ID").focus();
				return false;
			}
			if($("#hidPLAN_IDList").val().indexOf(objValue)>=0){
				alert("Input duplicate box id.\n输入了重复的箱号。");
				return false;
			}
		    
			var r=$("#form1").formSerialize()+"&asynid=4";				
			var data = window.showModalDialog("SplitPlan1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			var result=$.parseJSON(data);
			
			if(result.error==undefined){						
						//document.getElementById("txt_PLAN_ID").value="";
						var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.qty);						
						$("#txt_unitqty").val(totalQty);
						$("#txt_boxqty").val(parseInt($("#txt_boxqty").val())+parseInt(result.box_qty));
							document.getElementById("txt_PLAN_ID").disabled=true;
						
					}
					else{
						alert(result.error.replace("|","\n"));	
						document.getElementById("txt_PLAN_ID").select();
					}
					
			/*$.ajax({
				type: "post",
				url: "Stacking1.asp",
				async:false,
				data:r,
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){						
						document.getElementById("txt_inputboxid").value="";
						var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.qty);						
						$("#txt_unitqty").val(totalQty);
						$("#txt_boxqty").val(parseInt($("#txt_boxqty").val())+parseInt(result.box_qty));
						$("#hidBoxIdList").val($("#hidBoxIdList").val()+","+objValue);	
												
						$("#txt_boxid2").val($("#txt_boxid1").val());
						$("#txt_boxid1").val(objValue);
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});*/
			
			
		}
	}
	
	function doSplit(){		
		if($("#txt_opcode").val()==""){
			alert("Op Code cannot be blank.\n工号不能为空。");
			$("#txt_opcode").focus();
			return false;
		}
		
	if($("#get_opcode").val()==""){
			alert("Get Code cannot be blank.\n提取工号不能为空。");
			$("#get_opcode").focus();
			return false;
		}
		
		if($("#txt_PLAN_ID").val()==""){
			alert("Plan ID cannot be blank.\n计划号不能为空。");
			$("#get_opcode").focus();
			return false;
		}
		document.getElementById("txt_PLAN_ID").disabled=false;
		var r=$("#form1").formSerialize()+"&asynid=5";			
		var data = window.showModalDialog("SplitPlan1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		var result=$.parseJSON(data);
		var result=$.parseJSON(data);
				if(result.error==undefined){
					alert(result.message.replace("|","\n"));
					location.reload(); 
				}
				else{
					alert(result.error.replace("|","\n"));						
				}		
		
		/*$.ajax({
			type: "post",
			url: "Stacking1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);
				if(result.error==undefined){
					alert(result.message.replace("|","\n"));
					location.reload(); 
				}
				else{
					alert(result.error.replace("|","\n"));						
				}			
			}
		});		*/			
	}
	function doSplitOutWH(){		
		if($("#txt_opcode").val()==""){
			alert("Op Code cannot be blank.\n工号不能为空。");
			$("#txt_opcode").focus();
			return false;
		}
		if($("#get_opcode").val()==""){
			alert("Get Code cannot be blank.\n提取工号不能为空。");
			$("#get_opcode").focus();
			return false;
		}
		
	if($("#txt_PLAN_ID").val()==""){
			alert("Plan ID cannot be blank.\n计划号不能为空。");
			$("#get_opcode").focus();
			return false;
		}
		document.getElementById("txt_PLAN_ID").disabled=false;
		var r=$("#form1").formSerialize()+"&asynid=6";			
		var data = window.showModalDialog("SplitPlan1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		var result=$.parseJSON(data);
		var result=$.parseJSON(data);
				if(result.error==undefined){
					alert(result.message.replace("|","\n"));
					location.reload(); 
				}
				else{
					alert(result.error.replace("|","\n"));						
				}		
		
		/*$.ajax({
			type: "post",
			url: "Stacking1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);
				if(result.error==undefined){
					alert(result.message.replace("|","\n"));
					location.reload(); 
				}
				else{
					alert(result.error.replace("|","\n"));						
				}			
			}
		});		*/			
	}
</script>
</head>

<body bgcolor="#339966">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Split Plan 解除计划</td>
      </tr>
      <tr >
        <td align="right">工号<span class="red"> *</span></td>
        <td><input name="txt_opcode" type="text" id="txt_opcode" size="17"   value="<%=session("UserCode")%>" readonly="readonly"  style="background-color:#666666"/></td>
	  </tr>	  
      
	  <tr >
        <td align="right">计划号<span class="red"> *</span></td>
        <td><input type="text" id="txt_PLAN_ID" name="txt_PLAN_ID" onkeydown="inputPLAN_ID(this.value)" /></td>
	  </tr>
	 <tr >
        <td align="right">提取工号<span class="red"> *</span></td>
        <td><input type="text" id="get_opcode" name="get_opcode" onkeydown="if(event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr>
	  <tr >
        <td align="right"> 产品数</td>
        <td><input type="text" id="txt_unitqty" name="txt_unitqty" value="0" style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	  <tr align="center">
		<td colspan="2">&nbsp;</td>
      </tr>
	</thead>
	<tbody align="center">
		<tr><td colspan='2'>&nbsp;</td></tr>
		<tr><td colspan='2'>&nbsp;</td></tr>
    </tbody>		
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input name="btnSplit" type="button" id="btnSplit" onClick="doSplitOutWH()" value="解除计划拍号出库">
            &nbsp;		
            <input name="btnSplit" type="button" id="btnSplit" onClick="doSplit()" value="解除计划拍号">		
			&nbsp;
			<input name="Close" type="button" id="Close"  onClick="window.close()" value="Close 关闭">
		</td>		
      </tr>
    </tfoot>
</table>
<input type="hidden" id="hidPLAN_IDList" name="hidPLAN_IDList" value="">
</form>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>