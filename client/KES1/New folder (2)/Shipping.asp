<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
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
	function inputPalletId(objValue){
		if(event.keyCode==13){
			if($("#txt_delivery").val().length==0){
				alert("Delivery No cannot be blank.\n提货单号不能为空.");
				$("#txt_delivery").focus();
				return false;
			}
			if($("#hidPalletIdList").val().indexOf(objValue)>=0){
				alert("Input duplicate pallet id.\n输入了重复的拍号。");
				return false;
			}
					
			document.getElementById("txt_delivery").disabled=false;	
			var r=$("#form1").formSerialize()+"&asynid=2"				
			var data = window.showModalDialog("Shipping1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			var result=$.parseJSON(data);
			if(result.error==undefined){						
						document.getElementById("txt_pallet").value="";
						var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.unit_qty);						
						$("#txt_unitqty").val(totalQty);
						$("#txt_pallet_qty").val(parseInt($("#txt_pallet_qty").val())+parseInt(result.pallet_qty));
						$("#hidPalletIdList").val($("#hidPalletIdList").val()+","+objValue);	
						
						$("#txt_palletid2").val($("#txt_palletid1").val());
						$("#txt_palletid1").val(objValue);
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
			
			
			
			
			
			/*
			var r=$("#form1").formSerialize();	
			r+="&asynid=2";	
			$.ajax({
				type: "post",
				url: "Shipping1.asp",
				async:false,
				data:r,
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){						
						document.getElementById("txt_pallet").value="";
						var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.unit_qty);						
						$("#txt_unitqty").val(totalQty);
						$("#txt_pallet_qty").val(parseInt($("#txt_pallet_qty").val())+parseInt(result.pallet_qty));
						$("#hidPalletIdList").val($("#hidPalletIdList").val()+","+objValue);	
						
						$("#txt_palletid2").val($("#txt_palletid1").val());
						$("#txt_palletid1").val(objValue);
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});	*/
			document.getElementById("txt_pallet").select();
			document.getElementById("txt_delivery").disabled=true;	
		}
	}
	function SubmitForm(){
		if($("#txt_opcode").val().length==0){
			alert("Please input the op code.\n请输入工号.");
			$("#txt_opcode").focus();
			return false;
		}
		if($("#hidPalletIdList").val().length==0){
			alert("Please input pallet id.\n请输入拍号.");
			$("#txt_pallet").focus();
			return false;
		}
		document.getElementById("txt_delivery").disabled=false;
		
		var r=$("#form1").formSerialize()+"&asynid=3"				
	    var data = window.showModalDialog("Shipping1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		
		var result=$.parseJSON(data);
			if(result.error==undefined){
				alert(result.message.replace("|","\n"));
				location.reload();
			}
			else{
				alert(result.error.replace("|","\n"));
			} 
		
		
	}

</script>
<body bgcolor="#339966">
<form id="form1" method="post">
  <table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Shipping 出货</td>
      </tr>
      <tr >
        <td align="right"> 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode"  name="txt_opcode" onkeydown="if(event.keyCode==13){event.keyCode=9;}"  /></td>			        
      </tr>
      <tr >
	  	<td align="right"> 提货单号</td>
        <td><input type="text" id="txt_delivery" name="txt_delivery" onkeydown="if(event.keyCode==13){event.keyCode=9;}" /></td>        
      </tr>
	  <tr >
		<td align="right">拍号</td>
        <td><input type="text" id="txt_pallet" name="txt_pallet" onKeyDown="inputPalletId(this.value)" /></td>		
	  </tr>	  
	  <tr >
        <td align="right">拍数</td>
        <td><input type="text" id="txt_pallet_qty" name="txt_pallet_qty" value="0" style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	  <tr >
        <td align="right"> 产品数</td>
        <td><input type="text" id="txt_unitqty" name="txt_unitqty" value="0" style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	  <tr align="center">
		<td colspan="2"> 已扫描拍号</td>
      </tr>       
    </thead>
    <tbody align="center">
		<tr><td colspan='2'><input type='text' id='txt_palletid1' name='txt_palletid1' style='background-color:#666666'  readonly='true' /></td></tr>
		<tr><td colspan='2'><input type='text' id='txt_palletid2' name='txt_palletid2' style='background-color:#666666'  readonly='true' /></td></tr>	
    </tbody>
    <tfoot>	
      <tr>
        <td colspan="4" align="center" >		
			<input type="button" id="btn_print" value="OK 确定" onClick="SubmitForm()"/>  
			&nbsp;			
			<input name="Close" type="button" id="Close" onClick="window.close()" value="Close 关闭">        	
		</td>		
      </tr>
    </tfoot>	
  </table>
  <input type="hidden" id="hidPalletIdList" name="hidPalletIdList" value="">
</form>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>
