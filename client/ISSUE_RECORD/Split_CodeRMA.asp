<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.min.css" />
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script type="text/javascript" src="../Scripts/jquery.blockUI.js"></script>
<script type="text/javascript" src="../Scripts/jquery.form.js"></script>
<script type="text/javascript" src="../Scripts/common.js"></script>
<script type="text/javascript">
	$(function(){
		$(document).ajaxStart($.blockUI).ajaxStop($.unblockUI);		
		form1.txt_inputboxid.focus();		
	})
   
   
   
   function inputboxid(objValue){
	if(event.keyCode==13 && objValue!=""){
	
	
	var r=$("#form1").formSerialize()+"&asynid=3";				
	var data = window.showModalDialog("Split_CodeRMA1.asp?"+r,window,"dialogHeight:100px;dialogWidth:100px;alwaysLowered:yes");
	var result=$.parseJSON(data);
		if(result.error==undefined){						
		document.getElementById("txt_inputcode").focus();
	
		
		
		}
		else{
		alert(result.error.replace("|","\n"));	
		
	
	}  
	}
   }

   
	function inputCode(objValue){
		if(event.keyCode==13 && objValue!=""){
			if($("#txt_inputboxid").val()==""){
				alert("Box ID  cannot be blank.\n箱号不能为空。");
				
				return false;
			}
			
			
			 if($("#hidBoxIdList").val().indexOf(objValue)>=0){
				alert("Input duplicate Code id.\n输入了重复的二维码。");
				return false;
			}

						
			var r=$("#form1").formSerialize()+"&asynid=1";				
			var data = window.showModalDialog("Split_CodeRMA1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			//var data='{"error":"Part number of This box id(FG911130708002) is different with packing plan.|\u8BE5\u7BB1\u53F7(FG911130708002)\u7684\u6599\u53F7\u4E0E\u5305\u88C5\u8BA1\u5212\u4E0D\u4E00\u81F4."}';
			var result=$.parseJSON(data);
			if(result.error==undefined){						
				document.getElementById("txt_inputcode").value="";
				var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.qty);						
				$("#txt_unitqty").val(totalQty);
				$("#txt_boxqty").val(parseInt($("#txt_boxqty").val())+parseInt(result.box_qty));
				$("#hidBoxIdList").val($("#hidBoxIdList").val()+","+objValue);	
				$("#hidNewUnitQty").val(result.new_qty);
				
				$("#txt_boxid5").val($("#txt_boxid4").val())
				$("#txt_boxid4").val($("#txt_boxid3").val())
				$("#txt_boxid3").val($("#txt_boxid2").val())						
				$("#txt_boxid2").val($("#txt_boxid1").val());
				$("#txt_boxid1").val(objValue);
			}
			else{
				alert(result.error.replace("|","\n"));		
								
			}
			
		
			
			document.getElementById("txt_inputcode").select();
		}
	}
	
	function doPrint(){		
		var totalQty=parseInt($("#txt_unitqty").val());
		
		if($("#txt_opcode").val()==""){
			alert("Op code cannot be blank.\n工号不能为空。");
			return false;
		}
		
		if($("#txt_boxID").val()==""){
			alert("box_id cannot be blank.\n箱号不能为空。");
			return false;
		}
		if($("#hidBoxIdList").val()==""){
			alert("Code cannot be blank.\n没有可以解包的二维码。");
			document.getElementById("txt_inputcode").focus();
			return false;
			
		}
		
		
		var r=$("#form1").formSerialize()+"&asynid=2";		
		var data = window.showModalDialog("Split_CodeRMA1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
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
			url: "WarehouseRec1.asp",
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
		});			*/		
	}
	
</script>
</head>

<body bgcolor="#339966" onload= "javascript:document.all.txt_inputboxid.focus();">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" align="center" class="incent">扫描二维码</td>
      </tr>
       <tr align="left" >
        <td>工号<span class="red"> *</span></td>
        <td><input name="txt_opcode" type="text" id="txt_opcode" onkeydown="if(event.keyCode==13){event.keyCode=9;}" size="23"  /></td>
	  </tr>	 
      <tr align="left" >
        <td width="68">箱号<span class="red"> *</span></td>
        <td width="211"><input name="txt_inputboxid" type="text" id="txt_inputboxid"  value=""size="23" onKeyDown="inputboxid(this.value)" /></td>
	  </tr>
       
	 
	
	  <tr >
        <td align="left">二维码<span class="red"> *</span></td>
        <td><input name="txt_inputcode" type="text" id="txt_inputcode" onKeyDown="inputCode(this.value)" size="23"/></td>
	  </tr>
	  <tr >
        <td align="left">二维码数</td>
        <td><input name="txt_boxqty" type="text" id="txt_boxqty" style='background-color:#666666' value="0" size="17"  readonly='true'/></td>
	  </tr>
	  
	  <tr align="center">
		<td colspan="2"> 已扫二维码</td>
      </tr>
	</thead>
	<tbody align="center">
		<tr><td colspan='2'><input name='txt_boxid1' type='text' id='txt_boxid1' style='background-color:#666666' size="23"  readonly='true' /></td></tr>
		<tr><td colspan='2'><input name='txt_boxid2' type='text' id='txt_boxid2' style='background-color:#666666' size="23"  readonly='true' /></td></tr>
        <tr><td colspan='2'><input name='txt_boxid3' type='text' id='txt_boxid3' style='background-color:#666666' size="23"  readonly='true' /></td></tr>
        <tr><td colspan='2'><input name='txt_boxid4' type='text' id='txt_boxid4' style='background-color:#666666' size="23"  readonly='true' /></td></tr>
        <tr><td colspan='2'><input name='txt_boxid5' type='text' id='txt_boxid5' style='background-color:#666666' size="23"  readonly='true' /></td></tr>
    </tbody>		
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input type="button" id="btn_print" value="解包" onClick="doPrint()"/>  			
			&nbsp;
			<input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="关闭">			
		</td>		
      </tr>
    </tfoot>
</table>
<input type="hidden" id="hidBoxIdList" name="hidBoxIdList" value="">
<input type="hidden" id="hidNewUnitQty" name="hidNewUnitQty" value="0">
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>