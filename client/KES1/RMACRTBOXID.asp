<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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

function inputPart(objValue){
if(event.keyCode==13 && objValue!=""){

		
var r=$("#form1").formSerialize()+"&asynid=3";				
var data = window.showModalDialog("RMACRTBOXID1.asp?"+r,window,"dialogHeight:100px;dialogWidth:100px;alwaysLowered:yes");
var result=$.parseJSON(data);
		if(result.error==undefined){						
		
			$("#txt_Remarks").val(result.PARTDESCRIPTION);
		document.getElementById("txt_QTY").select();
		
		}
		else{
			alert(result.error.replace("|","\n"));	
			
		document.getElementById("txt_Part").select();
			
			}
}
}

	function doCRT(){
		
			if($("#txt_opcode").val()==""){
			alert("Op Code cannot be blank.\n工号不能为空。");
			$("#txt_opcode").focus();
			return false;
		}
		if($("#txt_RMANO").val()==""){
			alert("RMA Id cannot be blank.\n客退号不能为空。");
			$("#txt_RMANO").focus();
			return false;		
		}
		if($("#txt_Part").val()==""){
			alert("Part No cannot be blank.\n料号不能为空。");
			$("#txt_Part").focus();
			return false;		
		}
		
		if($("#txt_QTY").val()==""){
			alert("QTY cannot be blank.\n数量不能为空。");
			$("#txt_QTY").focus();
			return false;		
		}
				
			var r=$("#form1").formSerialize()+"&asynid=1";				
			var data = window.showModalDialog("RMACRTBOXID1.asp?"+r,window,"dialogHeight:100px;dialogWidth:100px;alwaysLowered:yes");
			var result=$.parseJSON(data);
			if(result.error==undefined){						
												
						$("#txt_boxId").val(result.boxId);
						
					}
					else{
						alert(result.error.replace("|","\n"));	
					}
					
			
		}
	
	
	
	
	
	
	function doPrint(){		
		if($("#txt_opcode").val()==""){
			alert("Op Code cannot be blank.\n工号不能为空。");
			$("#txt_opcode").focus();
			return false;
		}
		
		if($("#txt_RMANO").val()==""){
			alert("RMA Id cannot be blank.\n客退号不能为空。");
			$("#txt_RMANO").focus();
			return false;		
		}
		if($("#txt_Part").val()==""){
			alert("Part No cannot be blank.\n料号不能为空。");
			$("#txt_Part").focus();
			return false;		
		}
		
		if($("#txt_QTY").val()==""){
			alert("QTY cannot be blank.\n数量不能为空。");
			$("#txt_QTY").focus();
			return false;		
		}
		
		
		if($("#txt_boxId").val()==""){
			alert("BOX ID cannot be blank.\n没有生成箱号不能打印。");
			
			return false;		
		}
		
		var r=$("#form1").formSerialize()+"&asynid=2";			
		var data = window.showModalDialog("RMACRTBOXID1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
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
			url: "RMA1.asp",
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
        <td colspan="4" class="t-t-DarkBlue" align="center">RMA生成临时箱号</td>
      </tr>
      <tr >
        <td align="right">工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" name="txt_opcode" onkeydown="if(event.keyCode==13){event.keyCode=9;}"  /></td>
	  </tr>	 
      <tr >
        <td align="right">客退号<span class="red"> *</span></td>
        <td><input type="text" id="txt_RMANO" name="txt_RMANO" onkeydown="if(event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr> 
	  <tr >
        <td align="right">料号<span class="red"> *</span></td>
        <td><input type="text" id="txt_Part" name="txt_Part" onkeydown="inputPart(this.value)" /></td>
	  </tr>
       <tr >
        <td align="right">描述<span class="red"> *</span></td>
        <td><input type="text" id="txt_Remarks" name="txt_Remarks"   readonly='true'/></td>
	  </tr>
	  <tr >
        <td align="right">数量<span class="red"> *</span></td>
        <td><input type="text" id="txt_QTY" name="txt_QTY" onkeydown="if(event.keyCode==13){event.keyCode=9;}"/></td>
	  </tr>
	  <tr >
        <td align="right">箱号<span class="red"> *</span></td>
        <td><input type="text" id="txt_boxId" name="txt_boxId"  style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	 
	  <tr align="center">
		<td colspan="2">&nbsp;</td>
      </tr>
	</thead>
			
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input name="btnCRTBOX" type="button" id="btnCRTBOX" onClick="doCRT()" value="成生箱号">			
			&nbsp;
			<input name="btnPrint" type="button" id="btnPrint"  onClick="doPrint()" value="打印">
		</td>		
      </tr>
    </tfoot>
</table>
<input type="hidden" id="hidBoxIdList" name="hidBoxIdList" value="">
<input type="hidden" id="hidBoxIdQTY" name="hidBoxIdQTY" value="">
</form>
</body>
</html>
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>