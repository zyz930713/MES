<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
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

	function inputBoxId(objValue){
		if(event.keyCode==13){
			
			//if(objValue.indexOf("FG")!=0){
				//alert("This box id is not final good.\n此箱号不是良品箱号。");
				//return false;
			//}else
			 if($("#hidBoxIdList").val().indexOf(objValue)>=0){
				alert("Input duplicate box id.\n输入了重复的箱号。");
				return false;
			}

						
			var r=$("#form1").formSerialize()+"&asynid=1";				
			var data = window.showModalDialog("WarehouseRec1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			//var data='{"error":"Part number of This box id(FG911130708002) is different with packing plan.|\u8BE5\u7BB1\u53F7(FG911130708002)\u7684\u6599\u53F7\u4E0E\u5305\u88C5\u8BA1\u5212\u4E0D\u4E00\u81F4."}';
			var result=$.parseJSON(data);
			if(result.error==undefined){						
				document.getElementById("txt_inputboxid").value="";
				var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.qty);						
				$("#txt_unitqty").val(totalQty);
				$("#txt_boxqty").val(parseInt($("#txt_boxqty").val())+parseInt(result.box_qty));
				$("#hidBoxIdList").val($("#hidBoxIdList").val()+","+objValue);	
				$("#hidNewUnitQty").val(result.new_qty);
										
				$("#txt_boxid2").val($("#txt_boxid1").val());
				$("#txt_boxid1").val(objValue);
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
						document.getElementById("txt_inputboxid").value="";
						var totalQty=parseInt($("#txt_unitqty").val())+parseInt(result.qty);						
						$("#txt_unitqty").val(totalQty);
						$("#txt_boxqty").val(parseInt($("#txt_boxqty").val())+parseInt(result.box_qty));
						$("#hidBoxIdList").val($("#hidBoxIdList").val()+","+objValue);	
						$("#hidNewUnitQty").val(result.new_qty);
												
						$("#txt_boxid2").val($("#txt_boxid1").val());
						$("#txt_boxid1").val(objValue);
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});*/
			
			document.getElementById("txt_inputboxid").select();
		}
	}
	
	function doPrint(){		
		var totalQty=parseInt($("#txt_unitqty").val());
		if($("#txt_opcode").val()==""){
			alert("Op Code cannot be blank.\n工号不能为空。");
			return false;
		}
		if(totalQty==0 &&$("#txt_pallet").val()==""){
			alert("There is no data to be WH.\n没有数据可以入库。");
			return false;		
		}
		
		var r=$("#form1").formSerialize()+"&asynid=2";		
		var data = window.showModalDialog("WarehouseRec1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
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
	function doCancel(){
		
		var r=$("#form1").formSerialize()+"&asynid=3";	
		var data = window.showModalDialog("WarehouseRec1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		var result=$.parseJSON(data);		
		/*$.ajax({
			type: "post",
			url: "WarehouseRec1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);				
			}
		});*/
		window.close();
	}
</script>
</head>

<body bgcolor="#339966">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" align="center" class="incent">WH REC 确认入库</td>
      </tr>
      <tr >
        <td align="right">工号<span class="red"> *</span></td>
        <td><input name="txt_opcode" type="text" id="txt_opcode" size="17" onKeyDown="if(event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr>
	 
	
	  <tr >
        <td align="right">箱号<span class="red"> *</span></td>
        <td><input name="txt_inputboxid" type="text" id="txt_inputboxid" onKeyDown="inputBoxId(this.value)" size="17"/></td>
	  </tr>
	  <tr >
        <td align="right">箱数</td>
        <td><input name="txt_boxqty" type="text" id="txt_boxqty" style='background-color:#666666' value="0" size="17"  readonly='true'/></td>
	  </tr>
	  <tr >
        <td align="right">产品数</td>
        <td><input name="txt_unitqty" type="text" id="txt_unitqty" style='background-color:#666666' value="0" size="17"  readonly='true'/></td>
	  </tr>
	  <tr align="center">
		<td colspan="2">Box ID 已扫箱号</td>
      </tr>
	</thead>
	<tbody align="center">
		<tr><td colspan='2'><input type='text' id='txt_boxid1' name='txt_boxid1' style='background-color:#666666'  readonly='true' /></td></tr>
		<tr><td colspan='2'><input type='text' id='txt_boxid2' name='txt_boxid2' style='background-color:#666666'  readonly='true' /></td></tr>
    </tbody>		
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input type="button" id="btn_print" value="入库" onClick="doPrint()"/>  			
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