<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",HuaWei_ADMINISTRATOR"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
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

	function inputBoxId(objValue){
		if(event.keyCode==13){
			
			//if(objValue.indexOf("FG")!=0){
				//alert("This box id is not final good.\n此箱号不是良品箱号。");
				//return false;
			//}else
				
			var r=$("#form1").formSerialize()+"&asynid=1";				
			var data = window.showModalDialog("BOX_Bind_SPECIAL1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
			//var data='{"error":"Part number of This box id(FG911130708002) is different with packing plan.|\u8BE5\u7BB1\u53F7(FG911130708002)\u7684\u6599\u53F7\u4E0E\u5305\u88C5\u8BA1\u5212\u4E0D\u4E00\u81F4."}';
			var result=$.parseJSON(data);
			if(result.error==undefined){						
				document.getElementById("txt_inputboxid").value="";
				
				var totalQty=parseInt($("#Sum_unitqty").val())+parseInt(result.qty);						
				$("#Sum_unitqty").val(totalQty);
				
				$("#txt_unitqty").val(parseInt(result.qty));
				$("#txt_box").val(result.boxID);
				$("#part_number").val(result.part_number);
			    form1.Bind_boxid.focus();	
			
				
			}
			else{
				
				alert(result.error.replace("|","\n"));						
			}
			
		
		}
	}
	
	function bind_box(Bqty){
		
	if(event.keyCode==13){		
	var totalQty=parseInt($("#txt_unitqty").val());
	var Bind_boxid_Qty=parseInt($("#Bind_boxid_Qty").val());
	
	//var Bind_boxid_Qty=objValue
	
		if($("#txt_box").val()==""){
			alert("Op Code cannot be blank.\n箱号不能为空。");
			return false;
		}
		if(totalQty==0 ){
			alert("数量不能空。");
			return false;		
		}
		if(totalQty!=Bind_boxid_Qty ){
			alert("数量不对！！！");
			return false;		
		}
		var r=$("#form1").formSerialize()+"&asynid=2";		
		var data = window.showModalDialog("BOX_Bind_SPECIAL1.asp?"+r,window,"dialogHeight:10px;dialogWidth:10px;alwaysLowered:yes");
		var result=$.parseJSON(data);
		if(result.error==undefined){
			alert(result.message.replace("|","\n"));
			location.reload(); 
		}
		else{
			alert(result.error.replace("|","\n"));						
		}	
		//$.ajax({
//			type: "post",
//			url: "BOX_Bind_SPECIAL1.asp",
//			async:false,
//			data:r,
//			success:function(data,textStatus){
//				var result=$.parseJSON(data);
//				if(result.error==undefined){
//					alert(result.message.replace("|","\n"));
//					location.reload(); 
//				}
//				else{
//					alert(result.error.replace("|","\n"));						
//				}			
//			}
		//});					
	}
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
        <td><input name="txt_opcode" type="text" id="txt_opcode"  value="<%=session("UserCode")%>" size="17" onKeyDown="if(event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr>
      <tr >
        <td align="right">箱号<span class="red"> *</span></td>
        <td><input name="txt_inputboxid" type="text" id="txt_inputboxid" onKeyDown="inputBoxId(this.value)" size="17"/></td>
	  </tr>
	 
      <tr >
        <td align="right">绑定箱号<span class="red"> *</span></td>
        <td><input name="Bind_boxid" type="text" id="Bind_boxid" onKeyDown="if(event.keyCode==13){event.keyCode=9;}" size="17"/></td>
	  </tr>
       <tr >
        <td align="right">绑定料号<span class="red"> *</span></td>
        <td><input name="Bind_Part" type="text" id="Bind_Part" onKeyDown="if(event.keyCode==13){event.keyCode=9;}" size="17"/></td>
	  </tr>
      <tr >
        <td align="right">绑定箱号数量<span class="red"> *</span></td>
        <td><input name="Bind_boxid_Qty" type="text" id="Bind_boxid_Qty" onKeyDown="bind_box(this.value)" size="17"/></td>
	  </tr>
	  <tr align="center">
		<td colspan="2">当前箱号和数量</td>
      </tr>
	  <tr align="center">
	    <td>箱号</td>
	    <td><input name="txt_box" type="text" id="txt_box" style='background-color:#666666'  size="17"  readonly='readonly'/></td>
      </tr>
	  <tr align="center">
	    <td>产品数</td>
	    <td><input name="txt_unitqty" type="text" id="txt_unitqty" style='background-color:#666666' value="0" size="17"  readonly='readonly'/></td>
      </tr>
	</thead>
	<tbody align="center">
    </tbody>		
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			  			
			&nbsp;
            <input type="hidden" id="part_number" name="part_number" value="">
			<input name="Close" type="button" id="Close" onClick="javascript:window.close()" value="关闭">			
		</td>		
      </tr>
    </tfoot>
</table>

</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<script language="JavaScript" src="/Functions/NoRefresh.js" type="text/javascript"></script>