<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKEB.css" rel="stylesheet" type="text/css">
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
			if($("#slt_plan").val()==""){
				alert("Plan cannot be blank.\n计划不能为空。");
				return false;
			}
			if(objValue.indexOf("FG")!=0){
				alert("This box id is not final good.\n此箱号不是良品箱号。");
				return false;
			}else if($("#hidBoxIdList").val().indexOf(objValue)>=0){
				alert("Input duplicate box id.\n输入了重复的箱号。");
				return false;
			}

			document.getElementById("slt_plan").disabled=false;
			document.getElementById("txt_pallet").disabled=false;
			var planPart=$("#slt_plan").find("option:selected").text();		
			var r=$("#form1").formSerialize()+"&asynid=1&plan_part="+planPart;				
			$.ajax({
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
						$("#hidNewUnitQty").val(result.new_qty);
												
						$("#txt_boxid2").val($("#txt_boxid1").val());
						$("#txt_boxid1").val(objValue);
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});
			document.getElementById("slt_plan").disabled=true;
			document.getElementById("txt_pallet").disabled=true;	
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
			alert("There is no data to be printed.\n没有数据可以打印。");
			return false;		
		}
		document.getElementById("slt_plan").disabled=false;
		document.getElementById("txt_pallet").disabled=false;
		var r=$("#form1").formSerialize()+"&asynid=2";			
		$.ajax({
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
		});					
	}
	function doCancel(){
		document.getElementById("slt_plan").disabled=false;
		document.getElementById("txt_pallet").disabled=false;
		var r=$("#form1").formSerialize()+"&asynid=3";			
		$.ajax({
			type: "post",
			url: "Stacking1.asp",
			async:false,
			data:r,
			success:function(data,textStatus){
				var result=$.parseJSON(data);				
			}
		});
		window.close();
	}
</script>
</head>

<body bgcolor="#339966">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Stacking 堆栈</td>
      </tr>
      <tr >
        <td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" name="txt_opcode" /></td>
	  </tr>
	  <tr>
		<td align="right">Plan 计划<span class="red"> *</span></td>
        <td><select id="slt_plan" name="slt_plan">
				<option value=""></option>
				<%sql="select plan_id,plan_date,part_number,to_char(delivery_time,'mmddhh24mi') delivery_time from packing_plan where status in ('Initial','Wait') order by delivery_time,priority"
				rs.open sql,conn,1,3
				while not rs.eof 
					response.Write("<option value='"+rs("plan_id")+"'>"+rs("part_number")+"("+rs("delivery_time")+")</option>")
					rs.movenext
				wend
				rs.close
				%>
				
			</select>
		</td>
      </tr>
	  <tr >
        <td align="right">Pallet Id 拍号</td>
        <td><input type="text" id="txt_pallet" name="txt_pallet" onkeydown="if(event.keyCode==13){event.keyCode=9;}" /></td>
	  </tr>
	  <tr >
        <td align="right">Box ID 箱号<span class="red"> *</span></td>
        <td><input type="text" id="txt_inputboxid" name="txt_inputboxid" onKeyDown="inputBoxId(this.value)"/></td>
	  </tr>
	  <tr >
        <td align="right">Box Qty 箱数</td>
        <td><input type="text" id="txt_boxqty" name="txt_boxqty" value="0" style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	  <tr >
        <td align="right">Unit Qty 产品数</td>
        <td><input type="text" id="txt_unitqty" name="txt_unitqty" value="0" style='background-color:#666666'  readonly='true'/></td>
	  </tr>
	  <tr align="center">
		<td colspan="2">Stack Box ID 已堆箱号</td>
      </tr>
	</thead>
	<tbody align="center">
		<tr><td colspan='2'><input type='text' id='txt_boxid1' name='txt_boxid1' style='background-color:#666666'  readonly='true' /></td></tr>
		<tr><td colspan='2'><input type='text' id='txt_boxid2' name='txt_boxid2' style='background-color:#666666'  readonly='true' /></td></tr>
    </tbody>		
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input type="button" id="btn_print" value="Print 打印" onClick="doPrint()"/>  			
			&nbsp;
			<input name="Close" type="button" id="Close" onClick="doCancel()" value="Close 关闭">			
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