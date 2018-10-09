<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Job Inventory</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link rel="Stylesheet" type="text/css" href="../Scripts/jquery-ui-1.9.2/css/ui-lightness/jquery-ui-1.9.2.custom.css" />
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script type="text/javascript" src="../Scripts/jquery.blockUI.js"></script>
<script type="text/javascript" src="../Scripts/jquery.form.js"></script>
<script type="text/javascript" src="../Scripts/common.js"></script>
<script type="text/javascript">
	$(function(){
		$(document).ajaxStart($.blockUI).ajaxStop($.unblockUI);
		var $jobno=$("#txt_jobno");
		$jobno.focus();
		$("#divWindow").dialog({
			autoOpen: false,
			height: 600,
			width: 900,
			modal: true,
			draggable: false,
			resizable: false,
			title:"Packing Detail List 包装明细" 
		})
		$("#btn_search").bind("click",function(){
			$.get("JobInventory1.asp",{"asynid":"1","txt_jobno":$.trim($("#txt_jobno").val()),"txt_partno":$.trim($("#txt_partno").val())},function(data,textStatus){
				var $table=$("#tbsearch tbody");
				$table.empty();
				if(data.length>2){
					var result=$.parseJSON(data);
					var html="";						
					$.each(result,function(i,n){
						html+="<tr ><td><a href=\"#\" onclick=\"ShowPacked('"+n["JOB_NUMBER"]+"')\">"+n["JOB_NUMBER"]+"</a></td>";
						html+="<td>"+n["PART_NUMBER"]+"</td><td>"+n["GOOD_QTY"]+"</td><td>"+n["SCRAP_QTY"]+"</td><td>"+n["PACKED_QTY"]+"</td></tr>";
					})						
					if(html.length>0){							
						$(html).appendTo($table);
						$("#tbsearch").show();
					}
				}
				else{
					alert("No such job number 无此工单号");
					$jobno.focus();
				}
			})			
		})
	})
	
	function ShowPacked(obj){
		//alert(obj);
		$.get("JobInventory1.asp",{"asynid":"2","txt_jobno":obj},function(data,textStatus){
			var $table=$("#tbWindow tbody"); 
			$table.empty();
			if(data.length>2){
				var result=$.parseJSON(data);
				var html="";
				$.each(result,function(i,n){
					if(n["CUSTOMER"]==null){
						n["CUSTOMER"]="&nbsp;";
					}
					if(n["REMARKS"]==null){
						n["REMARKS"]="&nbsp;";
					}
					html+="<tr><td>"+n["BOX_ID"]+"</td><td>"+n["JOB_NUMBER"]+"</td><td>"+n["PART_NUMBER"]+"</td><td>"+n["PACKED_QTY"]+"</td><td>"+n["CUSTOMER"]+"</td><td>"+n["REMARKS"]+"</td></tr>";
				})
				$(html).appendTo($table);				
			}
			$("#divWindow").dialog("open");			
		})
	}
</script>
</head>
<body bgcolor="#339966">
<form id="form1">
  <table width="600"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td colspan="5" class="t-t-DarkBlue" align="center">Job Inventory 工单库存</td>
    </tr>
    <tr align="center">
      <td>Job Number 工单号</td>
      <td><input type="text" id="txt_jobno" name="txt_jobno" /></td>
      <td>Part Number 型号</td>
      <td><input type="text" id="txt_partno" name="txt_partno" /></td>	  
      <td><input type="button" id="btn_search" value="Query 查询" /></td>
    </tr>
  </table>
<br>
  <table id="tbsearch" width="600"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" style="display:none">
    <thead>
      <tr align="center" class='t-t-Book'>
        <td>Job Number 工单号</td>
		<td>Part Numbe 型号</td>
        <td>Good Qty 良品数量</td>
		<td>Scrap Qty 报废数量</td>
        <td>Packed Qty 包装数量</td>
      </tr>
    </thead>
    <tbody align="center">
    </tbody>
  </table>
  <div id="divWindow">
    <table id="tbWindow" width="800"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
      <thead>
        <tr align="center" class='t-t-Book'>
          <td>Box ID 箱号</td>
          <td>Job Number 工单号</td>
          <td>Part Number 物料号</td>
          <td>Pack QTY 包装数量</td>
          <td>Customer 客户</td>
		  <td>Remarks 备注</td>
        </tr>
      </thead>
      <tbody align="center">
      </tbody>
    </table>
  </div>
</form>
</body>
</html>
