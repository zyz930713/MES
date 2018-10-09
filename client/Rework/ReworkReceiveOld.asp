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
		AddRows();
		form1.txt_opcode.focus();		
	})

	function AddRows(){
		html="";
		for(var i=0;i<2;i++){
			html+="<tr><td ><input type='text' name='txt_boxid' style='background-color:#666666'  readonly='true' /></td><td ><input type='text' name='txt_qty' style='background-color:#666666' readonly='true' /></td></tr>"
		}
		$(html).appendTo($("#table1 tbody"));
	}
	function inputBoxId(objValue){
		if(event.keyCode==13){
			var jobNo=$("#txt_job").val();
			if(jobNo==""){
				alert("Job Number cannot be blank.\n工单号不能为空。");
				return false;
			}
			else if(jobNo.indexOf("RWK")!=0){
				alert("This job number is not rework job number.\n此工单不是重工工单。");
				return false;
			}
		
		
			var r=$("#form1").formSerialize();	
			r+="&asynid=1";	
			$.ajax({
				type: "post",
				url: "ReworkReceiveold1.asp",
				async:false,
				data:r,
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){
						document.getElementById("txt_RewNumber").value="";
						var totalQty=parseInt($("#hidTotalQty").val())+parseInt(result.qty);						
						$("#hidTotalQty").val(totalQty);
						$("#hidBoxIdList").val($("#hidBoxIdList").val()+","+objValue);
						var tempBoxId;
						var aryCount=$("#table1 tbody tr").length;
						$("#table1 tbody tr").each(function(i,n){
							tempBoxId=$.trim($(n).find("input[name='txt_boxid']").val());							
							if(tempBoxId.length==0 || tempBoxId==objValue){
								$(n).find("input[name='txt_boxid']").val(objValue)
								$(n).find("input[name='txt_qty']").val(result.qty);
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
			document.getElementById("txt_RewNumber").select();
		}
	}
	
	function doPrint(){		
		var totalQty=parseInt($("#hidTotalQty").val());
		if(totalQty>0){
			var jobNo=$("#txt_job").val();
			if($("#txt_opcode").val()==""){
				alert("Op Code cannot be blank.\n工号不能为空。");
				return false;
			}
			if($("#sltStartStation").val()==""){
				alert("Start Station cannot be blank.\n开始站点不能为空。");
				return false;
			}
			if(jobNo==""){
				alert("Job Number cannot be blank.\n工单号不能为空。");
				return false;
			}else if(jobNo.indexOf("RWK")!=0){
				alert("This job number is not rework job number.\n此工单不是重工工单。");
				return false;
			}
			
			var r=$("#form1").formSerialize()+"&asynid=2";	
			$.ajax({
				type: "post",
				url: "ReworkReceiveold1.asp",
				async:false,
				data:r,
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){
						//initial page
						$("#btnReset").click();
						$("table tbody").html("");
						AddRows();
						
						$("#txt_subjob").val(result.subjob);						
					}
					else{
						alert(result.error.replace("|","\n"));						
					}			
				}
			});	
		}
		if($("#txt_subjob").val() != "" ){
			window.showModalDialog('PrintReworkTicket.asp?sub_job='+$("#txt_subjob").val(),'','dialogHeight:600px;dialogWidth:800px;status:no');
			$("#txt_subjob").val("");
		}
	}
</script>
</head>

<body bgcolor="#339966">
<form id="form1" method="post">
<table id="table1" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="4" class="t-t-DarkBlue" align="center">Rework Receive 重工接收</td>
      </tr>
      <tr >
        <td align="right">Op Code 工号<span class="red"> *</span></td>
        <td><input type="text" id="txt_opcode" name="txt_opcode" /></td>
	  </tr>
	  <tr>
		<td align="right">Start Station 开始站点<span class="red"> *</span></td>
        <td><select id="sltStartStation" name="sltStartStation">
				<option value=""></option>
				<%sql="select nid,station_name,station_chinese_name from station_new where nid in ('SA00000988') order by nid"
				rs.open sql,conn,1,3
				while not rs.eof 
					response.Write("<option value='"+rs("nid")+"'>"+rs("station_name")+"("+rs("station_chinese_name")+")</option>")
					rs.movenext
				wend
				rs.close
				%>
				
			</select>
		</td>
      </tr>
	  <tr >
        <td align="right">Job Number 工单号<span class="red"> *</span></td>
        <td><input type="text" id="txt_job" name="txt_job" /></td>
	  </tr>
	  <tr >
        <td align="right">返工数量<span class="red"> *</span></td>
        <td><input type="text" id="txt_RewNumber" name="txt_RewNumber" onKeyDown="inputBoxId(this.value)"/></td>
	  </tr>
	  <tr >
        <td align="right">Sub Job Number 子工单号</td>
        <td><input type="text" id="txt_subjob" name="txt_subjob"/></td>
	  </tr>
	  <tr align="center">
		<td >返工数量</td>
        <td >Qty 数量</td>
      </tr>
	</thead>
	<tbody align="center">
    </tbody>
	<tfoot>	
      <tr>
        <td colspan="4" align="center" ><br>		
			<input type="button" id="btn_print" value="Print 打印" onClick="doPrint()"/>  
			&nbsp;
			<input name="btnReset" type="reset" id="btnReset" value="Reset 重置"> 
			&nbsp;
			<input name="Close" type="button" id="Close" onClick="window.close()" value="Close 关闭">        	
		</td>		
      </tr>
    </tfoot>
</table>
<input type="hidden" id="hidTotalQty" name="hidTotalQty" value="0">
<input type="hidden" id="hidBoxIdList" name="hidBoxIdList" value="">
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->