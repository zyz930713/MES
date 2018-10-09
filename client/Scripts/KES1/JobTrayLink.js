// JavaScript Document
$(function(){
	//var regExp=/^\d+(\.\d+)?$/;
	//var test1="1232.";
	//alert(regExp.test(test1));
	$(document).ajaxStart($.blockUI).ajaxStop($.unblockUI);
	$("#txt_jobno").focus();
	$("#txt_jobno").bind("change",function(){
		var jobno=$.trim($("#txt_jobno").val());
		if(jobno.indexOf("-")==-1){			
			alert("�ӹ�����ʽ����ȷ\nThe format of sub job number is wrong");
			$("#txt_jobno").focus();
			return false;
		}
		if(jobno.length>0){
			$.ajax({
				async:false,
				type:"get",
				url:"JobTrayLink1.asp",
				data:{"asynid":"1","txt_jobno":jobno},
				success:function(data,textStatus){
					var result=$.parseJSON(data);
					if(result.error==undefined){
						$("#txt_qty").val(result.startqty);
						$("#txt_partno").val(result.partno);
						$("#txt_opcode").val(result.opcode);
						var saveddata=null;
						$.ajax({
							async:false,
							type:"get",
							url:"JobTrayLink1.asp",
							data:{"asynid":"4","txt_jobno":jobno},
							success:function(data,textStatus){
								saveddata=data;//�����Ѿ������������
							}
						})
						
						$.ajax({
							async:false,
							type:"get",
							url:"JobTrayLink1.asp",
							data:{"asynid":"2","partno":result.partno},
							success:function(data,textStatus){
								var r=$.parseJSON(data);
								if(r.length>0){								
									var html="<table width='90%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolorlight='#000099' bordercolordark='#FFFFFF'>";
									var saveddata_length=0;
									if(saveddata!=null&&saveddata.length>2){
										saveddata=$.parseJSON(saveddata);
										saveddata_length=saveddata.length;
									}
									else{
										saveddata=null;
									}									
									var index=0;									
									$.each(r,function(i,n){								
										html+="<tr><td colspan='7'>"+n["STATION_NAME"]+"("+n["STATION_CHINESE_NAME"]+")--"+n["TRAY_TYPE"]+"</td></tr>";	
										var times=Math.ceil(result.startqty/n["TRAY_SIZE"]);
										//alert(times);
										for(var j=0;j<times;j++){
											//�����ѱ�������ݣ�traysizeû�б仯��stationid��ͬ������ͬһseq���У�
											if(saveddata!=null&&index<saveddata_length&&saveddata[index]["STATION_ID"]==n["STATION_ID"]&&saveddata[index]["DISPLAY_SEQ"]==(j+1)){
												html+="<tr><td>"+(j+1)+"</td><td>Tray ID</td><td><input type='text' name='txt_trayid' value='"+saveddata[index]["TRAY_ID"]+"' /></td>";
												if(saveddata[index]["MATERIAL_LOT_NO"]!=null){
													html+="<td>Material Lot No</td><td><input type='text' name='txt_lotno' value='"+saveddata[index]["MATERIAL_LOT_NO"]+"' /></td>";
												}
												else{
													html+="<td>Material Lot No</td><td><input type='text' name='txt_lotno' /></td>";
												}
												//alert(saveddata[index]["MATERIAL_QTY"]);
												if(saveddata[index]["MATERIAL_QTY"]!=null){
													html+="<td>Material Qty</td><td><input type='text' name='txt_mqty' value='"+saveddata[index]["MATERIAL_QTY"]+"' />";
												}
												else{
													html+="<td>Material Qty</td><td><input type='text' name='txt_mqty' />";
												}
												html+="<input type='hidden' name='txt_displayseq' value='"+saveddata[index]["DISPLAY_SEQ"]+"' />";
												index++;
											}
											else{
												html+="<tr><td>"+(j+1)+"</td><td>Tray ID</td><td><input type='text' name='txt_trayid' /></td>";
												html+="<td>Material Lot No</td><td><input type='text' name='txt_lotno' /></td>";
												html+="<td>Material Qty</td><td><input type='text' name='txt_mqty' />";
												html+="<input type='hidden' name='txt_displayseq' value='"+(j+1)+"' />";
									
											}	
											html+="<input type='hidden' name='txt_stationid' value='"+n["STATION_ID"]+"' /></tr>";											
										}
									})
									html+="<tr><td colspan='7'><input type='button' value='Save ����' onclick='Save()' /></td></tr></table>";
									$("#div_station").empty();
									$(html).appendTo($("#div_station"));
									
								}
								else{
									alert("�Ϻ�"+result.partno+"û��ά������\nLot No "+result.partno+" is not exist");
								}							
							}
						})
					}
					else{
						$("#txt_jobno").val("");
						$("#txt_qty").val("");
						$("#txt_partno").val("");
						alert(result.error.replace("|","\n"));
					}					
				}
			})
		}			
	})
})

function Save(){
	if($("#txt_jobno").val().length==0){
		alert("�����빤����\nPlease input job number");
		$("#txt_jobno").focus();
		return false;
	}
	if($("#txt_opcode").val().length==0){
		alert("�����빤��\nPlease input op code");
		$("#txt_opcode").focus();
		return false;
	}
	//��ֵ���
	var wrongnumber=false;
	var $wrongnumber=null;
	var regExp=/^\d+(\.\d+)?$/;
	$.each($("input[name='txt_mqty']"),function(i,n){
		if($.trim($(n).val()).length>0){
			if(!regExp.test($(n).val())){				
				wrongnumber=true;
				$wrongnumber=$(n);
				return false;
			}
		}									
	})
	if(wrongnumber){
		$wrongnumber.focus();
		alert("��ֵ����\nData error");		
		return false;
	}
	var r=$("#form1").formSerialize();	
	r+="&asynid=3";	
	$.post("JobTrayLink1.asp",r,function(data,textStatus){
		//alert(data);
		var result=$.parseJSON(data);
		if(result.error==undefined){
			alert(result.message.replace("|","\n"));
			location.reload(); 
		}
		else{
			alert(result.error.replace("|","\n"));
		}
	})
}