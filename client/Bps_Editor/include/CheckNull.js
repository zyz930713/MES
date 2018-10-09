function beforeSubmit(form){
	if(form.Hdate.value==''){
		alert('日期不能为空！');
		form.Hdate.focus();
		return false;
	}
	return true;
}

function beforeSubmit(form){
	if(form.d122.value==''){
		alert('日期不能为空！');
		form.d122.focus();
	return false;
	}

	if(form.d11.value!=''){
		if(form.d12.value==''){
			alert('停止时间不能为空！');
			form.d12.focus();
		return false;
		}
		if(form.DPos1.value==''){
			alert('工位名称不能为空！');
			form.DPos1.focus();
		return false;
		}
		if(form.DPro1.value==''){
			alert('故障名称不能为空！');
			form.DPro1.focus();
		return false;
		}

		if(form.DRes1.value==''){
			alert('故障原因不能为空！');
			form.DRes1.focus();
		return false;
		}
		if(form.DSol1.value==''){
			alert('解决方法不能为空！');
			form.DSol1.focus();
		return false;
		}
		if(form.DTim1.value==''){
			alert('停机时间不能为空！');
			form.DTim1.focus();
		return false;
		}
	}
return true;	
}

function trim(s){
	var count = s.length;
	var st = 0;       // start   
	var end = count-1; // end   
	if (s == "") return s;
	while (st < count){
		if (s.charAt(st) == " ")
			st ++;
		else
			break;
	}
	while (end > st){
		if (s.charAt(end) == " ")
			end --;
		else
			break;
	}
	return s.substring(st,end + 1);
}