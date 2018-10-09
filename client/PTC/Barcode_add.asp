
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">

</head>
<script type="text/javascript" src="JS/jquery-1.9.0.min.js"></script>
<script language="javascript" type="text/javascript">

var arrdd=[];

function selectCheck(testnum)
{
	var ttt=eval(testnum+1);
	
	
  	var evt=evt?evt:(window.event?window.event:null);//兼容IE和FF
	if (evt.keyCode==13){
		
		var name= 	$.trim($("#Dcode" + testnum).val()).toUpperCase();
       $("#Dcode" + testnum).val(name);
       var SNNO = $("div#SNNO").html();
       var QueryStringTest
       var SNNO = $("div#SNNO").html();
	   
	   
	         if (name=="SANEND")
			 {
	          }
			  else
			  {
	          if(name.length!=17) 
              {
	
               alert("对不起，2D码位数不对！");
	           return false;
              } 
              
	   
	        /* if (left(name ,3)!="DYD" || left(name ,3)!="MAP")
             {
	
            alert("对不起！北京工厂代码应为:DYD");
	         return false;
             } 
*/

  
       
	  
	  
     
  
  
     
  	  
     
	  
	  
  	   
	   }
	   
	   
	   
	   
	   
		
		
		
        var exist=$.inArray(name,arrdd);

	  if(exist>=0)
	  {
            alert('重复'+name);  
			 $("#Dcode" + testnum).select();
			 $("#Dcode" + testnum).css("background-color","red");
      }	
	  else
	  {
	  		arrdd.push(name);
			
		 QueryStringTest =  "'" + arrdd.join("','") + "'";
		
		

           if (testnum==10 || name=="SANEND")
			 { 
            ajax_post();
			}
	  
	         else
		  
		    {
		    
		  $("#Dcode" + testnum).css("background-color","#FFFFFF");
		  $("#Dcode" + ttt).select();
		  }
	  
	  
	  
	  }
	  
	  
	  
       
	
	
	function ajax_post()
	{
  $.post("Barcode_Save.asp",{Dcode:QueryStringTest,SNNO:SNNO},
  function(data){
    
  
   var reg=/,$/gi;                // 格式化最后一个逗号
   str=data.replace(reg,"");
  var a=str.split(",");           //以逗号为分隔符得到 a 值
	//alert(a);
	if (str=="OK")
	
	{
	
	window.location.reload()
	
	}
	else
	{
	
	$('#msg').html(str);
	for (var i=1; i<21 ;i++)
	{
    	var newname= $("#Dcode" + i).val();
	
	    var exist=$.inArray(newname,a);

	    if(exist>=0)
	 {
           // alert('重复'+name);  
			 $("#Dcode" + i).select();
			 $("#Dcode" + i).css("background-color","red");
            arrdd=  $.grep(arrdd, function(value, index)
			  {return value!=newname && value!='SANEND' ;});
	
	  }	
	
	
	}
	
	}

  
  
  
     },"html"
	 );//这里返回的类型有：json,html,xml,text
    }
	
		
	
	
	 }
	

		
		
		
		
		
		
		
		
}

function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}

function left(mainStr,lngLen) {
    if (lngLen>0) {return mainStr.substring(0,lngLen)}
    else{return null}
}

function right(mainStr,lngLen) {
    if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) {
    return mainStr.substring(mainStr.length-lngLen,mainStr.length)}
    else{return null}
}


function mid(mainStr,starnum,endnum){
    if (mainStr.length>=0){
        return mainStr.substr(starnum,endnum)
    }else{return null}
} 


function trim(str){ //删除左右两端的空格 
return str.replace(/(^\s*)|(\s*$)/g, ""); 
} 



function ltrim(str){ //删除左边的空格 
return str.replace(/(^\s*)/g,""); 
} 





function rtrim(str){ //删除右边的空格 
return str.replace(/(\s*$)/g,"");  
}







</script>


<script language="javascript"  type="text/javascript">

	function SaveData()
	{
	
	     if(confirm("确定扫描完成借出吗？"))
   { 
			document.form1.action="SendProd_Save.asp?PTC=OK";
			document.form1.submit();
	}
	
	
	
	}
	
	
	 
</script>
<body onload= "javascript:document.all.Dcode1.focus(); "  bgcolor="#339966">
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">


<tr bgcolor="#339966"> 
<td align="center" background="../images/admin_bg_1.gif" height="25">&nbsp;</td>
</tr>
<tr bgcolor="#E8F1FF"> 

<td> 
                                
</td>

</tr>
</table>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr bgcolor="#E8F1FF"> 
<form name="myform" method="post" action="KEB_JSS_save_line5.asp?action=add" OnSubmit="return checkkk()" >
<td> 
                                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#339966">
                                  <tr bgcolor="#E8F1FF" > 
                                    <td height="17" colspan="4" align="center" bgcolor="#339966" ><span class="STYLE3">请扫描2D码</span></td>
                                  </tr>
                                  <tr bgcolor="#E8F1FF" >
                                    <td height="17" colspan="4" align="center" bgcolor="#339966" id="msg">&nbsp;</td>
                                  </tr>
                                  <tr bgcolor="#E8F1FF" >
                                    <td height="35" align="center" bgcolor="#339966" ></td>
                                    <td width="38%" height="0" align="right" bgcolor="#339966" >单号：</td>
                                    <td width="47%" height="0" align="left" bgcolor="#339966" ><div id="SNNO"><%=request("SNNO")%></div></td>
                                    <td height="35" align="center" bgcolor="#339966" >&nbsp;</td>
                                  </tr>
								  
								  

                                  <tr bgcolor="#339966">
								  <td width="10%" align="right">2D Code ：</td>
<td colspan="2" align="center" bgcolor="#339966"><table width="792" height="240" border="0" cellpadding="1" cellspacing="1">
  <tr>
    <td width="149" height="33"><input name="Dcode1" type="text" id="Dcode1" onKeyDown="selectCheck(1)" size="22" /></td>
    <td width="152"><input name="Dcode2" type="text" id="Dcode2" onKeyDown="selectCheck(2)" size="22" /></td>
    <td width="154"><input name="Dcode3" type="text" id="Dcode3" onKeyDown="selectCheck(3)" size="22" /></td>
    <td width="155"><input name="Dcode4" type="text" id="Dcode4" onKeyDown="selectCheck(4)" size="22" /></td>
    <td width="148"><input name="Dcode5" type="text" id="Dcode5" onKeyDown="selectCheck(5)" size="22" /></td>
  </tr>
  <tr>
    <td height="69"><input name="Dcode6" type="text" id="Dcode6" onKeyDown="selectCheck(6)" size="22" /></td>
    <td><input name="Dcode7" type="text" id="Dcode7" onKeyDown="selectCheck(7)" size="22" /></td>
    <td><input name="Dcode8" type="text" id="Dcode8" onKeyDown="selectCheck(8)" size="22" /></td>
    <td><input name="Dcode9" type="text" id="Dcode9" onKeyDown="selectCheck(9)" size="22" /></td>
    <td><input name="Dcode10" type="text" id="Dcode10" onKeyDown="selectCheck(10)" size="22" /></td>
  </tr>

</table>       </td>
 
 
									
<td width="5%"></td>



<%
'rsjs.Close
'set rsjs=nothing
'end if%>
                                  </tr>                                         
                        

<tr bgcolor="#E8F1FF">
<td bgcolor="#339966"></td>
<td height="30" colspan="3" bgcolor="#339966">&nbsp;</td>
</tr>
</table></td>
</form>
</tr>
</table>


<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 

 <%  SNNO=request("SNNO")
 set rsJS=server.CreateObject("adodb.recordset")
			
sql= "select * from  PTC_SN where SNNO='"&SNNO&"'"

 rsJS.Open sql,conn,1,3

		       
	                       
	
	
	%>
<td height="139" colspan="5" align="center" bgcolor="#339966" class="STYLE5"><span class="STYLE5">2D扫描计数器：</span><span class="STYLE11">
    <%



LendNO=rsjs("LendNO")
response.Write(LendNO)

rsjs.Close
set rsjs=nothing
conn.close
set conn=nothing


%>
</span></td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">

<tr align="center"> 

<td height="100"> 
				
				
  		
       			



<form id="form1" method="post" name="form1" >  			
<table width="500" height="7" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td height=7 align="center"><input name="btnclose" type="button"  id="btnclose"  value="Close 关闭" onClick="javascript:window.close();">
  <input name="btnclose2" type="button"  id="btnclose2"  value="重新扫描 " onClick="javascript:window.location.reload();">
 <% 
 
	 
	 if 	 LendNO <>0 then%>
  <input name="btnclose" type="button"  id="btnclose"  value="扫描完成确认借出" onClick="SaveData();"> 
  <%
  end if%>
  <input type="hidden" id="SNNO" name="SNNO" value="<%=request("SNNO")%>">  </td>
</tr>
</table>
</form>
</td>

</tr>
</table>
</body>
</html>

