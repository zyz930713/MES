<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.tableBorder tr form td table tr td font {
	font-size: 14pt;
	font-weight: bold;
}
.tableBorder tr form td table tr td {
	font-size: 10pt;
}
.BT {
	font-family: "������";
	font-size: 16px;
	color: #000000;
	font-weight: bold;
}
.STYLE1 {color: #000000; font-weight: bold; font-family: "������";}
.STYLE3 {font-size: large}
.STYLE4 {font-size: 24px}
.STYLE5 {font-size: x-large}
.STYLE6 {color: #000000}
.STYLE11 {font-size: larger}

.color3 {background-color:#666; color:#FFF;}
-->
</style>
</head>
<script type="text/javascript" src="JS/jquery-1.9.0.min.js"></script>
<script language="javascript" type="text/javascript">

var arrdd=[];

function selectCheck(testnum)
{
	var ttt=eval(testnum+1);
	
	
  	var evt=evt?evt:(window.event?window.event:null);//����IE��FF
	if (evt.keyCode==13){
		
	  var name= 	$.trim($("#Dcode" + testnum).val()).toUpperCase();
       $("#Dcode" + testnum).val(name);
       var SNNO = $("div#SNNO").html();
	   
	         if (name=="SANEND")
			 {
	          }
			  else
			  {
	          if(name.length!=17) 
              {
	
               alert("�Բ���2D��λ�����ԣ�");
	           return false;
              } 
              
	   
	   /*      if (left(name ,3)!="DYD")
             {
	
            alert("�Բ��𣡱�����������ӦΪ:DYD");
	         return false;
             } 
*/

  
       
	  
	  
     
  
  
     
  	  
     
	  
	  
  	   
	   }
	   
	   
	   
	   
	   
		
		
		
        var exist=$.inArray(name,arrdd);

	  if(exist>=0)
	  {
            alert('�ظ�'+name);  
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
  $.post("Barcode_Return_Save.asp",{Dcode:QueryStringTest,SNNO:SNNO},
  function(data){
    
 // alert(data);
   var reg=/,$/gi;                // ��ʽ�����һ������
   str=data.replace(reg,"");
  var a=str.split(",");   
  
  for (var K=0;K<a.length;K++)
  
  {
  var strb= a[K];
 var jsqS=isNaN(strb);
  
  
  //alert(jsqS);
  
  
  
  
     if  (left(strb,2)=="OK")

	  
	
	{
	 alert('������'); 
	
	window.location.reload()
	//$('#msg').css({'font-size': '18pt', 'color':'#00FF00','font-weight': 'bold'});
	
	  // $('#msg').append(right(strb,17)+" ");
	
	
	
	
 
	}
	
	
    else if (jsqS==false)
	
	{
	//$('#JSQ').css({'font-size': '18pt', 'color':'#00FF00','font-weight': 'bold'});
	
	   $('#JSQ').html(strb+" ");
	
	
	}		
	
	
	
	
	
	
	
	
	
	else
	{
	
	for (var i=1; i<21 ;i++)
	{
    	var newname= $("#Dcode" + i).val();
	
	    var exist=$.inArray(newname,a);
		

	    if(exist>=0)
	 {
            //alert('�ظ�'+name);  
			 $("#Dcode" + i).select();
			 $("#Dcode" + i).css("background-color","red");
            arrdd=  $.grep(arrdd, function(value, index)  //ȥ���������ظ�����
			  {return value!=newname;});
	        // alert(arrdd);
	  }	
	
	
	}
	$('#errmsg').css({'font-size': '18pt', 'color':'#FF0000','font-weight': 'bold'});
	$('#errmsg').append(strb+" ");
	
	}

  }
  
  
     },"html"
	 );//���ﷵ�ص������У�json,html,xml,text
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


function trim(str){ //ɾ���������˵Ŀո� 
return str.replace(/(^\s*)|(\s*$)/g, ""); 
} 



function ltrim(str){ //ɾ����ߵĿո� 
return str.replace(/(^\s*)/g,""); 
} 





function rtrim(str){ //ɾ���ұߵĿո� 
return str.replace(/(\s*$)/g,"");  
}







</script>

<script language="javascript"  type="text/javascript">

	function RSaveData()
	{
	
	     if(confirm("ȷ��Ҫ���ֹ黹��"))
   { 
			document.form1.action="acceptProd_Save.asp?PTCstate=pending";
			document.form1.submit();
	}
	
	
	
	}
	
	
	
	function rAllSaveData()
	{
	
	     if(confirm("ȷ��Ҫȫ���黹��"))
   { 
			document.form1.action="acceptProd_Save.asp?PTCstate=Return";
			document.form1.submit();
	}
	
	
	
	}
	
	
	
	
	
	 
	
	 
	 
	 
	 
	 
</script>



<body onload= "javascript:document.all.Dcode1.focus(); "  bgcolor="#339966">
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">


<tr bgcolor="#339966"> 
<td height="25" align="center" background="../images/admin_bg_1.gif" bgcolor="#FFFFFF">��ɫ�������ɣ����Խ��ա� ��ɫ�����2D�벻�����</td>
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
                                    <td height="34" colspan="4" align="center" bgcolor="#339966" ><BR><span class="STYLE3">��ɨ��2D��</span></td>
                                  </tr>
								  <tr bgcolor="#E8F1FF" > 
                                    <td colspan="4" align="center" bgcolor="#339966"  id="msg" ></td>
                                  </tr>
								  <tr bgcolor="#E8F1FF" > 
                                    <td  colspan="4" align="center" bgcolor="#339966"  id="errmsg" ></td>
                                  </tr>
                                  <tr bgcolor="#E8F1FF" >
                                    <td height="35" align="center" bgcolor="#339966" ></td>
                                    <td width="38%" height="0" align="right" bgcolor="#339966" 	>���ţ�</td>
                                    <td width="47%" height="0" align="left" bgcolor="#339966" ><div id="SNNO"><%=request("SNNO")%></div></td>
                                    <td height="35" align="center" bgcolor="#339966" >&nbsp;</td>
                                  </tr>
								  
								  

                                  <tr bgcolor="#339966">
								  <td width="10%" align="right">2D Code ��</td>
<td colspan="2" align="center"><table width="792" height="173" border="0" cellpadding="1" cellspacing="1">
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

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr> 

 <%  SNNO=request("SNNO")
 set rsJS=server.CreateObject("adodb.recordset")
			
sql= "select * from  PTC_SN where SNNO='"&SNNO&"'"

 rsJS.Open sql,conn,1,3

		       
	                       
	
	
	%>
<td height="139" colspan="4" align="right" bgcolor="#339966" class="STYLE5"><span class="STYLE5">2Dɨ���������</span></td>
<td width="45%" height="139" align="left" bgcolor="#339966" class="STYLE5"><div id="JSQ"><%=rsjs("RETURNNOJSQ")%></div></td>
</tr>
</table>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">

<tr align="center"> 

<td height="100"> 				
  		     			

<form id="form1" method="post" name="form1" >  

			
<table width="604" height="48" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td height=7 align="center"><input name="btnclose" type="button"  id="btnclose"  value="Close �ر�" onClick="javascript:window.close();">
<input name="btnclose" type="button"  id="btnclose"  value="����ɨ�� " onClick="javascript:window.location.reload();">	  <input name="btnSubmit" type="button" id="btnSubmit"  value="���ֹ黹" onClick="RSaveData();">


<%
	 set rsRet=server.CreateObject("adodb.recordset")
			 sql="select * from PTC_SN where SNNO='"&SNNO&"'"
	
	 rsRet.open sql,conn,1,1
	 if rsRet.eof And rsRet.bof then
    
   	 else
	 
	if cint(rsRet("LendNO"))=cint(rsRet("ReturnNO"))+cint(rsRet("ReturnNOJSQ")) then %>
<input name="btnSubmit2" type="button" id="btnSubmit2"  value="ȫ���黹" onClick="rAllSaveData();">
<%else%>
<input name="btnclose" type="button"  id="btnclose"  value="ȷ��ȫ���黹 " onClick="javascript:window.location.reload();">

<%end if
end if%>
<input type="hidden" id="SNNO" name="SNNO" value="<%=request("SNNO")%>"> <input type="hidden" id="SelectType" name="SelectType" value="<%=request("SelectType")%>"></td>
</tr>
</table>
</form>
</td>

</tr>
</table>

</body>
</html>

