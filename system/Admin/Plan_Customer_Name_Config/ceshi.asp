<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>使用JQuery和ASP打造AutoComplete功能</title>
<script type="text/javascript" src="jquery.js"></script>
<script type="text/javascript">
 function lookup(inputString) {
  if(inputString.length == 0) {
   // Hide the suggestion box.
   $('#suggestions').hide();
  } else {
   $.post("showmember.asp", {queryString: ""+escape(inputString)+""}, function(data){
    if(data.length >0) {
     $('#suggestions').show();
     $('#autoSuggestionsList').html(unescape(data));
    }
   });
  }
 } // lookup
 
 function fill(thisValue) {
  $('#inputString').val(thisValue);
  setTimeout("$('#suggestions').hide();", 200);
 }
</script>

<style type="text/css">
 body {
  font-family: Helvetica;
  font-size: 11px;
  color: #000;
 }
 
 h3 {
  margin: 0px;
  padding: 0px; 
 }

 .suggestionsBox {
  position: relative;
  left: 30px;
  margin: 10px 0px 0px 0px;
  width: 200px;
  background-color: #212427;
  -moz-border-radius: 7px;
  -webkit-border-radius: 7px;
  border: 2px solid #000; 
  color: #fff;
 }
 
 .suggestionList {
  margin: 0px;
  padding: 0px;
 }
 
 .suggestionList li {
  
  margin: 0px 0px 3px 0px;
  padding: 3px;
  cursor: pointer;
 }
 
 .suggestionList li:hover {
  background-color: #659CD8;
 }
</style>
</head>
<body>
    <div> 
   <form>
   <div>
    请输入您的单位名称：
    <br />
    <input type="text" size="30" value="" id="inputString" onkeyup="lookup(this.value);" onblur="fill();" />
   </div>
   
   <div class="suggestionsBox" id="suggestions" style="display: none;">
    
    <div class="suggestionList" id="autoSuggestionsList">
     &nbsp;
    </div>
   </div>
  </form>
 </div>
</body>
</html>

