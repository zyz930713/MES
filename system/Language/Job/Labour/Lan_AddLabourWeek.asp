<script language="javascript">
var strBrowse="Add a New Labour Entry|浏览劳动力转移" 
var arrBrowse=strBrowse.split("|")
var strHostLine="Host Line|本线" 
var arrHostLine=strHostLine.split("|")
var strHostLineSelect="--Select Line--|--选择线别--" 
var arrHostLineSelect=strHostLineSelect.split("|")
var strYear="Year|年度" 
var arrYear=strYear.split("|")
var strWeek="Week|周次" 
var arrWeek=strWeek.split("|")
var strFrom="From|自" 
var arrFrom=strFrom.split("|")
var strTo="To|至" 
var arrTo=strTo.split("|")
var strSave="Save|保存" 
var arrSave=strSave.split("|")
var strReset="Reset|重置" 
var arrReset=strReset.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_HostLine.innerText=arrHostLine[<%=session("language")%>]}catch(e){}
try{document.all.line.options[0].text=arrHostLineSelect[<%=session("language")%>]}catch(e){}
try{inner_Year.innerText=arrYear[<%=session("language")%>]}catch(e){}
try{inner_Week.innerText=arrWeek[<%=session("language")%>]}catch(e){}
try{inner_From.innerText=arrFrom[<%=session("language")%>]}catch(e){}
try{inner_To.innerText=arrTo[<%=session("language")%>]}catch(e){}
try{document.all.Save.value=arrSave[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>
