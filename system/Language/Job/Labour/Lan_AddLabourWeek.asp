<script language="javascript">
var strBrowse="Add a New Labour Entry|����Ͷ���ת��" 
var arrBrowse=strBrowse.split("|")
var strHostLine="Host Line|����" 
var arrHostLine=strHostLine.split("|")
var strHostLineSelect="--Select Line--|--ѡ���߱�--" 
var arrHostLineSelect=strHostLineSelect.split("|")
var strYear="Year|���" 
var arrYear=strYear.split("|")
var strWeek="Week|�ܴ�" 
var arrWeek=strWeek.split("|")
var strFrom="From|��" 
var arrFrom=strFrom.split("|")
var strTo="To|��" 
var arrTo=strTo.split("|")
var strSave="Save|����" 
var arrSave=strSave.split("|")
var strReset="Reset|����" 
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
