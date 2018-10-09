<script language="javascript">
var strSearch="Search Scrap Change Records|搜索报废修改记录" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|线别" 
var arrSearchLine=strSearchLine.split("|")
var strSearchChangeTime="Scrap Time|报废时间" 
var arrSearchChangeTime=strSearchChangeTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strSearchCode="Code|修改人" 
var arrSearchCode=strSearchCode.split("|")

var strBrowse="Browse Scrap Change Records|浏览报废修改记录" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strCode="Code|报废人" 
var arrCode=strCode.split("|")
var strOldQuantity="Old Quantity|旧的数量" 
var arrOldQuantity=strOldQuantity.split("|")
var strNewQuantity="New Quantity|新的数量" 
var arrNewQuantity=strNewQuantity.split("|")
var strChangeTime="Change Time|修改时间" 
var arrChangeTime=strChangeTime.split("|")
var strChangeReason="Change Reason|修改理由" 
var arrChangeReason=strChangeReason.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchChangeTime.innerText=arrSearchChangeTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_OldQuantity.innerText=arrOldQuantity[<%=session("language")%>]}catch(e){}
try{inner_NewQuantity.innerText=arrNewQuantity[<%=session("language")%>]}catch(e){}
try{inner_ChangeTime.innerText=arrChangeTime[<%=session("language")%>]}catch(e){}
try{inner_ChangeReason.innerText=arrChangeReason[<%=session("language")%>]}catch(e){}
}
</script>