<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Rework Job Number|ά�޹�����" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchReworkJobType="Rework Job Type|ά������" 
var arrSearchReworkJobType=strSearchReworkJobType.split("|")
var strSearchStartTime="Rework Start Time|ά�޿�ʼʱ��" 
var arrSearchStartTime=strSearchStartTime.split("|")
var strSearchEndTime="Rework End Time|ά�޽���ʱ��" 
var arrSearchEndTime=strSearchEndTime.split("|")

var strSearchOperatorCode="Operator Code|����������" 
var arrSearchOperatorCode=strSearchOperatorCode.split("|")

var strBrowse="Browse Rework Job Records (Default in past 7 days)|���ά�޹�����¼��Ĭ��7�����ڣ�" 
var arrBrowse=strBrowse.split("|")

var strNO="NO|����" 
var arrNO=strNO.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")

var strJobNumber="Rework Job Number|ά�޹�����" 
var arrJobNumber=strJobNumber.split("|")

var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")

var strJobType="Rework Job Type|ά������" 
var arrJobType=strJobType.split("|")

var strReworkQty="Rework Quantity|ά������" 
var arrReworkQty=strReworkQty.split("|")

var strGoodQty="Good Quantity|�ϸ�����" 
var arrGoodQtys=strGoodQty.split("|")

var strRejectQty="Reject Quantity|��������" 
var arrRejectQty=strRejectQty.split("|")

var strStartTime="Start Time|��ʼʱ��" 
var arrStartTime=strStartTime.split("|")

var strEndTime="End Time|����ʱ��" 
var arrEndTime=strEndTime.split("|")

var strOperatorCode="Operator Code|����������" 
var arrOperatorCode=strOperatorCode.split("|")

var strYield="Yield|����" 
var arrYield=strYield.split("|")

var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")

var strbtnUpdate="Update|����" 
var arrbtnUpdate=strbtnUpdate.split("|")

var strUpdateTitle="Update Job Quantity|����ά�޹�������" 
var arrUpdateTitle=strUpdateTitle.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchReworkType.innerText=arrSearchReworkJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchStartTime.innerText=arrSearchStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchEndTime.innerText=arrSearchEndTime[<%=session("language")%>]}catch(e){}
try{inner_SearchOperatorCode.innerText=arrSearchOperatorCode[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Part_Number.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}

try{inner_JobType.innerText=arrJobType[<%=session("language")%>]}catch(e){}

try{inner_ReworkQty.innerText=arrReworkQty[<%=session("language")%>]}catch(e){}
try{inner_GoodQty.innerText=arrGoodQtys[<%=session("language")%>]}catch(e){}
try{inner_RejectQty.innerText=arrRejectQty[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_EndTime.innerText=arrEndTime[<%=session("language")%>]}catch(e){}
try{inner_OperatorCode.innerText=arrOperatorCode[<%=session("language")%>]}catch(e){}
try{inner_Yield.innerText=arrYield[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{document.all.btnUpdate.value=arrbtnUpdate[<%=session("language")%>]}catch(e){}
try{inner_UpdateTitle.innerText=arrUpdateTitle[<%=session("language")%>]}catch(e){}
}
</script>