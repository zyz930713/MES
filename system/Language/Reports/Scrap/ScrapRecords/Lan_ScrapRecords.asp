<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchScrapTime="Scrap Time|����ʱ��" 
var arrSearchScrapTime=strSearchScrapTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strSearchCode="Code|������" 
var arrSearchCode=strSearchCode.split("|")

var strBrowse="Browse Scrap Records|������ϼ�¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strCheck="Check|����" 
var arrCheck=strCheck.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strCode="Code|������" 
var arrCode=strCode.split("|")
var strInputQuantity="Input Quantity|��������" 
var arrInputQuantity=strInputQuantity.split("|")
var strScrapQuantity="Scrap Quantity|��������" 
var arrScrapQuantity=strScrapQuantity.split("|")
var strOntimeYield="OntimeYield|��ʱ����" 
var arrOntimeYield=strOntimeYield.split("|")
var strScrapTime="Scrap Time|����ʱ��" 
var arrScrapTime=strScrapTime.split("|")
var strScrapType="Scrap Type|��������" 
var arrScrapType=strScrapType.split("|")
var strReason="Reason|����" 
var arrReason=strReason.split("|")
var strFormID="Form ID|�����" 
var arrFormID=strFormID.split("|")
var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")
var strNote="Note|ע��" 
var arrNote=strNote.split("|")
var arrERPJobStatus=["ERP Job Status","ERP����״̬"]
var arrScrapAccount=["Scrap Account","�����˺�"]
var arrScrapReason=["Scrap Reason","����ԭ��"]
var arrScrapReference=["Scrap Reference","���ϲ���"]
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchScrapTime.innerText=arrSearchScrapTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Check.innerText=arrCheck[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_ScrapQuantity.innerText=arrScrapQuantity[<%=session("language")%>]}catch(e){}
try{inner_OntimeYield.innerText=arrOntimeYield[<%=session("language")%>]}catch(e){}
try{inner_ScrapTime.innerText=arrScrapTime[<%=session("language")%>]}catch(e){}
try{inner_ScrapType.innerText=arrScrapType[<%=session("language")%>]}catch(e){}
try{inner_Reason.innerText=arrReason[<%=session("language")%>]}catch(e){}
try{inner_FormID.innerText=arrFormID[<%=session("language")%>]}catch(e){}
try{inner_NoRecords.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_Note.innerText=arrNote[<%=session("language")%>]}catch(e){}
try{td_ERPJobStatus.innerText=arrERPJobStatus[<%=session("language")%>]}catch(e){}
try{td_ScrapAccount.innerText=arrScrapAccount[<%=session("language")%>]}catch(e){}
try{td_ScrapReason.innerText=arrScrapReason[<%=session("language")%>]}catch(e){}
try{td_ScrapReference.innerText=arrScrapReference[<%=session("language")%>]}catch(e){}
}
</script>