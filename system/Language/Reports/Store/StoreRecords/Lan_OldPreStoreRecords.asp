<script language="javascript">
var strSearch="Search Pre Store Records|����Ԥ����¼����ϵͳ��" 
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchLine="Line|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchSection="Section|����" 
var arrSearchSection=strSearchSection.split("|")
var strSearchChangeTime="Change Time|�޸�ʱ��" 
var arrSearchChangeTime=strSearchChangeTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strSearchCode="Code|�޸���" 
var arrSearchCode=strSearchCode.split("|")
var strSearchStoreType="Store Type|�������" 
var arrSearchStoreType=strSearchStoreType.split("|")
var strSearchStoreTypeSelect="-- Select Type --|-- ѡ������ --" 
var arrSearchStoreTypeSelect=strSearchStoreTypeSelect.split("|")
var strSearchSectionSelect="-- Select Section --|-- ѡ������ --" 
var arrSearchSectionSelect=strSearchSectionSelect.split("|")
var strSearchStoreNormal="Normal|��������" 
var arrSearchStoreNormal=strSearchStoreNormal.split("|")
var strSearchStoreRework="Rework|���޹���" 
var arrSearchStoreRework=strSearchStoreRework.split("|")
var strSearchCheckStatus="Check Status|����״̬" 
var arrSearchCheckStatus=strSearchCheckStatus.split("|")
var strSearchCheckStatusSelect="-- Select Status --|-- ѡ��״̬ --" 
var arrSearchCheckStatusSelect=strSearchCheckStatusSelect.split("|")
var strSearchChecked="Checked|�Ѹ���" 
var arrSearchChecked=strSearchChecked.split("|")
var strSearchUnchecked="Unchecked|δ����" 
var arrSearchUnchecked=strSearchUnchecked.split("|")

var strBrowse="Browse Pre Store Change Records|���Ԥ����޸ļ�¼����ϵͳ��" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strCode="Code|�޸���" 
var arrCode=strCode.split("|")
var strInputQuantity="Input Quantity|��������" 
var arrInputQuantity=strInputQuantity.split("|")
var strInspectQuantity="Inspect Quantity|��������" 
var arrInspectQuantity=strInspectQuantity.split("|")
var strStoreQuantity="Store Quantity|�������" 
var arrStoreQuantity=strStoreQuantity.split("|")
var strOntimeYield="OntimeYield|��ʱ����" 
var arrOntimeYield=strOntimeYield.split("|")
var strStoreTime="Store Time|���ʱ��" 
var arrStoreTime=strStoreTime.split("|")
var strStoreType="Store Type|�������" 
var arrStoreType=strStoreType.split("|")
var strNote="Note|ע��" 
var arrNote=strNote.split("|")
var strSubJobs="SubJobs|ϸ�ֹ�����" 
var arrSubJobs=strSubJobs.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchSection.innerText=arrSearchSection[<%=session("language")%>]}catch(e){}
try{document.all.section.options[0].text=arrSearchSectionSelect[<%=session("language")%>]}catch(e){}
try{inner_SearchStoreTime.innerText=arrSearchStoreTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchCode.innerText=arrSearchCode[<%=session("language")%>]}catch(e){}
try{inner_SearchStoreType.innerText=arrSearchStoreType[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[0].text=arrSearchStoreTypeSelect[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[1].text=arrSearchStoreNormal[<%=session("language")%>]}catch(e){}
try{document.all.storetype.options[2].text=arrSearchStoreRework[<%=session("language")%>]}catch(e){}
try{inner_SearchCheckStatus.innerText=arrSearchCheckStatus[<%=session("language")%>]}catch(e){}
try{document.all.checkstatus.options[0].text=arrSearchCheckStatusSelect[<%=session("language")%>]}catch(e){}
try{document.all.checkstatus.options[1].text=arrSearchChecked[<%=session("language")%>]}catch(e){}
try{document.all.checkstatus.options[2].text=arrSearchUnchecked[<%=session("language")%>]}catch(e){}

try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Check.innerText=arrCheck[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Code.innerText=arrCode[<%=session("language")%>]}catch(e){}
try{inner_InputQuantity.innerText=arrInputQuantity[<%=session("language")%>]}catch(e){}
try{inner_InspectQuantity.innerText=arrInspectQuantity[<%=session("language")%>]}catch(e){}
try{inner_StoreQuantity.innerText=arrStoreQuantity[<%=session("language")%>]}catch(e){}
try{inner_OntimeYield.innerText=arrOntimeYield[<%=session("language")%>]}catch(e){}
try{inner_StoreTime.innerText=arrStoreTime[<%=session("language")%>]}catch(e){}
try{inner_StoreType.innerText=arrStoreType[<%=session("language")%>]}catch(e){}
try{inner_Note.innerText=arrNote[<%=session("language")%>]}catch(e){}
try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
}
</script>