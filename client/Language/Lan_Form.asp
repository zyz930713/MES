<script language="javascript">
var strTitle="Create New NCMR Ticket|�½��쳣����"
var arrTitle=strTitle.split("|")
var strFactory ="Factory|רҵ��" 
var arrFactory=strFactory.split("|")
var strCreator="User Code|�����˹���" 
var arrCreator=strCreator.split("|")
var strLine="Family|����" 
var arrLine=strLine.split("|")
var strKeyWord="Key Word|�ؼ���" 
var arrKeyWord=strKeyWord.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strModel="Model|�ͺ�" 
var arrModel=strModel.split("|")
var strStation="Station|��վ" 
var arrStation=strStation.split("|")
var strProlem="Problem Description|��������" 
var arrProlem=strProlem.split("|")
var strUpload="Upload File|�ϴ��ļ�" 
var arrUpload=strUpload.split("|")
var strActionPerson="NCMR Owner|NCMR ������" 
var arrActionPerson=strActionPerson.split("|")
var strBackUp="Backup|��ѡ" 
var arrBackUp=strBackUp.split("|")
var strDueDate="Happen Date|��������" 
var arrDueDate=strDueDate.split("|")
var strDotice="Notice Persons|֪ͨ�����" 
var arrDotice=strDotice.split("|")
var strSimilar ="Similar Recodes|���������¼" 
var arrSimilar=strSimilar.split("|")
var strGreen="Green:Finished;|��ɫ����ɣ�" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|��ɫ�������У�" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|��ɫ��δ���У�" 
var arrBlue=strBlue.split("|")
var strNum="NO.|����" 
var arrNum=strNum.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strNumber1 ="NCMR NO.|�쳣�����" 
var arrNumber1=strNumber1.split("|")
var strKeyWord1="Key Word|�ؼ���" 
var arrKeyWord1=strKeyWord1.split("|")
var strJobNumber1="Job Number|������" 
var arrJobNumber1=strJobNumber1.split("|")
var strModel1="Model|�ͺ�" 
var arrModel1=strModel1.split("|")
var strCreateTime1="Create Time|����ʱ��" 
var arrCreateTime1=strCreateTime1.split("|")
var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")

function language()
{
try{inner_Title.innerText=arrTitle[<%=session("language")%>]}catch(e){} 
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){} 
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_KeyWord.innerText=arrKeyWord[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_JobNumber1.innerText=arrJobNumber1[<%=session("language")%>]}catch(e){}
try{inner_Model.innerText=arrModel[<%=session("language")%>]}catch(e){}
try{inner_Model1.innerText=arrModel1[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Prolem.innerText=arrProlem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_BackUp.innerText=arrBackUp[<%=session("language")%>]}catch(e){} 
try{inner_ActionPerson.innerText=arrActionPerson[<%=session("language")%>]}catch(e){}
try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_Dotice.innerText=arrDotice[<%=session("language")%>]}catch(e){}
try{inner_Similar.innerText=arrSimilar[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Number1.innerText=arrNumber1[<%=session("language")%>]}catch(e){}
try{inner_KeyWord1.innerText=arrKeyWord1[<%=session("language")%>]}catch(e){}
try{inner_CreateTime1.innerText=arrCreateTime1[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}

}
</script>