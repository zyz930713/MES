<script language="javascript">
var strJobisopened="Job is opened|��������"
var arrJobisopened=strJobisopened.split("|")
var strJobispaused="Job is paused|������ͣ"
var arrJobispaused=strJobispaused.split("|")
var strJobisclosed="Job is closed|�����ر�"
var arrJobisclosed=strJobisclosed.split("|")
var strJobisshiftouted="Job is shift-out|����ͣ��"
var arrJobisshiftouted=strJobisshiftouted.split("|")
var strJobislocked="Job is locked|��������"
var arrJobislocked=strJobislocked.split("|")
var strJobisaborted="Job is aborted|�����ϳ�"
var arrJobisaborted=strJobisaborted.split("|")
var strStartjob="Start job|��������"
var arrStartjob=strStartjob.split("|")
var strPausejob="Pause job|��ͣ����"
var arrPausejob=strPausejob.split("|")
var strAbortjob="Abort job|�ϳ�����"
var arrAbortjob=strAbortjob.split("|")
var strRepeatjob="Repeat job|��������"
var arrRepeatjob=strRepeatjob.split("|")
var strHistory="History|���Ƽ�¼"
var arrHistory=strHistory.split("|")
var strRepeated="Repeated|������¼"
var arrRepeated=strRepeated.split("|")
var strGreen="Green|��ɫ"
var arrGreen=strGreen.split("|")
var strFinishedstation="Finished station|��վ���"
var arrFinishedstation=strFinishedstation.split("|")
var strRed="Green|��ɫ"
var arrRed=strRed.split("|")
var strCurrentstation="Current station|��ǰ��վ"
var arrCurrentstation=strCurrentstation.split("|")
var strBlue="Green|��ɫ"
var arrBlue=strBlue.split("|")
var strWaitingstation="Waiting station|δ����վ"
var arrWaitingstation=strWaitingstation.split("|")
var strItalic="Italic|б��"
var arrItalic=strItalic.split("|")
var strOptionalstations="Optional stations|��ѡվ"
var arrOptionalstations=strOptionalstations.split("|")
var strConjuctive="[]|[]"
var arrConjuctive=strConjuctive.split("|")
var strConjuctivestations="Conjuctive stations|���վ"
var arrConjuctivestations=strConjuctivestations.split("|")
var strPractised="Practised|���վ"
var arrPractised=strPractised.split("|")
function language_jobnote()
{
try{inner_Jobisopened.innerText=arrJobisopened[<%=session("language")%>]}catch(e){} 
try{inner_Jobispaused.innerText=arrJobispaused[<%=session("language")%>]}catch(e){} 
try{inner_Jobisclosed.innerText=arrJobisclosed[<%=session("language")%>]}catch(e){} 
try{inner_Jobisshiftouted.innerText=arrJobisshiftouted[<%=session("language")%>]}catch(e){} 
try{inner_Jobislocked.innerText=arrJobislocked[<%=session("language")%>]}catch(e){} 
try{inner_Jobisaborted.innerText=arrJobisaborted[<%=session("language")%>]}catch(e){}
try{inner_Startjob.innerText=arrStartjob[<%=session("language")%>]}catch(e){}
try{inner_Pausejob.innerText=arrPausejob[<%=session("language")%>]}catch(e){}
try{inner_Abortjob.innerText=arrAbortjob[<%=session("language")%>]}catch(e){}
try{inner_Repeatjob.innerText=arrRepeatjob[<%=session("language")%>]}catch(e){}
try{inner_History.innerText=arrHistory[<%=session("language")%>]}catch(e){}
try{inner_Repeated.innerText=arrRepeated[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Finishedstation.innerText=arrFinishedstation[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Currentstation.innerText=arrCurrentstation[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Waitingstation.innerText=arrWaitingstation[<%=session("language")%>]}catch(e){}
try{inner_Italic.innerText=arrItalic[<%=session("language")%>]}catch(e){}
try{inner_Optionalstations.innerText=arrOptionalstations[<%=session("language")%>]}catch(e){}
try{inner_Conjuctivestations.innerText=arrConjuctivestations[<%=session("language")%>]}catch(e){}
try{inner_Finishedstation2.innerText=arrFinishedstation[<%=session("language")%>]}catch(e){}
try{inner_Currentstation2.innerText=arrCurrentstation[<%=session("language")%>]}catch(e){}
try{inner_Practised.innerText=arrPractised[<%=session("language")%>]}catch(e){}
}
</script>
