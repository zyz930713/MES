<script language="javascript">
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strADID="AD ID|ID ����"
var arrADID=strADID.split("|") 
var strProductName="Product Name|��Ʒ����"
var arrProductName=strProductName.split("|") 
var strLineName="Line Name|������"
var arrLineName=strLineName.split("|")
var strBoxName="Box Name|������"
var arrBoxName=strBoxName.split("|")
var strMeasureCoun="Measure Count|��������"
var arrMeasureCount=strMeasureCoun.split("|")
var strAMSMeasureDateTime="AMS Measure Time|AMS����ʱ��"
var arrAMSMeasureDateTime=strAMSMeasureDateTime.split("|")
var strTestDay="Test Day|������"
var arrTestDay=strTestDay.split("|")
var strSerialnumber="Serialnumber|��ά��"
var arrSerialnumber=strSerialnumber.split("|")
var strYear="Year|��"
var arrYear=strYear.split("|")
var strWeek="Week|��"
var arrWeek=strWeek.split("|")
var strWeekDay="WeekDay|�ڼ���"
var arrWeekDay=strWeekDay.split("|")
var strProdDate="Prod Date|��Ʒ����ʱ��"
var arrProdDate=strProdDate.split("|")
var strWeekNum="Week Num|������"
var arrWeekNum=strWeekNum.split("|")
var strDeltaTime="Delta Time|�ȴ�ʱ��"
var arrDeltaTime=strDeltaTime.split("|")
var strDeltaDay="Delta Day|�ȴ���"
var arrDeltaDay=strDeltaDay.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strADFAIL="ADFAIL|ͨ��/ʧ��" 
var arrADFAI=strADFAIL.split("|")
var strERROR="ERROR|������Ϣ" 
var arrERROR=strERROR.split("|")
var strCriterionError="Criterion Error|��׼������Ϣ" 
var arrCriterionError=strCriterionError.split("|")

var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")

var strSubJobs="SubJobs|ϸ�ֹ�����" 
var arrSubJobs=strSubJobs.split("|")
var strLabelId="LabelId|��ǩ��" 
var arrLabelId=strLabelId.split("|")

var strPot="POT|��·" 
var arrPot=strPot.split("|")

var strTop="TOP|�ϸ�" 
var arrTop=strTop.split("|")

var strBottom="Bottom|�¸�" 
var arrBottom=strBottom.split("|")

var strFrame="Frame|���" 
var arrFrame=strFrame.split("|")


var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")

function language()
{


try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_ADID.innerText=arrADID[<%=session("language")%>]}catch(e){}
try{inner_ProductName.innerText=arrProductName[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_BoxName.innerText=arrBoxName[<%=session("language")%>]}catch(e){}
try{inner_MeasureCount.innerText=arrMeasureCount[<%=session("language")%>]}catch(e){}
try{inner_AMSMeasureDateTime.innerText=arrAMSMeasureDateTime[<%=session("language")%>]}catch(e){}
try{inner_TestDay.innerText=arrTestDay[<%=session("language")%>]}catch(e){}
try{inner_Serialnumber.innerText=arrSerialnumber[<%=session("language")%>]}catch(e){}
try{inner_Year.innerText=arrYear[<%=session("language")%>]}catch(e){}
try{inner_Week.innerText=arrWeek[<%=session("language")%>]}catch(e){}
try{inner_WeekDay.innerText=arrWeekDay[<%=session("language")%>]}catch(e){}
try{inner_ProdDate.innerText=arrProdDate[<%=session("language")%>]}catch(e){}
try{inner_WeekNum.innerText=arrWeekNum[<%=session("language")%>]}catch(e){}
try{inner_DeltaTime.innerText=arrDeltaTime[<%=session("language")%>]}catch(e){}
try{inner_DeltaDay.innerText=arrDeltaDay[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_ADFAI.innerText=arrADFAI[<%=session("language")%>]}catch(e){}
try{inner_ERROR.innerText=arrERROR[<%=session("language")%>]}catch(e){}
try{inner_CriterionError.innerText=arrCriterionError[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_SubJobs.innerText=arrSubJobs[<%=session("language")%>]}catch(e){}
try{inner_LabelId.innerText=arrLabelId[<%=session("language")%>]}catch(e){}
try{inner_Pot.innerText=arrPot[<%=session("language")%>]}catch(e){}
try{inner_Top.innerText=arrTop[<%=session("language")%>]}catch(e){}
try{inner_Bottom.innerText=arrBottom[<%=session("language")%>]}catch(e){}
try{inner_Frame.innerText=arrFrame[<%=session("language")%>]}catch(e){}
try{inner_NoRecords.innerText=arrRecords[<%=session("language")%>]}catch(e){}

}
</script>