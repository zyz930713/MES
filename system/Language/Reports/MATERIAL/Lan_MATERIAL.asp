<script language="javascript">
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strADID="AD ID|ID 号码"
var arrADID=strADID.split("|") 
var strProductName="Product Name|产品名称"
var arrProductName=strProductName.split("|") 
var strLineName="Line Name|线名称"
var arrLineName=strLineName.split("|")
var strBoxName="Box Name|测试屋"
var arrBoxName=strBoxName.split("|")
var strMeasureCoun="Measure Count|测试数量"
var arrMeasureCount=strMeasureCoun.split("|")
var strAMSMeasureDateTime="AMS Measure Time|AMS测试时间"
var arrAMSMeasureDateTime=strAMSMeasureDateTime.split("|")
var strTestDay="Test Day|测试日"
var arrTestDay=strTestDay.split("|")
var strSerialnumber="Serialnumber|二维码"
var arrSerialnumber=strSerialnumber.split("|")
var strYear="Year|年"
var arrYear=strYear.split("|")
var strWeek="Week|周"
var arrWeek=strWeek.split("|")
var strWeekDay="WeekDay|第几天"
var arrWeekDay=strWeekDay.split("|")
var strProdDate="Prod Date|产品喷码时间"
var arrProdDate=strProdDate.split("|")
var strWeekNum="Week Num|喷码日"
var arrWeekNum=strWeekNum.split("|")
var strDeltaTime="Delta Time|等待时间"
var arrDeltaTime=strDeltaTime.split("|")
var strDeltaDay="Delta Day|等待日"
var arrDeltaDay=strDeltaDay.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strADFAIL="ADFAIL|通过/失败" 
var arrADFAI=strADFAIL.split("|")
var strERROR="ERROR|错误信息" 
var arrERROR=strERROR.split("|")
var strCriterionError="Criterion Error|标准错误信息" 
var arrCriterionError=strCriterionError.split("|")

var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")

var strSubJobs="SubJobs|细分工作单" 
var arrSubJobs=strSubJobs.split("|")
var strLabelId="LabelId|标签号" 
var arrLabelId=strLabelId.split("|")

var strPot="POT|磁路" 
var arrPot=strPot.split("|")

var strTop="TOP|上盖" 
var arrTop=strTop.split("|")

var strBottom="Bottom|下盖" 
var arrBottom=strBottom.split("|")

var strFrame="Frame|盆架" 
var arrFrame=strFrame.split("|")


var strRecords="No Records|没有记录" 
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