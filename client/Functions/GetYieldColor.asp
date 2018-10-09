<%
function GetYieldColor(actual,target)
	if csng(actual)>=csng(target)/100 and csng(actual)<>0 then
	GetYieldColor="t-b-Green"
	elseif (csng(target)/100-csng(actual))<=0.005 and csng(actual)<csng(target)/100 and csng(actual)<>0 then
	GetYieldColor="t-b-Yellow"
	elseif (csng(target)/100-csng(actual))>0.005 and csng(actual)<>0 then
	GetYieldColor="t-b-Red"
	end if
end function

function GetFirstYieldColor(output,actual,target)
	if csng(actual)>=csng(target)/100 and csng(actual)<>0 then
	GetFirstYieldColor="t-b-Green"
	elseif (csng(target)/100-csng(actual))>0.05 and csng(actual)<>0 then
	GetFirstYieldColor="t-b-Red"
	end if
	if csng(output)>10000 and (csng(target)/100-csng(actual))>0.03 then
	GetFirstYieldColor="t-b-Red"
	end if
end function

function GetNormalYieldColor(actual,target)
	if csng(actual)>=csng(target) and csng(actual)<>0 then
	GetNormalYieldColor="t-b-Green"
	elseif (csng(target)-csng(actual))<=0.005 and csng(actual)<csng(target) and csng(actual)<>0 then
	GetNormalYieldColor="t-b-Yellow"
	elseif (csng(target)-csng(actual))>0.005 and csng(actual)<>0 then
	GetNormalYieldColor="t-b-Red"
	end if
end function

function GetYieldExcelColor(actual,target)
	if csng(actual)>=csng(target)/100 and csng(actual)<>0 then
	GetYieldExcelColor=xlColorGreen
	elseif (csng(target)/100-csng(actual))<=0.005 and csng(actual)<csng(target)/100 and csng(actual)<>0 then
	GetYieldExcelColor=xlColorYellow
	elseif (csng(target)/100-csng(actual))>0.005 and csng(actual)<>0 then
	GetYieldExcelColor=xlColorRed
	end if
end function

function GetFirstYieldExcelColor(output,actual,target)
	if csng(actual)>=csng(target)/100 and csng(actual)<>0 then
	GetFirstYieldExcelColor=xlColorGreen
	elseif (csng(target)/100-csng(actual))>0.05 and csng(actual)<>0 then
	GetFirstYieldExcelColor=xlColorRed
	end if
	if csng(output)>10000 and (csng(target)/100-csng(actual))>0.03 then
	GetFirstYieldExcelColor=xlColorRed
	end if
end function

function GetNormalYieldExcelColor(actual,target)
	if csng(actual)>=csng(target) and csng(actual)<>0 then
	GetNormalYieldExcelColor=xlColorGreen
	elseif (csng(target)-csng(actual))<=0.005 and csng(actual)<csng(target) and csng(actual)<>0 then
	GetNormalYieldExcelColor=xlColorYellow
	elseif (csng(target)-csng(actual))>0.005 and csng(actual)<>0 then
	GetNormalYieldExcelColor=xlColorRed
	end if
end function

function GetScrapColor(actual,target)
	if csng(actual)<=csng(target) and csng(actual)<>0 then
	GetScrapColor="t-b-Green"
	elseif (csng(target)-csng(actual))<=0.005 and csng(actual)>csng(target) and csng(actual)<>0 then
	GetScrapColor="t-b-Yellow"
	elseif (csng(target)-csng(actual))>0.005 and csng(actual)<>0 then
	GetScrapColor="t-b-Red"
	end if
end function
%>
