<%
function getFactoryID(classid)
	select case classid
	case "C10-WIP4"
	getFactoryID="FA00000001"'Deltek
	case "C10-WIP1"
	getFactoryID="FA00000002"'KE
	case "C11-WIP1"
	getFactoryID="FA00000002"'KE
	case "C10-WIP3"
	getFactoryID="FA00000003"'VAM
	case else
	getFactoryID=classid
	end select
end function
%>