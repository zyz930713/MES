<%
	response.Expires=0
	response.CacheControl="no-cache"
	
	UserCode = session("UserCode")
	if UserCode = "" then
		HtmUrl = "http://"
		HtmUrl = HtmUrl & Request.ServerVariables("Server_Name") & ":9810"
		HtmUrl = HtmUrl & Request.ServerVariables("URL")
		if request.ServerVariables("QUERY_STRING") <> "" then 
			HtmUrl = HtmUrl &"?"& Request.ServerVariables("QUERY_STRING")
		end if
		session("HtmUrl") = HtmUrl
		response.write "<script>alert('连接超时,或者您还没有登录,请先登录');location.href='../../Functions/user_login.asp'</script>"
		response.end()
	end if
	'response.write HtmUrl
%>