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
		response.write "<script>alert('���ӳ�ʱ,��������û�е�¼,���ȵ�¼');location.href='../../Functions/user_login.asp'</script>"
		response.end()
	end if
	'response.write HtmUrl
%>