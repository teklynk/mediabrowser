
<div id="footer"></div>
</div><!--close class container-->
</body>
</html>

<%
'clear session variables on log off
	if request.querystring("logoff") <> "" then
		'delete session variable
		Session("logon")=""
		Session.Contents.RemoveAll()
		Session.Abandon()
		response.redirect("default.asp")
	end if
	
	if request.querystring("debug")="true" then
		response.write Session("userlevel")
	end if
%>