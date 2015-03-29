<!--#include file="db/dbcon.asp"-->
<!--#include file="includes/header.asp"-->
<!--
Filemanager
Created by: Ryan Jones
Date: Sept. 21 2013
Version: 1.0
www.teklynk.com
-->
<style type="text/css">
  body {
	padding-top: 40px;
	padding-bottom: 40px;
	/*background-color: #f5f5f5;*/
	background: url(./img/cloudsbg.png) no-repeat center center /cover; 
	-webkit-background-size: cover;
	-moz-background-size: cover;
	-o-background-size: cover;
	background-size: cover;
  }

  .form-signin {
	max-width: 300px;
	padding: 19px 29px 29px;
	margin: 0 auto 20px;
	background-color: #fff;
	border: 1px solid #e5e5e5;
	-webkit-border-radius: 5px;
	   -moz-border-radius: 5px;
			border-radius: 5px;
	-webkit-box-shadow: 0 1px 2px rgba(0,0,0,.05);
	   -moz-box-shadow: 0 1px 2px rgba(0,0,0,.05);
			box-shadow: 0 1px 2px rgba(0,0,0,.05);
  }
  .form-signin .form-signin-heading,
  .form-signin .checkbox {
	margin-bottom: 10px;
  }
  .form-signin input[type="text"],
  .form-signin input[type="password"] {
	font-size: 16px;
	height: auto;
	margin-bottom: 15px;
	padding: 7px 9px;
  }

</style>
<%
'logon stuff
response.buffer=true
Session.Timeout=120

Sub LoginForm
	response.write "<form name=""logonForm"" id=""logonForm"" class=""form-signin"" action=""default.asp"" method=""post"">"&vbCrLf
	response.write "<h2 class=""form-signin-heading"">Please sign in</h2>"&vbCrLf
	if Session("logon")="false" then
		response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""default.asp?dismiss=1"">×</a>Incorrect Username or Password!</div>"&vbCrLf
		Session("logon")=""
	end if

    response.write "<input type=""text"" name=""username"" class=""input-block-level"" placeholder=""Username"" />"&vbCrLf
    response.write "<input type=""password"" name=""password"" class=""input-block-level"" placeholder=""Password"" />"&vbCrLf
	response.write "<input type=""hidden"" name=""send"" value=""1"" />"&vbCrLf
    response.write "<button type=""submit"" name=""submit"" class=""btn btn-large btn-primary"">Sign in</button>"&vbCrLf
	response.write "<div><p></p></div>"&vbCrLf

    response.write "</form>"&vbCrLf
End Sub

Sub LogonInfo
	Session("logon")="false"
	
	set dbconn = server.createobject("ADODB.Connection")
	dbconn.open DBName
	SQL="SELECT username, password, userlevel, id FROM users WHERE '"&request.form("username")&"' = username"
	set rs=dbconn.execute(SQL)
		do while not rs.eof
			username = rs("username")
			password = rs("password")
			userlevel = rs("userlevel")
			userid = rs("id")
			rs.movenext
		loop
	rs.close
	
	if (Strcomp(Request.Form("username"),username,1)=0 AND Request.Form("password") = password) then
		'create session variable
		Session("logon")="true"
		Session("username")=username
		Session("userlevel")=userlevel
		Session("userid")=userid
	end if

End Sub

'''''''''''''''''''''Calling the Subs'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'connect to DB if logon=1 and create session cookie
	if request.form("send")=1 then
		LogonInfo
	end if	

'show the rest of page or show logon box if not logged in
	if Session("logon")="true" then
		response.redirect("filemgr.asp?mobileFiles=true")
	elseif request.querystring("dismiss")=1 then
		'set session cookie to nothing to clear data
		Session("logon")=""
		response.redirect("default.asp")
	else
		LoginForm
	end if
'end logon stuff
%>
<!--#include file="includes/footer.asp"-->